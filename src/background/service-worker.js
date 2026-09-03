/*
 * Fygaro Catalog Automation
 * Background worker. It is the single owner of the run state.
 *
 * The content script executes steps and the side panel renders them, but every
 * decision about what happens next is made here and written to storage first.
 * A service worker can be shut down at any moment, so nothing is held in memory
 * between messages: each handler reads the run, mutates it, and writes it back.
 */
'use strict';

importScripts('../shared/util.js', '../shared/state.js');

var FYG = self.FYG;
var S = FYG.state;
var U = FYG.util;

/* Mutations are serialised through one promise chain so two messages arriving
 * together cannot both read the same state and clobber each other. */
var queue = Promise.resolve();
function serial(task) {
  var run = queue.then(task, task);
  queue = run.catch(function () {});
  return run;
}

function log(bucket, level, message) {
  bucket.log = U.pushLog(bucket.log, level, message, S.LOG_CAP);
}

function broadcast() {
  chrome.runtime.sendMessage({ type: S.MSG.STATE_CHANGED }).catch(function () {
    // No side panel open. Nothing to tell.
  });
}

function persist(bucket) {
  var jobs = [S.writeRun(bucket.run)];
  if (bucket.rowsDirty) jobs.push(S.writeRows(bucket.rows));
  if (bucket.logDirty) jobs.push(S.writeLog(bucket.log));
  return Promise.all(jobs).then(function () {
    return updateBadge(bucket.run);
  }).then(broadcast);
}

function updateBadge(run) {
  var text = '';
  var color = '#2457D6';
  if (run.status === S.STATUS.RUNNING) {
    text = String(run.stats.done);
  } else if (run.status === S.STATUS.ATTENTION) {
    text = '!';
    color = '#C02626';
  } else if (run.status === S.STATUS.PAUSED) {
    text = '||';
    color = '#B45309';
  } else if (run.status === S.STATUS.DONE) {
    text = 'OK';
    color = '#17803D';
  }
  return Promise.all([
    chrome.action.setBadgeText({ text: text }),
    chrome.action.setBadgeBackgroundColor({ color: color })
  ]).catch(function () {});
}

function notify(title, message) {
  try {
    chrome.notifications.create({
      type: 'basic',
      iconUrl: chrome.runtime.getURL('icons/icon128.png'),
      title: title,
      message: message,
      priority: 2
    }, function () { void chrome.runtime.lastError; });
  } catch (e) {
    // Notifications are a convenience, never a requirement.
  }
}

/** Index of the next row that still needs work, or -1 when the run is finished. */
function nextRowIndex(rows, from) {
  for (var i = Math.max(0, from); i < rows.length; i++) {
    if (rows[i].status === S.ROW.PENDING || rows[i].status === S.ROW.ACTIVE) return i;
  }
  return -1;
}

function currentRow(bucket) {
  var i = bucket.run.cursor;
  return (i >= 0 && i < bucket.rows.length) ? bucket.rows[i] : null;
}

/** Sends the driven tab to a url, tolerating a tab the user already closed. */
function navigate(run, url) {
  if (run.tabId == null) return Promise.resolve();
  return chrome.tabs.update(run.tabId, { url: url }).catch(function () {
    run.tabId = null;
  });
}

/*
 * Zoom is deliberately real browser zoom rather than a CSS transform, because
 * only real zoom changes the layout viewport, which is what Fygaro's responsive
 * breakpoints actually respond to.
 *
 * The default per-origin scope is what is wanted here: per-tab zoom is reset on
 * every navigation, and this run navigates constantly.
 */
function applyZoom(run) {
  if (run.tabId == null) return Promise.resolve({ ok: false, changed: false, reason: 'there is no Fygaro tab' });

  var target = (run.settings.zoomPercent || 100) / 100;
  var CLOSE_ENOUGH = 0.005;

  /*
   * A tab that was only just created is still loading, and zooming it can fail
   * until it settles, so this retries. It also reads the zoom back afterwards
   * rather than assuming setZoom worked, because a silent no-op here would
   * leave the run driving a layout the steps were not written for.
   */
  function attempt(triesLeft) {
    return chrome.tabs.getZoom(run.tabId).then(function (current) {
      // Only the first capture counts, so pausing and resuming cannot record
      // the automation's own zoom as the value to restore later.
      if (run.originalZoom == null) run.originalZoom = current;
      if (Math.abs(current - target) < CLOSE_ENOUGH) {
        return { ok: true, changed: false, zoom: current };
      }
      return chrome.tabs.setZoom(run.tabId, target)
        .then(function () { return chrome.tabs.getZoom(run.tabId); })
        .then(function (after) {
          if (Math.abs(after - target) < CLOSE_ENOUGH) return { ok: true, changed: true, zoom: after };
          throw new Error('the tab reported ' + Math.round(after * 100) + '% afterwards');
        });
    }).catch(function (err) {
      if (triesLeft > 1) return U.sleep(400).then(function () { return attempt(triesLeft - 1); });
      return { ok: false, changed: false, reason: String(err && err.message ? err.message : err) };
    });
  }

  return attempt(5);
}

/**
 * Applies the zoom and always says what happened, including when nothing needed
 * doing. Staying quiet on the no-change case once made a wrong zoom setting look
 * exactly like a broken feature, so the log now names the number every time.
 */
function applyZoomAndReport(bucket) {
  var wanted = bucket.run.settings.zoomPercent;

  return applyZoom(bucket.run).then(function (result) {
    bucket.logDirty = true;

    if (!result.ok) {
      log(bucket, 'warn', 'The page zoom could not be set to ' + wanted + '% (' + result.reason +
        '). The run will carry on, but if a step cannot find a field, zoom the Fygaro page out ' +
        'manually with Ctrl and minus.');
    } else if (result.changed) {
      log(bucket, 'info', 'Page zoom set to ' + Math.round(result.zoom * 100) +
        '% so the side panel cannot change the layout. It is put back when the run ends.');
    } else {
      log(bucket, 'info', 'Page zoom is already ' + Math.round(result.zoom * 100) + '%, left as it is.');
    }
    return result;
  });
}

/** Hands the page back the zoom it had before the run started. */
function restoreZoom(run) {
  var back = run.originalZoom;
  run.originalZoom = null;
  if (run.tabId == null || back == null) return Promise.resolve();
  return chrome.tabs.setZoom(run.tabId, back).catch(function () {});
}

function raiseAttention(bucket, message, url) {
  var row = currentRow(bucket);
  bucket.run.status = S.STATUS.ATTENTION;
  bucket.run.pending = {
    rowIndex: bucket.run.cursor,
    step: bucket.run.step,
    stepLabel: S.STEP_LABEL[bucket.run.step] || bucket.run.step,
    code: row ? row.code : '',
    name: row ? row.name : '',
    sheetRow: row ? row.sheetRow : null,
    url: url || '',
    message: message
  };
  log(bucket, 'error', (row ? 'Row ' + row.sheetRow + ' (' + row.code + '): ' : '') + message);
  bucket.logDirty = true;
  notify('Fygaro automation needs you', U.truncate(message, 180));
}

/** Marks the current row finished and moves the cursor to the next one. */
function completeRow(bucket, status, link) {
  var run = bucket.run;
  var row = currentRow(bucket);
  if (row) {
    row.status = status;
    if (link) row.link = link;
    row.finishedAt = Date.now();
    if (run.rowStartedAt) {
      var took = Date.now() - run.rowStartedAt;
      run.durations.push(took);
      if (run.durations.length > 25) run.durations = run.durations.slice(-25);
    }
    bucket.rowsDirty = true;
  }

  run.stats = S.recount(bucket.rows);
  run.pending = null;
  run.attempt = 0;
  run.waitingSince = null;

  var next = nextRowIndex(bucket.rows, run.cursor + 1);
  if (next === -1) {
    run.cursor = -1;
    run.status = S.STATUS.DONE;
    run.finishedAt = Date.now();
    run.step = S.STEP.NAV_TO_PRODUCTS;
    log(bucket, 'success', 'Run finished. ' + run.stats.done + ' done, ' +
      run.stats.skipped + ' skipped, ' + run.stats.failed + ' failed.');
    bucket.logDirty = true;
    notify('Fygaro automation finished', run.stats.done + ' links created, ' + run.stats.failed + ' failed.');
    return restoreZoom(run);
  }

  run.cursor = next;
  run.step = S.STEP.NAV_TO_PRODUCTS;
  run.rowStartedAt = Date.now();
  bucket.rows[next].status = S.ROW.ACTIVE;
  bucket.rowsDirty = true;
  // Re-asserted once per row so a reset, a new tab or a stray Ctrl and zero
  // cannot quietly put the run back onto the narrow layout.
  return applyZoom(run).then(function () {});
}

/* ---------------------------------------------------------------- handlers */

var handlers = {};

handlers[S.MSG.GET_STATE] = function (msg, bucket) {
  return Promise.resolve({ ok: true });
};

handlers[S.MSG.LOAD_CATALOG] = function (msg, bucket) {
  var rows = (msg.rows || []).map(function (r, i) {
    return {
      i: i,
      sheetRow: r.sheetRow,
      code: r.code,
      name: r.name,
      priceRaw: r.priceRaw,
      priceText: r.priceText,
      link: r.link || '',
      // A row that already carries a link is left alone, which is what makes a
      // restart safe. A row missing a name, code or readable price is marked
      // failed up front rather than breaking the run halfway through.
      status: r.link ? S.ROW.SKIPPED : (r.blocked ? S.ROW.FAILED : S.ROW.PENDING),
      error: r.link ? '' : (r.blocked ? r.blockedReason || 'This row cannot be processed.' : ''),
      productUuid: '',
      finishedAt: null
    };
  });

  bucket.rows = rows;
  bucket.rowsDirty = true;

  var settings = bucket.run.settings;
  var originalZoom = bucket.run.originalZoom;
  var tabId = bucket.run.tabId;
  bucket.run = S.defaultRun();
  bucket.run.settings = settings;
  // Carried over so a zoom this extension applied is still owed back.
  bucket.run.originalZoom = originalZoom;
  bucket.run.tabId = tabId;
  bucket.run.file = msg.file || null;
  bucket.run.stats = S.recount(rows);

  log(bucket, 'info', 'Catalog loaded: ' + rows.length + ' rows, ' +
    bucket.run.stats.toProcess + ' to process, ' + bucket.run.stats.skipped + ' already have a link.');
  bucket.logDirty = true;
  return Promise.resolve({ ok: true });
};

handlers[S.MSG.UPDATE_SETTINGS] = function (msg, bucket) {
  bucket.run.settings = Object.assign({}, bucket.run.settings, msg.settings || {});
  return Promise.resolve({ ok: true });
};

handlers[S.MSG.START] = function (msg, bucket) {
  if (!bucket.rows.length) return Promise.resolve({ ok: false, error: 'Load a catalog file first.' });

  var first = nextRowIndex(bucket.rows, 0);
  if (first === -1) {
    bucket.run.status = S.STATUS.DONE;
    return Promise.resolve({ ok: false, error: 'Every row already has a link. Nothing to do.' });
  }

  return findOrOpenTab().then(function (tabId) {
    var run = bucket.run;
    run.tabId = tabId;
    run.cursor = first;
    run.step = S.STEP.NAV_TO_PRODUCTS;
    run.attempt = 0;
    run.pending = null;
    run.status = S.STATUS.RUNNING;
    run.startedAt = Date.now();
    run.finishedAt = null;
    run.rowStartedAt = Date.now();
    run.waitingSince = null;
    run.durations = [];
    bucket.rows[first].status = S.ROW.ACTIVE;
    bucket.rowsDirty = true;
    run.stats = S.recount(bucket.rows);

    log(bucket, 'info', run.settings.dryRun
      ? 'Dry run started. The product form will be filled for row ' + bucket.rows[first].sheetRow + ' and left unsaved.'
      : 'Run started at row ' + bucket.rows[first].sheetRow + '.');
    bucket.logDirty = true;

    // Zoom first, then navigate, so the very first page already renders at the
    // layout the steps expect.
    return applyZoomAndReport(bucket).then(function () {
      return navigate(run, S.APP_URL);
    });
  }).then(function () {
    return { ok: true };
  });
};

handlers[S.MSG.PAUSE] = function (msg, bucket) {
  if (bucket.run.status === S.STATUS.RUNNING) {
    bucket.run.status = S.STATUS.PAUSED;
    log(bucket, 'info', 'Paused.');
    bucket.logDirty = true;
  }
  return Promise.resolve({ ok: true });
};

handlers[S.MSG.RESUME] = function (msg, bucket) {
  if (bucket.run.status !== S.STATUS.PAUSED && bucket.run.status !== S.STATUS.ATTENTION) {
    return Promise.resolve({ ok: true });
  }
  bucket.run.status = S.STATUS.RUNNING;
  bucket.run.pending = null;
  bucket.run.attempt = 0;
  bucket.run.waitingSince = null;
  log(bucket, 'info', 'Resumed.');
  bucket.logDirty = true;
  return applyZoomAndReport(bucket).then(function () { return { ok: true }; });
};

handlers[S.MSG.RETRY] = function (msg, bucket) {
  bucket.run.status = S.STATUS.RUNNING;
  bucket.run.pending = null;
  bucket.run.attempt = 0;
  bucket.run.waitingSince = null;
  log(bucket, 'info', 'Retrying ' + (S.STEP_LABEL[bucket.run.step] || bucket.run.step) + '.');
  bucket.logDirty = true;
  return Promise.resolve({ ok: true });
};

/* Skipping restarts the row cycle from the dashboard so the next row begins
 * from a known page rather than wherever the failure left the browser. */
handlers[S.MSG.SKIP_ROW] = function (msg, bucket) {
  var row = currentRow(bucket);
  if (row) {
    row.error = bucket.run.pending ? bucket.run.pending.message : 'Skipped by the user.';
    log(bucket, 'warn', 'Row ' + row.sheetRow + ' (' + row.code + ') skipped by the user.');
    bucket.logDirty = true;
  }
  return completeRow(bucket, S.ROW.FAILED, null).then(function () {
    if (bucket.run.status === S.STATUS.DONE) return null;
    bucket.run.status = S.STATUS.RUNNING;
    return navigate(bucket.run, S.APP_URL);
  }).then(function () {
    return { ok: true };
  });
};

handlers[S.MSG.STOP] = function (msg, bucket) {
  var row = currentRow(bucket);
  if (row && row.status === S.ROW.ACTIVE) row.status = S.ROW.PENDING;
  bucket.rowsDirty = true;
  bucket.run.status = S.STATUS.IDLE;
  bucket.run.pending = null;
  bucket.run.stats = S.recount(bucket.rows);
  log(bucket, 'info', 'Stopped. Progress is kept, so Start will carry on from here.');
  bucket.logDirty = true;
  return restoreZoom(bucket.run).then(function () { return { ok: true }; });
};

handlers[S.MSG.RESET] = function (msg, bucket) {
  return restoreZoom(bucket.run).then(function () {
    return S.clear();
  }).then(function () {
    return updateBadge(S.defaultRun());
  }).then(broadcast).then(function () {
    return { ok: true, handled: true };
  });
};

handlers[S.MSG.OPEN_FYGARO] = function (msg, bucket) {
  return findOrOpenTab().then(function (tabId) {
    bucket.run.tabId = tabId;
    return chrome.tabs.update(tabId, { active: true }).catch(function () {});
  }).then(function () {
    return { ok: true };
  });
};

/**
 * The content script asks what it should be doing on the page it is looking at.
 * The answer is derived purely from stored state plus the reported url, so it is
 * safe to ask repeatedly.
 */
handlers[S.MSG.REQUEST_JOB] = function (msg, bucket, sender) {
  var run = bucket.run;
  var tabId = sender && sender.tab ? sender.tab.id : null;
  var row = currentRow(bucket);

  // Every reply carries the on page status widget's data, whatever the answer.
  var hud = {
    status: run.status,
    done: run.stats.done,
    total: run.stats.total,
    sheetRow: row ? row.sheetRow : null,
    code: row ? row.code : '',
    stepLabel: S.STEP_LABEL[run.step] || run.step,
    message: run.pending ? run.pending.message : ''
  };

  if (run.status !== S.STATUS.RUNNING) return Promise.resolve({ act: 'idle', hud: hud });
  if (run.tabId == null && tabId != null) run.tabId = tabId;
  if (tabId != null && run.tabId !== tabId) return Promise.resolve({ act: 'idle', hud: null });
  if (!row) return Promise.resolve({ act: 'idle', hud: hud });

  var wantRoute = S.STEP_ROUTE[run.step];
  if (msg.route !== wantRoute) {
    // The page is not where this step belongs. Say nothing and wait for the
    // navigation to land rather than acting on the wrong screen.
    if (!run.waitingSince) run.waitingSince = Date.now();

    // If it never lands, the run would sit here looking busy forever. Waiting
    // far longer than a step could legitimately take means something moved the
    // browser somewhere unexpected, so ask the user rather than hang.
    var patience = Math.max(60000, (run.settings.stepTimeoutMs || 20000) * 3);
    if (Date.now() - run.waitingSince > patience) {
      run.waitingSince = null;
      raiseAttention(bucket,
        'The browser is on the ' + (msg.route || 'unrecognised') + ' page but step "' +
        (S.STEP_LABEL[run.step] || run.step) + '" needs the ' + wantRoute + ' page. ' +
        'Go back to that page and press Retry, or skip this row.', msg.url);
      return Promise.resolve({ act: 'idle', hud: hud });
    }
    return Promise.resolve({ act: 'wait', expect: wantRoute, step: run.step, hud: hud });
  }

  run.waitingSince = null;
  return Promise.resolve({
    act: 'run',
    step: run.step,
    attempt: run.attempt,
    settings: run.settings,
    hud: hud,
    row: {
      sheetRow: row.sheetRow,
      code: row.code,
      name: row.name,
      priceText: row.priceText,
      productUuid: row.productUuid || ''
    }
  });
};

handlers[S.MSG.STEP_DONE] = function (msg, bucket, sender) {
  var run = bucket.run;
  if (run.status !== S.STATUS.RUNNING) return Promise.resolve({ ok: true });
  if (msg.step !== run.step) return Promise.resolve({ ok: true, stale: true });

  var row = currentRow(bucket);
  if (!row) return Promise.resolve({ ok: true });

  run.attempt = 0;
  run.waitingSince = null;

  if (msg.productUuid) {
    row.productUuid = msg.productUuid;
    bucket.rowsDirty = true;
  }

  // A dry run stops once the product form is filled, because going further
  // would mean saving a record.
  if (msg.step === S.STEP.FILL_PRODUCT && run.settings.dryRun) {
    run.status = S.STATUS.DONE;
    run.finishedAt = Date.now();
    log(bucket, 'success',
      'Dry run complete. The product form is filled for row ' + row.sheetRow +
      ' and was not saved. Check the values on screen, then turn Dry run off.');
    bucket.logDirty = true;
    notify('Dry run complete', 'The form is filled and was not saved. Check it, then turn Dry run off.');
    return Promise.resolve({ ok: true });
  }

  if (msg.step === S.STEP.CAPTURE_LINK) {
    var link = String(msg.link || '').trim();
    if (!link) {
      raiseAttention(bucket, 'The link page did not contain a Fygaro link.', msg.url);
      return Promise.resolve({ ok: true });
    }
    log(bucket, 'success', 'Row ' + row.sheetRow + ' (' + row.code + ') done: ' + link);
    bucket.logDirty = true;
    return completeRow(bucket, S.ROW.DONE, link).then(function () {
      return { ok: true };
    });
  }

  // Saving an item normally returns to the list, but if Fygaro ever drops
  // straight onto the new product there is nothing left to look up.
  if (msg.step === S.STEP.FILL_PRODUCT && msg.landedRoute === 'productDetail') {
    run.step = S.STEP.CLICK_CREATE_LINK;
    log(bucket, 'info', 'Saving opened the product directly, so the lookup step was not needed.');
    bucket.logDirty = true;
    return Promise.resolve({ ok: true });
  }

  var at = S.STEP_ORDER.indexOf(msg.step);
  run.step = S.STEP_ORDER[at + 1] || S.STEP.NAV_TO_PRODUCTS;
  return Promise.resolve({ ok: true });
};

handlers[S.MSG.STEP_FAILED] = function (msg, bucket) {
  var run = bucket.run;
  if (run.status !== S.STATUS.RUNNING) return Promise.resolve({ ok: true });
  if (msg.step !== run.step) return Promise.resolve({ ok: true, stale: true });

  run.attempt = (run.attempt || 0) + 1;
  if (run.attempt < (run.settings.maxAttempts || 1)) {
    log(bucket, 'warn', 'Step "' + (S.STEP_LABEL[msg.step] || msg.step) + '" failed, retrying. ' + msg.message);
    bucket.logDirty = true;
    return Promise.resolve({ ok: true, retry: true });
  }

  raiseAttention(bucket, msg.message || 'The step could not be completed.', msg.url);
  return Promise.resolve({ ok: true });
};

handlers[S.MSG.CONTENT_LOG] = function (msg, bucket) {
  log(bucket, msg.level || 'info', msg.message);
  bucket.logDirty = true;
  return Promise.resolve({ ok: true });
};

/* ------------------------------------------------------------------- tabs */

/*
 * Prefers a tab already inside the Fygaro app. Falling back to any fygaro.com
 * tab would be enough to find one, but that tab could be a payment link the user
 * opened to look at, and it is about to be navigated away.
 */
function findOrOpenTab() {
  return chrome.tabs.query({ url: [S.ORIGIN + '/en/app/*', S.ORIGIN + '/es/app/*'] })
    .then(function (tabs) {
      if (tabs && tabs.length) return tabs[0].id;
      return chrome.tabs.create({ url: S.APP_URL, active: true }).then(function (tab) { return tab.id; });
    });
}

chrome.tabs.onRemoved.addListener(function (tabId) {
  serial(function () {
    return S.read().then(function (bucket) {
      if (bucket.run.tabId !== tabId) return null;
      bucket.run.tabId = null;
      if (bucket.run.status === S.STATUS.RUNNING) {
        raiseAttention(bucket, 'The Fygaro tab was closed. Reopen it and press Retry.', '');
      }
      return persist(bucket);
    });
  });
});

/* --------------------------------------------------------------- lifecycle */

chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
  var handler = msg && handlers[msg.type];
  if (!handler) return false;

  serial(function () {
    return S.read().then(function (bucket) {
      bucket.rowsDirty = false;
      bucket.logDirty = false;
      // The content script asks for work every couple of seconds for hours on
      // end. Most of those answers change nothing, so compare before and after
      // and skip the write and the redraw when the run is untouched.
      var before = JSON.stringify(bucket.run);
      return Promise.resolve(handler(msg, bucket, sender)).then(function (result) {
        result = result || { ok: true };
        if (result.handled) return result;
        var changed = bucket.rowsDirty || bucket.logDirty || JSON.stringify(bucket.run) !== before;
        if (!changed) return result;
        return persist(bucket).then(function () { return result; });
      });
    });
  }).then(sendResponse, function (err) {
    sendResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  });

  return true; // response is asynchronous
});

function enableSidePanel() {
  if (chrome.sidePanel && chrome.sidePanel.setPanelBehavior) {
    chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(function () {});
  }
}

chrome.runtime.onInstalled.addListener(function () {
  enableSidePanel();
  serial(function () {
    return S.read().then(function (bucket) {
      // A worker restart must never leave the UI claiming work is in flight.
      if (bucket.run.status === S.STATUS.RUNNING) {
        bucket.run.status = S.STATUS.PAUSED;
        log(bucket, 'warn', 'The extension reloaded mid run. Press Resume to carry on.');
        bucket.logDirty = true;
      }
      return persist(bucket);
    });
  });
});

chrome.runtime.onStartup.addListener(enableSidePanel);
enableSidePanel();
