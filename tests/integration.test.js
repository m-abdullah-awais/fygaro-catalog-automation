/*
 * Fygaro Catalog Automation
 * End to end test driver, loaded by tests/integration.test.html.
 *
 * Nothing in src/ is reimplemented here. The Chrome APIs the extension uses are
 * stubbed in memory, then the real service worker and the real side panel are
 * loaded and put through a full run.
 */
(function () {
  'use strict';

  var U = FYG.util;
  var S = FYG.state;
  var X = FYG.xlsx;

  var WORKBOOK_PATH = '../docs/Cat%C3%A1logo%20de%20Productos%20y%20Servicios%20Fygaro.xlsx';
  var WORKBOOK_NAME = 'Catálogo de Productos y Servicios Fygaro.xlsx';
  var SHEET = 'Logros ';

  var results = [];
  var recorded = { badges: [], notifications: [], tabUpdates: [], downloads: [], actions: [] };
  var currentZoom = 1;

  function assert(condition, message) {
    if (!condition) throw new Error(message);
  }

  function check(name, fn) {
    return Promise.resolve().then(fn)
      .then(function (detail) { results.push({ ok: true, name: name, detail: detail || '' }); })
      .catch(function (err) {
        results.push({ ok: false, name: name, detail: String(err && err.message ? err.message : err) });
      });
  }

  function waitUntil(probe, label, timeoutMs) {
    var deadline = Date.now() + (timeoutMs || 5000);
    return new Promise(function (resolve, reject) {
      (function poll() {
        var value;
        try { value = probe(); } catch (e) { value = null; }
        if (value) return resolve(value);
        if (Date.now() > deadline) return reject(new Error('Timed out waiting for ' + label));
        setTimeout(poll, 20);
      })();
    });
  }

  /* ------------------------------------------------------ Chrome API stubs */

  var store = {};
  var messageListeners = [];
  var storageListeners = [];
  var nextTabId = 100;

  function clone(value) {
    return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
  }

  window.importScripts = function () {}; // the shared files are already loaded

  window.chrome = {
    runtime: {
      lastError: null,
      getURL: function (path) { return 'chrome-extension://test/' + path; },
      sendMessage: function (message) {
        return new Promise(function (resolve) {
          var answered = false;
          var settle = function (response) {
            if (answered) return;
            answered = true;
            resolve(clone(response));
          };
          var wantsAsync = false;
          messageListeners.forEach(function (listener) {
            var result = listener(clone(message), { tab: null }, settle);
            if (result === true) wantsAsync = true;
          });
          // No listener took the message, so resolve like Chrome would.
          if (!wantsAsync) setTimeout(function () { settle(undefined); }, 0);
        });
      },
      onMessage: { addListener: function (fn) { messageListeners.push(fn); } },
      onInstalled: { addListener: function () {} },
      onStartup: { addListener: function () {} }
    },
    storage: {
      local: {
        get: function (keys) {
          var wanted = typeof keys === 'string' ? [keys] : keys;
          var out = {};
          wanted.forEach(function (k) { if (k in store) out[k] = clone(store[k]); });
          return Promise.resolve(out);
        },
        set: function (items) {
          var changes = {};
          Object.keys(items).forEach(function (k) {
            changes[k] = { oldValue: clone(store[k]), newValue: clone(items[k]) };
            store[k] = clone(items[k]);
          });
          storageListeners.forEach(function (fn) { fn(changes, 'local'); });
          return Promise.resolve();
        },
        remove: function (keys) {
          var list = typeof keys === 'string' ? [keys] : keys;
          var changes = {};
          list.forEach(function (k) { changes[k] = { oldValue: clone(store[k]) }; delete store[k]; });
          storageListeners.forEach(function (fn) { fn(changes, 'local'); });
          return Promise.resolve();
        }
      },
      onChanged: { addListener: function (fn) { storageListeners.push(fn); } }
    },
    action: {
      setBadgeText: function (o) { recorded.badges.push(o.text); return Promise.resolve(); },
      setBadgeBackgroundColor: function () { return Promise.resolve(); }
    },
    notifications: {
      create: function (options, cb) { recorded.notifications.push(options); if (cb) cb('id'); }
    },
    tabs: {
      query: function () { return Promise.resolve([]); },
      create: function (props) {
        var tab = { id: nextTabId++, url: props.url };
        recorded.tabUpdates.push(props.url);
        recorded.actions.push('nav:' + props.url);
        return Promise.resolve(tab);
      },
      update: function (id, props) {
        if (props.url) {
          recorded.tabUpdates.push(props.url);
          recorded.actions.push('nav:' + props.url);
        }
        return Promise.resolve({ id: id });
      },
      getZoom: function () { return Promise.resolve(currentZoom); },
      setZoom: function (id, factor) {
        currentZoom = factor;
        recorded.actions.push('zoom:' + factor);
        return Promise.resolve();
      },
      onRemoved: { addListener: function () {} }
    },
    sidePanel: { setPanelBehavior: function () { return Promise.resolve(); } }
  };

  // Catch exports instead of writing files. The anchor click is intercepted too,
  // because a real download keeps headless Chrome alive waiting to save it.
  var realCreateObjectURL = URL.createObjectURL;
  URL.createObjectURL = function (blob) {
    recorded.downloads.push(blob);
    return realCreateObjectURL(blob);
  };

  var realAnchorClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    if (this.hasAttribute('download')) return;
    return realAnchorClick.apply(this, arguments);
  };

  /* ------------------------------------------------------------ page setup */

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var el = document.createElement('script');
      el.src = src;
      el.onload = resolve;
      el.onerror = function () { reject(new Error('Could not load ' + src)); };
      document.body.appendChild(el);
    });
  }

  function mountPanelMarkup() {
    return fetch('../src/sidepanel/sidepanel.html')
      .then(function (r) { return r.text(); })
      .then(function (html) {
        var doc = new DOMParser().parseFromString(html, 'text/html');
        var app = doc.querySelector('.app');
        assert(app, 'sidepanel.html has no .app container');
        document.getElementById('panel').appendChild(document.importNode(app, true));
      });
  }

  /* -------------------------------------------------- content script mimic */

  var ROUTE_URL = {
    dashboard: 'https://www.fygaro.com/en/app/dashboard/',
    productList: 'https://www.fygaro.com/en/app/products/',
    productAdd: 'https://www.fygaro.com/en/app/products/add/',
    productDetail: 'https://www.fygaro.com/en/app/products/UUID/permalink/product/',
    linkAdd: 'https://www.fygaro.com/en/app/payment-buttons/payments/payment-buttons/add/',
    linkDone: 'https://www.fygaro.com/en/app/payment-buttons/payments/UUID/payment-buttons/permalink/'
  };

  function askForJob(route) {
    return chrome.runtime.sendMessage({
      type: S.MSG.REQUEST_JOB, route: route, url: ROUTE_URL[route]
    });
  }

  function finishStep(step, extra) {
    return chrome.runtime.sendMessage(Object.assign({
      type: S.MSG.STEP_DONE, step: step, url: ROUTE_URL[S.STEP_ROUTE[step]]
    }, extra || {}));
  }

  /** Plays one row through all seven steps, exactly as the content script would. */
  function playRow(uuid, link) {
    var seen = [];
    function stage(route, extra) {
      return askForJob(route).then(function (job) {
        assert(job && job.act === 'run',
          'expected work on ' + route + ' but got "' + (job && job.act) + '"');
        seen.push(job.step);
        return finishStep(job.step, extra);
      });
    }
    return stage('dashboard')
      .then(function () { return stage('productList'); })
      .then(function () { return stage('productAdd', { landedRoute: 'productList' }); })
      .then(function () { return stage('productList', { productUuid: uuid }); })
      .then(function () { return stage('productDetail'); })
      .then(function () { return stage('linkAdd'); })
      .then(function () { return stage('linkDone', { link: link }); })
      .then(function () { return seen; });
  }

  function readState() {
    return S.read();
  }

  var $ = function (id) { return document.getElementById(id); };

  /* ----------------------------------------------------------------- suite */

  var workbookBytes = null;

  function suite() {
    return check('the panel markup and script load without errors', function () {
      return waitUntil(function () { return $('statusPill'); }, 'the panel to mount')
        .then(function () {
          assert($('statusPill').textContent === 'Idle', 'starts as "' + $('statusPill').textContent + '"');
          assert($('btnStart').disabled, 'Start should be disabled before a catalog is loaded');
          assert($('btnExportXlsx').disabled, 'Export should be disabled before any link exists');
          return 'panel mounted, controls correctly disabled';
        });
    })

    .then(function () {
      return check('choosing the catalog file loads all 699 rows', function () {
        var file = new File([workbookBytes], WORKBOOK_NAME,
          { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var transfer = new DataTransfer();
        transfer.items.add(file);
        $('fileInput').files = transfer.files;
        $('fileInput').dispatchEvent(new Event('change'));

        return waitUntil(function () { return $('statTotal').textContent === '699'; }, 'the rows to load', 20000)
          .then(function () {
            assert($('sheetSelect').value === SHEET, 'sheet is "' + $('sheetSelect').value + '"');
            assert($('mapName').value === 'D', 'Name mapped to ' + $('mapName').value);
            assert($('mapCode').value === 'C', 'Code mapped to ' + $('mapCode').value);
            assert($('mapPrice').value === 'E', 'Price mapped to ' + $('mapPrice').value);
            assert($('mapLink').value === 'H', 'Link mapped to ' + $('mapLink').value);
            assert($('statSkipped').textContent === '1', 'expected 1 pre-skipped row, saw ' + $('statSkipped').textContent);
            assert(!$('btnStart').disabled, 'Start should now be enabled');
            return '699 rows, sheet and columns detected automatically';
          });
      });
    })

    .then(function () {
      return check('the row that already had a link is skipped, not redone', function () {
        return readState().then(function (data) {
          var row18 = data.rows.filter(function (r) { return r.sheetRow === 18; })[0];
          assert(row18, 'row 18 is missing');
          assert(row18.status === S.ROW.SKIPPED, 'row 18 is "' + row18.status + '"');
          assert(data.rows[0].status === S.ROW.PENDING, 'row 2 should be pending');
          return 'row 18 skipped because column H was already filled';
        });
      });
    })

    .then(function () {
      return check('Start opens Fygaro and begins at the first unfinished row', function () {
        $('btnStart').click();
        return waitUntil(function () { return $('statusPill').dataset.status === 'running'; }, 'the run to start')
          .then(readState)
          .then(function (data) {
            assert(data.run.cursor === 0, 'cursor is ' + data.run.cursor);
            assert(data.run.step === S.STEP.NAV_TO_PRODUCTS, 'first step is ' + data.run.step);
            assert(data.rows[0].status === S.ROW.ACTIVE, 'first row is "' + data.rows[0].status + '"');
            assert(recorded.tabUpdates.indexOf(S.APP_URL) !== -1, 'the Fygaro tab was not opened');
            return 'running at sheet row ' + data.rows[0].sheetRow;
          });
      });
    })

    .then(function () {
      return check('the page is zoomed out before the run navigates to work', function () {
        assert(Math.abs(currentZoom - 0.67) < 0.001, 'the page zoom is ' + currentZoom + ', expected 0.67');

        // Chrome keys zoom to the origin, so a fygaro.com page has to be loaded
        // before the zoom can be set at all. What must be guaranteed is that the
        // navigation the run actually starts from happens after the zoom.
        var zoomAt = recorded.actions.indexOf('zoom:0.67');
        var lastNavAt = recorded.actions.lastIndexOf('nav:' + S.APP_URL);
        assert(zoomAt !== -1, 'the zoom was never applied');
        assert(lastNavAt !== -1, 'the tab was never sent to the dashboard');
        assert(zoomAt < lastNavAt, 'the run must reload the page after zooming, order was ' +
          recorded.actions.join(' then '));

        return readState().then(function (data) {
          assert(data.run.originalZoom === 1, 'the original zoom was recorded as ' + data.run.originalZoom);
          return 'zoomed to 67%, then reloaded at that zoom, original 100% remembered';
        });
      });
    })

    .then(function () {
      return check('a step is only offered on the page it belongs to', function () {
        return askForJob('linkAdd').then(function (job) {
          assert(job.act === 'wait', 'expected to wait on the wrong page, got "' + job.act + '"');
          assert(job.expect === 'productList' || job.expect === 'dashboard',
            'expected route is ' + job.expect);
          return 'the wrong page yields "wait", never an action';
        });
      });
    })

    .then(function () {
      return check('the first row walks all seven steps in order and captures its link', function () {
        var link = 'https://www.fygaro.com/en/pb/aaaaaaaa-1111-1111-1111-111111111111/';
        return playRow('aaaaaaaa-0000-0000-0000-000000000001', link).then(function (seen) {
          assert(seen.join(',') === S.STEP_ORDER.join(','), 'steps ran as ' + seen.join(','));
          return readState().then(function (data) {
            assert(data.rows[0].status === S.ROW.DONE, 'row is "' + data.rows[0].status + '"');
            assert(data.rows[0].link === link, 'link is "' + data.rows[0].link + '"');
            assert(data.run.stats.done === 1, 'done count is ' + data.run.stats.done);
            assert(data.run.cursor === 1, 'cursor did not advance, it is ' + data.run.cursor);
            assert(data.run.step === S.STEP.NAV_TO_PRODUCTS, 'next row should start at step 1');
            return seen.length + ' steps, link stored on sheet row ' + data.rows[0].sheetRow;
          });
        });
      });
    })

    .then(function () {
      return check('the panel shows the captured link and enables export', function () {
        return waitUntil(function () { return !$('btnExportXlsx').disabled; }, 'export to become available')
          .then(function () {
            assert(!$('btnExportCsv').disabled, 'CSV export should be enabled');
            assert($('statDone').textContent === '1', 'done stat is ' + $('statDone').textContent);
            assert($('progressText').textContent.indexOf('of 699') !== -1,
              'progress reads "' + $('progressText').textContent + '"');
            return $('progressText').textContent;
          });
      });
    })

    .then(function () {
      return check('a second row runs and the skipped row 18 is never offered', function () {
        var link = 'https://www.fygaro.com/en/pb/aaaaaaaa-2222-2222-2222-222222222222/';
        return playRow('aaaaaaaa-0000-0000-0000-000000000002', link).then(function () {
          return readState();
        }).then(function (data) {
          assert(data.run.stats.done === 2, 'done count is ' + data.run.stats.done);
          var row18 = data.rows.filter(function (r) { return r.sheetRow === 18; })[0];
          assert(row18.status === S.ROW.SKIPPED, 'row 18 changed to ' + row18.status);
          assert(data.rows[data.run.cursor].sheetRow !== 18, 'the cursor landed on the skipped row');
          return 'two rows done, row 18 left alone';
        });
      });
    })

    .then(function () {
      return check('a failing step retries, then asks the user', function () {
        return readState().then(function (before) {
          var step = before.run.step;
          var attempts = before.run.settings.maxAttempts;
          function failOnce() {
            return chrome.runtime.sendMessage({
              type: S.MSG.STEP_FAILED, step: step, url: 'https://www.fygaro.com/en/app/dashboard/',
              message: 'The Products link was not found in the sidebar.'
            });
          }
          var chain = Promise.resolve();
          for (var i = 0; i < attempts - 1; i++) chain = chain.then(failOnce);

          return chain.then(readState).then(function (mid) {
            assert(mid.run.status === S.STATUS.RUNNING, 'should still be retrying, is ' + mid.run.status);
            return failOnce();
          }).then(function () {
            return waitUntil(function () { return $('statusPill').dataset.status === 'attention'; },
              'the attention banner');
          }).then(readState).then(function (after) {
            assert(after.run.pending, 'no pending failure was recorded');
            assert(after.run.pending.message.indexOf('Products link') !== -1,
              'message is "' + after.run.pending.message + '"');
            assert(!$('attentionBanner').classList.contains('hidden'), 'the banner is hidden');
            assert($('attentionMessage').textContent === after.run.pending.message, 'the banner text is wrong');
            assert(recorded.notifications.length > 0, 'no desktop notification was raised');
            return 'retried ' + (attempts - 1) + ' time(s), then paused for the user';
          });
        });
      });
    })

    .then(function () {
      return check('no work is handed out while the run waits for the user', function () {
        return askForJob('dashboard').then(function (job) {
          assert(job.act === 'idle', 'expected idle while paused, got "' + job.act + '"');
          return 'idle until the user decides';
        });
      });
    })

    .then(function () {
      return check('Skip this row marks it failed and moves on cleanly', function () {
        return readState().then(function (before) {
          var skipped = before.rows[before.run.cursor];
          $('btnSkip').click();
          return waitUntil(function () { return $('statusPill').dataset.status === 'running'; }, 'the run to resume')
            .then(readState)
            .then(function (after) {
              assert(after.rows[skipped.i].status === S.ROW.FAILED, 'the row is "' + after.rows[skipped.i].status + '"');
              assert(after.rows[skipped.i].error, 'no reason was recorded on the row');
              assert(after.run.cursor > skipped.i, 'the cursor did not advance');
              assert(after.run.step === S.STEP.NAV_TO_PRODUCTS, 'the next row must restart at step 1');
              assert(after.run.pending === null, 'the pending failure was not cleared');
              assert(recorded.tabUpdates[recorded.tabUpdates.length - 1] === S.APP_URL,
                'skipping should return to the dashboard');
              return 'row ' + skipped.sheetRow + ' marked failed, run continues at row ' +
                after.rows[after.run.cursor].sheetRow;
            });
        });
      });
    })

    .then(function () {
      return check('Pause stops handing out work and Resume restarts it', function () {
        $('btnPause').click();
        return waitUntil(function () { return $('statusPill').dataset.status === 'paused'; }, 'the pause')
          .then(function () { return askForJob('dashboard'); })
          .then(function (job) {
            assert(job.act === 'idle', 'work was handed out while paused');
            $('btnStart').click();
            return waitUntil(function () { return $('statusPill').dataset.status === 'running'; }, 'the resume');
          })
          .then(function () { return askForJob('dashboard'); })
          .then(function (job) {
            assert(job.act === 'run', 'no work after resuming, got "' + job.act + '"');
            return 'paused and resumed without losing the cursor';
          });
      });
    })

    .then(function () {
      return check('declining the reload prompt keeps the captured links and the mapping', function () {
        var realConfirm = window.confirm;
        var asked = null;
        window.confirm = function (text) { asked = text; return false; };

        return readState().then(function (before) {
          $('mapCode').value = 'B';
          $('mapCode').dispatchEvent(new Event('change'));
          return waitUntil(function () { return asked !== null; }, 'the confirmation prompt');
        }).then(function () {
          window.confirm = realConfirm;
          assert(/clears the 2 links/.test(asked), 'the prompt did not name the cost: "' + asked + '"');
          return readState();
        }).then(function (after) {
          assert(after.run.file.mapping.code === 'C', 'the stored mapping changed to ' + after.run.file.mapping.code);
          assert(after.rows.filter(function (r) { return r.link; }).length === 3,
            'links were lost, ' + after.rows.filter(function (r) { return r.link; }).length + ' remain');
          assert($('mapCode').value === 'C', 'the dropdown still shows ' + $('mapCode').value);
          return 'nothing reloaded, dropdown reverted to C';
        }).catch(function (err) {
          window.confirm = realConfirm;
          throw err;
        });
      });
    })

    .then(function () {
      return check('a navigation that never lands is escalated instead of hanging forever', function () {
        // The browser ends up somewhere the current step cannot act on. Waiting
        // is right at first, but it must not wait for ever.
        return askForJob('linkDone').then(function (job) {
          assert(job.act === 'wait', 'the first wrong page should simply wait, got "' + job.act + '"');

          var realNow = Date.now;
          Date.now = function () { return realNow() + 200000; };
          return askForJob('linkDone').then(function (later) {
            Date.now = realNow;
            assert(later.act === 'idle', 'expected the run to stop handing out work, got "' + later.act + '"');
            return readState();
          }, function (err) {
            Date.now = realNow;
            throw err;
          });
        }).then(function (data) {
          assert(data.run.status === S.STATUS.ATTENTION, 'status is ' + data.run.status);
          assert(/needs the .* page/.test(data.run.pending.message),
            'unhelpful message: "' + data.run.pending.message + '"');
          // Put the run back so the remaining checks continue from a clean state.
          return chrome.runtime.sendMessage({ type: S.MSG.RESUME });
        }).then(function () {
          return 'waited, then asked the user rather than stalling silently';
        });
      });
    })

    .then(function () {
      return check('the exported workbook carries the captured links in the right rows', function () {
        return readState().then(function (data) {
          var done = data.rows.filter(function (r) { return r.status === S.ROW.DONE; });
          assert(done.length === 2, 'expected 2 finished rows, saw ' + done.length);

          recorded.downloads.length = 0;
          $('btnExportXlsx').click();

          return waitUntil(function () { return recorded.downloads.length > 0; }, 'the export', 20000)
            .then(function () { return recorded.downloads[0].arrayBuffer(); })
            .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
            .then(function (wb) {
              var sheet = wb.readSheet(SHEET);
              var byRow = {};
              sheet.rows.forEach(function (r) { byRow[r.r] = r.cells; });

              done.forEach(function (row) {
                assert(byRow[row.sheetRow].H === row.link,
                  'sheet row ' + row.sheetRow + ' has "' + byRow[row.sheetRow].H + '" instead of its link');
              });
              // Everything else must be exactly as it was.
              assert(byRow[18].H.indexOf('/es/pb/5c03f135') !== -1, 'the pre-existing link was disturbed');
              assert(byRow[2].D === 'Cita de Ingreso Presencial con Psicóloga Clínica del Equipo',
                'row 2 name changed');
              assert(wb.sheets.length === 6, 'sheets were lost');

              // Rows the run never reached must still be empty.
              var touched = done.map(function (r) { return r.sheetRow; }).concat([18]);
              var untouched = data.rows.filter(function (r) { return touched.indexOf(r.sheetRow) === -1; });
              assert(untouched.length > 600, 'expected most rows to be untouched');
              untouched.slice(0, 50).forEach(function (r) {
                assert(!byRow[r.sheetRow].H,
                  'sheet row ' + r.sheetRow + ' should be empty, saw "' + byRow[r.sheetRow].H + '"');
              });
              return done.length + ' links written to sheet rows ' +
                done.map(function (r) { return r.sheetRow; }).join(' and ');
            });
        });
      });
    })

    .then(function () {
      return check('stopping hands the page zoom back to the user', function () {
        assert(Math.abs(currentZoom - 0.67) < 0.001, 'precondition: the run should still be zoomed out');
        $('btnStop').click();
        return waitUntil(function () { return $('statusPill').dataset.status === 'idle'; }, 'the stop')
          .then(function () {
            return waitUntil(function () { return Math.abs(currentZoom - 1) < 0.001; },
              'the zoom to be restored');
          })
          .then(readState)
          .then(function (data) {
            assert(data.run.originalZoom === null, 'the remembered zoom should be cleared once given back');
            return 'page returned to 100%';
          });
      });
    })

    .then(function () {
      return check('a dry run stops with the form filled and creates nothing', function () {
        $('btnStop').click();
        return waitUntil(function () { return $('statusPill').dataset.status === 'idle'; }, 'the stop')
          .then(function () {
            $('dryRun').checked = true;
            $('dryRun').dispatchEvent(new Event('change'));
            return waitUntil(function () { return $('runHint').textContent.indexOf('Dry run') !== -1; },
              'the dry run hint');
          })
          .then(function () {
            $('btnStart').click();
            return waitUntil(function () { return $('statusPill').dataset.status === 'running'; }, 'the dry run');
          })
          .then(function () { return askForJob('dashboard'); })
          .then(function (job) { return finishStep(job.step); })
          .then(function () { return askForJob('productList'); })
          .then(function (job) { return finishStep(job.step); })
          .then(function () { return askForJob('productAdd'); })
          .then(function (job) {
            assert(job.step === S.STEP.FILL_PRODUCT, 'expected the form step, got ' + job.step);
            assert(job.settings.dryRun === true, 'the dry run flag did not reach the page');
            // The content script fills the form and reports back without saving.
            return finishStep(job.step);
          })
          .then(function () {
            return waitUntil(function () { return $('statusPill').dataset.status === 'done'; },
              'the dry run to finish');
          })
          .then(readState)
          .then(function (data) {
            assert(data.run.status === S.STATUS.DONE, 'status is ' + data.run.status);
            var active = data.rows.filter(function (r) { return r.status === S.ROW.ACTIVE; });
            assert(active.length === 1, 'the row should still be waiting, not completed');
            assert(!active[0].link, 'a dry run must not record a link');
            var last = data.log[data.log.length - 1];
            assert(last.message.indexOf('Dry run complete') !== -1, 'last log line is "' + last.message + '"');
            return 'stopped at the filled form, nothing recorded';
          });
      });
    })

    .then(function () {
      return check('Clear all wipes the run, the rows, the log and the cached file', function () {
        return chrome.runtime.sendMessage({ type: S.MSG.RESET }).then(function () {
          return waitUntil(function () { return !(S.KEY_ROWS in store); }, 'the reset');
        }).then(function () {
          assert(!(S.KEY_RUN in store), 'the run was left behind');
          assert(!(S.KEY_LOG in store), 'the log was left behind');
          assert(!(S.KEY_FILE in store), 'the cached workbook was left behind');
          return 'all four storage keys cleared';
        });
      });
    })

    .then(report);
  }

  function report() {
    var failed = results.filter(function (r) { return !r.ok; });
    var summary = document.getElementById('summary');
    summary.className = failed.length ? 'fail' : 'pass';
    summary.textContent = failed.length
      ? failed.length + ' of ' + results.length + ' checks FAILED'
      : 'All ' + results.length + ' checks passed';
    document.title = (failed.length ? 'FAIL ' + failed.length + '/' : 'PASS 0/') + results.length;

    var list = document.getElementById('results');
    list.innerHTML = '';
    results.forEach(function (r) {
      var li = document.createElement('li');
      li.className = r.ok ? 'ok' : 'bad';
      var tag = document.createElement('span');
      tag.className = 'tag';
      tag.textContent = r.ok ? 'PASS' : 'FAIL';
      var name = document.createElement('span');
      name.textContent = r.name;
      var detail = document.createElement('div');
      detail.className = 'detail';
      detail.textContent = r.detail;
      li.appendChild(tag);
      li.appendChild(name);
      li.appendChild(detail);
      list.appendChild(li);
    });
  }

  /* ------------------------------------------------------------------ boot */

  fetch(WORKBOOK_PATH)
    .then(function (response) {
      if (!response.ok) throw new Error('the workbook could not be read: status ' + response.status);
      return response.arrayBuffer();
    })
    .then(function (buffer) {
      workbookBytes = buffer;
      return mountPanelMarkup();
    })
    // The worker registers its message listener first, exactly as in the browser.
    .then(function () { return loadScript('../src/background/service-worker.js'); })
    .then(function () { return loadScript('../src/sidepanel/sidepanel.js'); })
    .then(suite)
    .catch(function (err) {
      results.push({ ok: false, name: 'test harness', detail: String(err && err.message ? err.message : err) });
      report();
    });
})();
