/*
 * Fygaro Catalog Automation
 * Side panel: the whole user interface.
 *
 * The panel never decides anything about the run. It loads the catalog, sends
 * commands to the background worker, and renders whatever state comes back, so
 * closing and reopening it cannot lose or contradict a run in progress.
 */
(function () {
  'use strict';

  var U = FYG.util;
  var S = FYG.state;
  var X = FYG.xlsx;

  var RESULT_LIMIT = 150;

  /* The parsed workbook, kept so the updated file can be exported. It is
   * restored from storage when the panel is reopened. */
  var workbook = null;
  var fileName = '';
  var sheetData = null;
  var headerInfo = null;

  /* A handle to the file on disk, when the browser gave us one. It is what
   * makes "update the original" possible: a plain file input only yields a
   * snapshot of the bytes and no way to write back. Handles cannot go in
   * chrome.storage because they are not JSON, so they live in IndexedDB. */
  var fileHandle = null;

  /* Confirmation of the last successful save, shown in the Export card. */
  var lastSaveMessage = '';

  /* Columns this sheet does not have and the export therefore has to create,
   * along with the headers to write above them. Empty when the sheet already
   * had a link column, which is the case on every run after the first. */
  var proposed = { link: '', note: '' };

  /* Where the reason goes for a row that finished without a link. */
  var noteColumn = '';

  var view = { run: S.defaultRun(), rows: [], log: [] };
  var filter = 'all';
  var searchTerm = '';
  /* Set while the user is deliberately picking a different file, so a re-render
   * triggered by the running job cannot snap the picker shut under them. */
  var choosingFile = false;

  var $ = function (id) { return document.getElementById(id); };

  function send(type, payload) {
    return chrome.runtime.sendMessage(Object.assign({ type: type }, payload || {}))
      .catch(function () { return null; });
  }

  function setHidden(el, hidden) {
    el.classList.toggle('hidden', !!hidden);
  }

  /* ------------------------------------------------------- workbook storage */

  function bytesToBase64(bytes) {
    var chunk = 0x8000;
    var parts = [];
    for (var i = 0; i < bytes.length; i += chunk) {
      parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
    }
    return btoa(parts.join(''));
  }

  function base64ToBytes(b64) {
    var binary = atob(b64);
    var out = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  }

  function storeWorkbook(name, bytes) {
    var payload = {};
    payload[S.KEY_FILE] = { name: name, data: bytesToBase64(bytes) };
    return chrome.storage.local.set(payload).catch(function (err) {
      // Not fatal: the run still works, only the xlsx export needs the file back.
      note('warn', 'The catalog file could not be cached: ' + err.message);
    });
  }

  function restoreWorkbook() {
    return chrome.storage.local.get(S.KEY_FILE).then(function (data) {
      var saved = data[S.KEY_FILE];
      if (!saved || !saved.data) return null;
      return X.load(base64ToBytes(saved.data)).then(function (wb) {
        workbook = wb;
        fileName = saved.name;
        return wb;
      });
    }).catch(function () { return null; });
  }

  function note(level, message) {
    send(S.MSG.CONTENT_LOG, { level: level, message: message });
  }

  /* ---------------------------------------------------------- file handles */

  var HANDLE_DB = 'fygaro-file';
  var HANDLE_STORE = 'handles';

  function openHandleDb() {
    return new Promise(function (resolve, reject) {
      // IndexedDB is absent in some contexts, a file:// page among them.
      if (typeof indexedDB === 'undefined' || !indexedDB) {
        reject(new Error('IndexedDB is not available here.'));
        return;
      }
      var request;
      try {
        request = indexedDB.open(HANDLE_DB, 1);
      } catch (err) {
        reject(err);
        return;
      }
      request.onupgradeneeded = function () {
        if (!request.result.objectStoreNames.contains(HANDLE_STORE)) {
          request.result.createObjectStore(HANDLE_STORE);
        }
      };
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error || new Error('IndexedDB could not be opened.')); };
      request.onblocked = function () { reject(new Error('IndexedDB is blocked by another tab.')); };
    });
  }

  /**
   * Every path here is time limited. Remembering a file handle is a convenience,
   * so if the database misbehaves the panel must carry on rather than sit
   * waiting for an event that may never arrive.
   */
  function handleStore(mode, action) {
    var work = openHandleDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var request;
        try {
          var tx = db.transaction(HANDLE_STORE, mode);
          request = action(tx.objectStore(HANDLE_STORE));
          tx.oncomplete = function () { db.close(); resolve(request ? request.result : undefined); };
          tx.onerror = function () { db.close(); reject(tx.error); };
          tx.onabort = function () { db.close(); reject(tx.error || new Error('The write was aborted.')); };
        } catch (err) {
          db.close();
          reject(err);
        }
      });
    });

    var timeout = new Promise(function (resolve, reject) {
      setTimeout(function () { reject(new Error('IndexedDB did not respond.')); }, 3000);
    });
    return Promise.race([work, timeout]);
  }

  function rememberHandle(handle) {
    fileHandle = handle;
    if (!handle) return Promise.resolve();
    return handleStore('readwrite', function (store) { return store.put(handle, 'workbook'); })
      .catch(function () { /* the run still works, only in place saving needs it */ });
  }

  function forgetHandle() {
    fileHandle = null;
    return handleStore('readwrite', function (store) { return store.delete('workbook'); })
      .catch(function () {});
  }

  function restoreHandle() {
    return handleStore('readonly', function (store) { return store.get('workbook'); })
      .then(function (handle) { fileHandle = handle || null; return fileHandle; })
      .catch(function () { return null; });
  }

  function canPickHandles() {
    return typeof window.showOpenFilePicker === 'function';
  }

  /** Whether the handle we hold is still allowed to be written to. */
  function handleWritable(prompt) {
    if (!fileHandle || !fileHandle.queryPermission) return Promise.resolve(false);
    return fileHandle.queryPermission({ mode: 'readwrite' }).then(function (state) {
      if (state === 'granted') return true;
      if (!prompt || !fileHandle.requestPermission) return false;
      return fileHandle.requestPermission({ mode: 'readwrite' }).then(function (asked) {
        return asked === 'granted';
      });
    }).catch(function () { return false; });
  }

  /* -------------------------------------------------------- catalog parsing */

  function optionList(select, values, selected) {
    select.innerHTML = '';
    values.forEach(function (v) {
      var option = document.createElement('option');
      option.value = v.value;
      option.textContent = v.label;
      if (v.value === selected) option.selected = true;
      select.appendChild(option);
    });
  }

  /**
   * Sets a select, adding the value as an option first when it is not already
   * there. A column the sheet has no header for is not in the option list, and
   * assigning a missing value silently snaps the select back to "Not used".
   */
  function setSelectValue(select, value, label) {
    if (value && !Array.prototype.some.call(select.options, function (o) { return o.value === value; })) {
      var option = document.createElement('option');
      option.value = value;
      option.textContent = label || value;
      select.appendChild(option);
    }
    select.value = value || '';
  }

  function pickSheet(wb) {
    var match = wb.findSheet('Logros');
    return match ? match.name : wb.sheets[0].name;
  }

  function renderSheetChoices(wb, chosen) {
    optionList($('sheetSelect'), wb.sheets.map(function (s) {
      return { value: s.name, label: s.name.trim() || '(unnamed)' };
    }), chosen);
  }

  function renderMappingChoices(header, chosen, newColumn) {
    var choices = [{ value: '', label: 'Not used' }].concat(header.labels.map(function (l) {
      return { value: l.col, label: l.col + '  ' + U.truncate(l.label, 26) };
    }));
    optionList($('mapName'), choices, chosen.name);
    optionList($('mapCode'), choices, chosen.code);
    optionList($('mapPrice'), choices, chosen.price);

    // A column that does not exist yet has no header to list, so it is offered
    // explicitly and labelled as new. Every real header stays on offer too, so
    // the choice can still be overridden.
    optionList($('mapLink'), newColumn
      ? choices.concat([{ value: newColumn, label: newColumn + '  (new column "' + S.LINK_HEADER + '")' }])
      : choices, chosen.link);
  }

  function currentMapping() {
    return {
      name: $('mapName').value,
      code: $('mapCode').value,
      price: $('mapPrice').value,
      link: $('mapLink').value
    };
  }

  /** Turns the chosen sheet and column mapping into rows for the worker. */
  function buildRows(mapping) {
    var rows = [];
    var problems = 0;

    sheetData.rows.forEach(function (r) {
      if (r.r <= headerInfo.headerRow) return;

      var name = U.normText(mapping.name ? r.cells[mapping.name] : '');
      var code = U.normText(mapping.code ? r.cells[mapping.code] : '');
      var priceRaw = U.normText(mapping.price ? r.cells[mapping.price] : '');
      var link = U.normText(mapping.link ? r.cells[mapping.link] : '');
      if (!name && !code) return;

      var parsed = FYG.price.parse(priceRaw);
      if (!parsed.ok && !link) problems++;

      rows.push({
        sheetRow: r.r,
        name: name,
        code: code,
        priceRaw: priceRaw,
        priceText: parsed.ok ? parsed.text : '',
        link: link,
        // A row the automation cannot possibly complete is flagged now rather
        // than failing halfway through the run.
        blocked: !link && (!name || !code || !parsed.ok),
        blockedReason: !name ? 'This row has no service name.'
          : !code ? 'This row has no code.'
          : !parsed.ok ? 'The price "' + priceRaw + '" could not be read.'
          : ''
      });
    });

    return { rows: rows, problems: problems };
  }

  /**
   * Reloading the catalog resets the run, which would throw away links that
   * have already been captured. Ask before doing that, and never silently.
   */
  function confirmDiscardProgress(what) {
    var captured = view.rows.filter(function (r) {
      return r.link && r.status === S.ROW.DONE;
    }).length;
    if (!captured) return true;
    return window.confirm(what + ' reloads the catalog and clears the ' + captured +
      ' link' + (captured === 1 ? '' : 's') + ' captured so far.\n\n' +
      'Export them first if you still need them. Continue?');
  }

  /** Puts the sheet and column dropdowns back to what the run is actually using. */
  function restoreMappingSelects() {
    var file = view.run.file;
    if (!file) return;
    if (file.sheetName) $('sheetSelect').value = file.sheetName;
    var mapping = file.mapping || {};
    [['mapName', 'name'], ['mapCode', 'code'], ['mapPrice', 'price'], ['mapLink', 'link']]
      .forEach(function (pair) { setSelectValue($(pair[0]), mapping[pair[1]] || ''); });
  }

  function applyMapping() {
    if (!sheetData || !headerInfo) return;
    var mapping = currentMapping();
    var built = buildRows(mapping);

    var withLink = built.rows.filter(function (r) { return r.link; }).length;
    var blocked = built.rows.filter(function (r) { return r.blocked; }).length;

    var summary = built.rows.length + ' rows found. ' +
      (built.rows.length - withLink - blocked) + ' to process, ' +
      withLink + ' already have a link';
    if (blocked) summary += ', ' + blocked + ' cannot be processed';
    if (proposed.link) {
      summary += '. Links will go into a new column ' + proposed.link;
    }
    $('fileSummary').textContent = summary + '.';

    return send(S.MSG.LOAD_CATALOG, {
      file: {
        name: fileName,
        sheetName: $('sheetSelect').value,
        mapping: mapping,
        headerRow: headerInfo.headerRow,
        // Only set when the export has to write the header itself, which is
        // exactly when the sheet did not already have that column.
        linkHeader: mapping.link && mapping.link === proposed.link ? S.LINK_HEADER : '',
        noteColumn: noteColumn,
        noteHeader: proposed.note && proposed.note === noteColumn ? S.NOTE_HEADER : ''
      },
      rows: built.rows
    }).then(refresh);
  }

  function loadSheet(name) {
    sheetData = workbook.readSheet(name);
    headerInfo = X.readHeader(sheetData);

    var detected = {
      name: X.findColumn(headerInfo, S.HEADERS.name),
      code: X.findColumn(headerInfo, S.HEADERS.code),
      price: X.findColumn(headerInfo, S.HEADERS.price),
      link: X.findColumn(headerInfo, S.HEADERS.link)
    };

    /*
     * This catalog has no link column at all, and its last column holds the
     * product images, so one is proposed just past everything the sheet uses.
     * On the next run the header written by the export is found by findColumn
     * above and nothing is proposed, which is what lets a restart skip the rows
     * that are already done.
     */
    proposed = { link: '', note: '' };
    if (!detected.link) {
      detected.link = X.nextFreeColumn(headerInfo, sheetData);
      proposed.link = detected.link;
    }
    noteColumn = X.findColumn(headerInfo, S.HEADERS.note);
    if (!noteColumn) {
      noteColumn = X.colName(X.colIndex(detected.link) + 1);
      proposed.note = noteColumn;
    }

    renderMappingChoices(headerInfo, detected, proposed.link);
    setHidden($('fileEmpty'), true);
    setHidden($('fileLoaded'), false);
    setHidden($('btnChangeFile'), false);
    return applyMapping();
  }

  /**
   * Opens the file picker. When the browser supports file handles it is used in
   * preference to the plain input, because only a handle can be written back to
   * later. Everything still works without one, just without the in place option.
   */
  function choosePrimaryFile() {
    if (!canPickHandles()) {
      $('fileInput').click();
      return;
    }
    window.showOpenFilePicker({
      multiple: false,
      types: [{
        description: 'Excel workbook',
        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
      }]
    }).then(function (handles) {
      var handle = handles && handles[0];
      if (!handle) return null;
      return rememberHandle(handle).then(function () { return handle.getFile(); });
    }).then(function (file) {
      if (file) openWorkbook(file);
    }).catch(function (err) {
      // AbortError just means the picker was dismissed, so do nothing.
      if (err && err.name === 'AbortError') return;
      $('fileInput').click();
    });
  }

  function openWorkbook(file) {
    choosingFile = false;
    $('fileSummary').textContent = 'Reading ' + file.name + '...';
    setHidden($('fileEmpty'), true);
    setHidden($('fileLoaded'), false);

    return file.arrayBuffer()
      .then(function (buffer) {
        var bytes = new Uint8Array(buffer);
        return X.load(bytes).then(function (wb) {
          workbook = wb;
          fileName = file.name;
          renderSheetChoices(wb, pickSheet(wb));
          return storeWorkbook(file.name, bytes).then(function () {
            return loadSheet($('sheetSelect').value);
          });
        });
      })
      .catch(function (err) {
        choosingFile = true;
        setHidden($('fileEmpty'), false);
        setHidden($('fileLoaded'), true);
        $('fileSummary').textContent = '';
        window.alert('That file could not be read.\n\n' + (err && err.message ? err.message : err));
      });
  }

  /* -------------------------------------------------------------- rendering */

  var STATUS_TEXT = {
    idle: 'Idle',
    running: 'Running',
    paused: 'Paused',
    attention: 'Needs attention',
    done: 'Finished'
  };

  function renderStatus(run) {
    var pill = $('statusPill');
    pill.dataset.status = run.status;
    pill.textContent = STATUS_TEXT[run.status] || run.status;
  }

  function renderAttention(run) {
    var showing = run.status === S.STATUS.ATTENTION && run.pending;
    setHidden($('attentionBanner'), !showing);
    if (!showing) return;

    var p = run.pending;
    $('attentionTitle').textContent = p.sheetRow
      ? 'Row ' + p.sheetRow + ' stopped at "' + p.stepLabel + '"'
      : 'The run stopped at "' + p.stepLabel + '"';
    $('attentionMessage').textContent = p.message;
    $('attentionWhere').textContent = [p.code, p.url].filter(Boolean).join('  |  ');
  }

  function renderStats(run) {
    $('statTotal').textContent = run.stats.total;
    $('statDone').textContent = run.stats.done;
    $('statSkipped').textContent = run.stats.skipped;
    $('statFailed').textContent = run.stats.failed;
  }

  /**
   * Decides which half of the catalog card to show. A run can outlive the cached
   * workbook, so this is driven by the stored run rather than by whether the
   * file happens to be in memory. Showing "Choose your catalog" over a run in
   * progress would be alarming and wrong.
   */
  function renderCatalog(run, rows) {
    if (choosingFile) return;
    var loaded = !!(run.file || rows.length);
    setHidden($('fileEmpty'), loaded);
    setHidden($('fileLoaded'), !loaded);
    setHidden($('btnChangeFile'), !loaded);
    if (!loaded) return;

    var file = run.file || {};

    if (!workbook) {
      // Without the file the sheet and column pickers cannot offer real
      // choices, so show what the run is using and disable them.
      var mapping = file.mapping || {};
      optionList($('sheetSelect'), [{ value: file.sheetName || '', label: file.sheetName || 'Unknown sheet' }],
        file.sheetName || '');
      [['mapName', 'name'], ['mapCode', 'code'], ['mapPrice', 'price'], ['mapLink', 'link']]
        .forEach(function (pair) {
          var col = mapping[pair[1]] || '';
          optionList($(pair[0]), [{ value: col, label: col || 'Not used' }], col);
          $(pair[0]).disabled = true;
        });
      $('sheetSelect').disabled = true;
      $('fileSummary').textContent = 'Using ' + (file.name || 'the catalog') +
        '. Load the same file again to export an updated .xlsx.';
    } else if (!$('fileSummary').textContent) {
      $('fileSummary').textContent = 'Using ' + (file.name || fileName) + '.';
    }
  }

  function renderProgress(run, rows) {
    var finished = run.stats.done + run.stats.skipped + run.stats.failed;
    var total = run.stats.total;
    var percent = total ? Math.round((finished / total) * 100) : 0;
    $('progressBar').style.width = percent + '%';

    $('progressText').textContent = total
      ? finished + ' of ' + total + ' rows (' + percent + '%)'
      : 'Nothing loaded yet';

    var remaining = total - finished;
    if (run.status === S.STATUS.RUNNING && run.durations.length && remaining > 0) {
      var average = run.durations.reduce(function (a, b) { return a + b; }, 0) / run.durations.length;
      $('etaText').textContent = 'about ' + U.formatDuration(average * remaining) + ' left';
    } else if (run.status === S.STATUS.DONE) {
      $('etaText').textContent = 'complete';
    } else {
      $('etaText').textContent = '';
    }

    var row = (run.cursor >= 0 && rows[run.cursor]) ? rows[run.cursor] : null;
    var busy = run.status === S.STATUS.RUNNING || run.status === S.STATUS.PAUSED ||
      run.status === S.STATUS.ATTENTION;
    setHidden($('currentCard'), !(row && busy));

    if (row && busy) {
      $('currentCode').textContent = 'Row ' + row.sheetRow + '  ' + row.code;
      $('currentName').textContent = row.name;

      var dots = $('stepDots');
      dots.innerHTML = '';
      var at = S.STEP_ORDER.indexOf(run.step);
      S.STEP_ORDER.forEach(function (step, i) {
        var dot = document.createElement('span');
        dot.className = 'step' + (i < at ? ' done' : i === at ? ' now' : '');
        dot.title = S.STEP_LABEL[step];
        dots.appendChild(dot);
      });
      $('stepName').textContent = 'Step ' + (at + 1) + ' of 7 · ' + (S.STEP_LABEL[run.step] || run.step);
    }
  }

  function renderControls(run, rows) {
    var hasRows = rows.length > 0;
    var running = run.status === S.STATUS.RUNNING;
    var resumable = run.status === S.STATUS.PAUSED || run.status === S.STATUS.ATTENTION;

    $('btnStart').textContent = resumable ? 'Resume' : 'Start';
    $('btnStart').disabled = !hasRows || running;
    $('btnPause').disabled = !running;
    $('btnStop').disabled = run.status === S.STATUS.IDLE || run.status === S.STATUS.DONE;

    ['minDelay', 'maxDelay', 'stepTimeout', 'maxAttempts', 'zoomPercent', 'dryRun'].forEach(function (id) {
      $(id).disabled = running;
    });
    ['sheetSelect', 'mapName', 'mapCode', 'mapPrice', 'mapLink'].forEach(function (id) {
      // Without the workbook these hold a single placeholder option, so there is
      // nothing to choose between.
      $(id).disabled = running || !workbook;
    });

    var hint = 'Open Fygaro and sign in first, then press Start.';
    if (!hasRows) hint = 'Load your catalog file above to begin.';
    else if (run.settings.dryRun && !running) hint = 'Dry run is on. The first row will be filled in but not saved.';
    else if (running) hint = 'Running. You can keep using other tabs, but leave the Fygaro tab open.';
    else if (run.status === S.STATUS.DONE) hint = 'Finished. Export the updated file below.';
    else if (resumable) hint = 'Paused. Press Resume to carry on from row ' +
      (rows[run.cursor] ? rows[run.cursor].sheetRow : '?') + '.';
    $('runHint').textContent = hint;
  }

  function renderSettings(run) {
    // Only the field being typed into is left alone. Every other field is
    // written back from the run, so what is on screen is always what will
    // actually be used, including after a value was clamped or corrected.
    var focused = document.activeElement;
    var fields = [
      ['minDelay', run.settings.minDelayMs],
      ['maxDelay', run.settings.maxDelayMs],
      ['stepTimeout', run.settings.stepTimeoutMs],
      ['maxAttempts', run.settings.maxAttempts],
      ['zoomPercent', run.settings.zoomPercent]
    ];
    fields.forEach(function (pair) {
      var el = $(pair[0]);
      if (el === focused) return;
      if (String(el.value) !== String(pair[1])) el.value = pair[1];
    });
    if ($('dryRun') !== focused) $('dryRun').checked = !!run.settings.dryRun;
  }

  function matchesFilter(row, index, run) {
    if (filter !== 'all') {
      if (filter === 'pending' && row.status !== S.ROW.PENDING && row.status !== S.ROW.ACTIVE) return false;
      if (filter !== 'pending' && row.status !== filter) return false;
    }
    if (searchTerm) {
      var haystack = U.foldText(row.code + ' ' + row.name + ' ' + row.sheetRow);
      if (haystack.indexOf(searchTerm) === -1) return false;
    }
    return true;
  }

  var BADGE = { pending: 'Pending', active: 'Working', done: 'Done', skipped: 'Skipped', failed: 'Failed' };

  function renderResults(run, rows) {
    var list = $('results');
    list.innerHTML = '';

    var matched = rows.filter(function (r, i) { return matchesFilter(r, i, run); });
    if (!matched.length) {
      var empty = document.createElement('div');
      empty.className = 'empty';
      empty.textContent = rows.length ? 'No rows match this filter.' : 'No catalog loaded yet.';
      list.appendChild(empty);
      $('resultsMeta').textContent = '';
      return;
    }

    // Centre the window on the row being worked on, so a long run keeps the
    // interesting part in view without rendering all 699 rows.
    var start = 0;
    if (matched.length > RESULT_LIMIT && run.cursor >= 0) {
      var at = matched.findIndex(function (r) { return r.i === run.cursor; });
      if (at >= 0) start = Math.max(0, Math.min(at - 5, matched.length - RESULT_LIMIT));
    }
    var window_ = matched.slice(start, start + RESULT_LIMIT);

    window_.forEach(function (row) {
      var item = document.createElement('div');
      item.className = 'result ' + row.status + (row.i === run.cursor ? ' active' : '');

      var n = document.createElement('span');
      n.className = 'n';
      n.textContent = row.sheetRow;

      var who = document.createElement('span');
      who.className = 'who';
      var code = document.createElement('b');
      code.textContent = row.code || '(no code)';
      var name = document.createElement('small');
      name.textContent = row.error || row.name;
      name.title = row.error || row.name;
      who.appendChild(code);
      who.appendChild(name);

      // The status word is always shown, never colour alone, and the copy
      // action sits beside it rather than replacing it.
      var badge = document.createElement('span');
      badge.className = 'badge';
      var word = document.createElement('span');
      // A row Fygaro already had is worth telling apart from one the sheet had
      // already linked. Both are skipped, but only one of them means the product
      // exists in Fygaro without a link of ours.
      word.textContent = (row.status === S.ROW.SKIPPED && row.reason === S.SKIP.EXISTS)
        ? 'Exists'
        : (BADGE[row.status] || row.status);
      badge.appendChild(word);

      if (row.link) {
        var copy = document.createElement('button');
        copy.type = 'button';
        copy.className = 'subtle copy';
        copy.textContent = 'Copy';
        copy.title = row.link;
        copy.setAttribute('aria-label', 'Copy the link for ' + (row.code || 'this row'));
        copy.addEventListener('click', function () {
          navigator.clipboard.writeText(row.link).then(function () {
            copy.textContent = 'Copied';
            setTimeout(function () { copy.textContent = 'Copy'; }, 1200);
          });
        });
        badge.appendChild(copy);
      }

      item.appendChild(n);
      item.appendChild(who);
      item.appendChild(badge);
      list.appendChild(item);
    });

    $('resultsMeta').textContent = matched.length > RESULT_LIMIT
      ? 'Showing ' + (start + 1) + ' to ' + (start + window_.length) + ' of ' + matched.length + ' matching rows.'
      : matched.length + ' row' + (matched.length === 1 ? '' : 's') + '.';
  }

  function renderLog(log) {
    var box = $('logLines');
    var atBottom = box.scrollHeight - box.scrollTop - box.clientHeight < 40;
    box.innerHTML = '';

    log.slice(-200).forEach(function (entry) {
      var line = document.createElement('div');
      line.className = entry.level;
      var time = document.createElement('time');
      time.textContent = U.clockTime(entry.ts);
      var msg = document.createElement('span');
      msg.className = 'msg';
      msg.textContent = entry.message;
      line.appendChild(time);
      line.appendChild(msg);
      box.appendChild(line);
    });

    if (atBottom) box.scrollTop = box.scrollHeight;
  }

  function renderExport(rows) {
    var withLinks = rows.filter(function (r) { return r.link; }).length;
    var ready = withLinks > 0;
    $('btnExportCsv').disabled = !ready;
    $('btnCopyLinks').disabled = !ready;
    $('btnExportXlsx').disabled = !ready || !workbook;

    // Writing back in place needs a handle to the file, which only the file
    // picker can give. A dropped or input chosen file is bytes and nothing more.
    var canOverwrite = !!fileHandle;
    $('targetOriginal').disabled = !canOverwrite;
    if (!canOverwrite && $('targetOriginal').checked) $('targetCopy').checked = true;

    $('targetOriginalNote').textContent = canOverwrite
      ? 'Writes the links straight into ' + U.truncate(fileName, 44) + '. Close it in Excel first, or ' +
        'the write will fail.'
      : (canPickHandles()
        ? 'Not available for this file. Choose it again with the Choose your catalog button, rather than ' +
          'dropping it, so the browser can grant write access.'
        : 'Not available in this browser. Only saving into a copy is possible here.');

    $('btnExportXlsx').textContent = $('targetOriginal').checked
      ? 'Update the original file'
      : 'Download updated .xlsx';

    if (!ready) {
      $('exportHint').textContent = 'Nothing to save yet. Links appear here as each one is created.';
    } else if (!workbook) {
      $('exportHint').textContent = withLinks + ' links ready. Load the catalog file again to write an ' +
        'updated .xlsx, or use CSV which needs no file.';
    } else {
      $('exportHint').textContent = withLinks + ' link' + (withLinks === 1 ? '' : 's') + ' ready.' +
        (lastSaveMessage ? '  ' + lastSaveMessage : '');
    }
  }

  /** Says exactly what is stored, so Clear everything is never a surprise. */
  function renderReset(run, rows, log) {
    var parts = [];
    if (rows.length) parts.push(rows.length + ' rows');
    var links = rows.filter(function (r) { return r.link; }).length;
    if (links) parts.push(links + ' link' + (links === 1 ? '' : 's'));
    if (run.file && run.file.name) parts.push('the file ' + U.truncate(run.file.name, 34));
    if (log.length) parts.push(log.length + ' log line' + (log.length === 1 ? '' : 's'));

    $('resetSummary').textContent = parts.length
      ? 'Currently stored: ' + parts.join(', ') + '.'
      : 'Nothing is stored yet.';
    $('btnResetAll').disabled = !parts.length && run.status === S.STATUS.IDLE;
  }

  function render() {
    renderStatus(view.run);
    renderAttention(view.run);
    renderCatalog(view.run, view.rows);
    renderStats(view.run);
    renderProgress(view.run, view.rows);
    renderControls(view.run, view.rows);
    renderSettings(view.run);
    renderResults(view.run, view.rows);
    renderLog(view.log);
    renderExport(view.rows);
    renderReset(view.run, view.rows, view.log);
  }

  function refresh() {
    return S.read().then(function (data) {
      view = data;
      render();
      return data;
    });
  }

  /* ----------------------------------------------------------------- export */

  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 30000);
  }

  /** Writes the produced workbook straight back into the file on disk. */
  function saveOverOriginal(blob, updates) {
    return handleWritable(true).then(function (allowed) {
      if (!allowed) {
        throw new Error('Permission to write to the file was not granted. Choose the file again, or ' +
          'save into a copy instead.');
      }
      return fileHandle.createWritable();
    }).then(function (writable) {
      // createWritable buffers and only commits on close, so a failure part way
      // through does not leave a half written spreadsheet on disk.
      return writable.write(blob).then(function () { return writable.close(); });
    }).then(function () {
      note('success', 'Wrote ' + updates.length + ' links into ' + fileName + '.');
      // Confirmed in place rather than through a modal, so a long run does not
      // stop dead behind a dialog nobody is there to dismiss.
      lastSaveMessage = 'Updated ' + fileName + ' with ' + updates.length + ' link' +
        (updates.length === 1 ? '' : 's') + ' at ' + U.clockTime() + '.';
    });
  }

  function exportXlsx() {
    if (!workbook) return;
    var file = view.run.file || {};
    var column = file.mapping && file.mapping.link;
    var sheetName = file.sheetName || $('sheetSelect').value;
    var toOriginal = $('targetOriginal').checked && !!fileHandle;

    /*
     * No silent fallback to a literal column. The last column of this catalog
     * holds the product images, so guessing would write links underneath them.
     */
    if (!column) {
      window.alert('There is no column to write the links into.\n\n' +
        'Choose one under Link in the catalog card.');
      return;
    }

    var updates = view.rows
      .filter(function (r) { return r.link; })
      .map(function (r) { return { row: r.sheetRow, value: r.link }; });

    if (!updates.length) return;

    // A row that finished without a link says why, in the sheet rather than
    // only in the panel, so the reason is still there tomorrow.
    var notes = view.rows
      .filter(function (r) { return !r.link && r.error; })
      .map(function (r) { return { row: r.sheetRow, value: r.error }; });

    var batches = [{ column: column, updates: updates }];
    if (file.linkHeader) {
      batches[0].updates = [{ row: file.headerRow || 1, value: file.linkHeader }].concat(updates);
    }
    if (file.noteColumn && notes.length) {
      if (file.noteHeader) notes.unshift({ row: file.headerRow || 1, value: file.noteHeader });
      batches.push({ column: file.noteColumn, updates: notes });
    }

    if (toOriginal && !window.confirm('Write ' + updates.length + ' link' +
      (updates.length === 1 ? '' : 's') + ' into ' + fileName + '?\n\n' +
      'This replaces the file on your disk. Close it in Excel first. Everything else in the workbook is ' +
      'kept exactly as it is.')) return;

    $('btnExportXlsx').disabled = true;
    X.writeColumns(workbook, sheetName, batches)
      .then(function (blob) {
        if (toOriginal) return saveOverOriginal(blob, updates);
        download(blob, X.exportName(fileName || file.name));
        note('success', 'Saved ' + updates.length + ' links into a new copy of the workbook.');
        return null;
      })
      .catch(function (err) {
        var message = err && err.message ? err.message : String(err);
        note('error', 'The workbook could not be saved: ' + message);
        window.alert((toOriginal ? 'The original file could not be updated.'
          : 'The updated file could not be created.') + '\n\n' + message +
          (toOriginal ? '\n\nIf the file is open in Excel, close it and try again.' : ''));
      })
      .then(function () { renderExport(view.rows); });
  }

  function exportCsv() {
    var rows = [['Sheet row', 'Código', 'Servicios', 'Link']];
    view.rows.forEach(function (r) {
      if (r.link) rows.push([r.sheetRow, r.code, r.name, r.link]);
    });
    download(new Blob([U.toCsv(rows)], { type: 'text/csv;charset=utf-8' }),
      'fygaro-links-' + U.todayStamp() + '.csv');
  }

  function copyLinks() {
    var text = view.rows.filter(function (r) { return r.link; })
      .map(function (r) { return r.sheetRow + '\t' + r.code + '\t' + r.link; })
      .join('\n');
    navigator.clipboard.writeText(text).then(function () {
      var button = $('btnCopyLinks');
      var original = button.textContent;
      button.textContent = 'Copied';
      setTimeout(function () { button.textContent = original; }, 1200);
    });
  }

  /* ------------------------------------------------------------------ wiring */

  /**
   * Reads one numeric setting from its field.
   *
   * An empty or unreadable field falls back to the value the run is already
   * using, and only then to the documented default. It never falls back to a
   * literal written at the call site: doing that is what once let a blank zoom
   * field save 100 percent, which is not the default and meant the page was
   * told to zoom in rather than out.
   */
  function numberSetting(id, key, min, max) {
    var raw = String($(id).value).trim();
    var parsed = raw === '' ? NaN : parseInt(raw, 10);
    if (!isFinite(parsed)) {
      var current = view.run.settings[key];
      parsed = isFinite(current) ? current : S.DEFAULT_SETTINGS[key];
    }
    return U.clamp(parsed, min, max);
  }

  function pushSettings() {
    var min = numberSetting('minDelay', 'minDelayMs', 0, 60000);
    var max = Math.max(min, numberSetting('maxDelay', 'maxDelayMs', 0, 60000));

    return send(S.MSG.UPDATE_SETTINGS, {
      settings: {
        minDelayMs: min,
        maxDelayMs: max,
        stepTimeoutMs: numberSetting('stepTimeout', 'stepTimeoutMs', 5000, 120000),
        maxAttempts: numberSetting('maxAttempts', 'maxAttempts', 1, 5),
        zoomPercent: numberSetting('zoomPercent', 'zoomPercent', 25, 200),
        dryRun: $('dryRun').checked
      }
    // Write the corrected values back to the fields, so what is on screen is
    // always what the run will actually use.
    }).then(function (result) {
      return refresh().then(function () { return result; });
    });
  }

  function wire() {
    var dropzone = $('dropzone');
    var fileInput = $('fileInput');

    fileInput.addEventListener('change', function () {
      if (!fileInput.files || !fileInput.files[0]) return;
      if (!confirmDiscardProgress('Loading a different file')) { fileInput.value = ''; return; }
      openWorkbook(fileInput.files[0]);
    });

    ['dragenter', 'dragover'].forEach(function (type) {
      dropzone.addEventListener(type, function (e) {
        e.preventDefault();
        dropzone.classList.add('over');
      });
    });
    ['dragleave', 'drop'].forEach(function (type) {
      dropzone.addEventListener(type, function (e) {
        e.preventDefault();
        dropzone.classList.remove('over');
      });
    });

    dropzone.addEventListener('click', function () {
      if (!confirmDiscardProgress('Loading a different file')) return;
      choosePrimaryFile();
    });
    dropzone.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      e.preventDefault();
      dropzone.click();
    });

    dropzone.addEventListener('drop', function (e) {
      var item = e.dataTransfer && e.dataTransfer.items && e.dataTransfer.items[0];
      var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (!file) return;
      if (!confirmDiscardProgress('Loading a different file')) return;

      // A drop can carry a real file handle too, which keeps the option to
      // update the original file open. Without it only a copy can be saved.
      if (item && typeof item.getAsFileSystemHandle === 'function') {
        item.getAsFileSystemHandle()
          .then(function (handle) {
            return (handle && handle.kind === 'file') ? rememberHandle(handle) : null;
          })
          .catch(function () {})
          .then(function () { openWorkbook(file); });
        return;
      }
      openWorkbook(file);
    });

    $('btnChangeFile').addEventListener('click', function () {
      choosingFile = true;
      setHidden($('fileEmpty'), false);
      setHidden($('fileLoaded'), true);
      setHidden($('btnChangeFile'), true);
      fileInput.value = '';
    });

    $('sheetSelect').addEventListener('change', function () {
      if (!confirmDiscardProgress('Changing the sheet')) return restoreMappingSelects();
      loadSheet($('sheetSelect').value);
    });
    ['mapName', 'mapCode', 'mapPrice', 'mapLink'].forEach(function (id) {
      $(id).addEventListener('change', function () {
        if (!confirmDiscardProgress('Changing the column mapping')) return restoreMappingSelects();
        applyMapping();
      });
    });

    ['minDelay', 'maxDelay', 'stepTimeout', 'maxAttempts', 'zoomPercent'].forEach(function (id) {
      $(id).addEventListener('change', pushSettings);
    });
    $('dryRun').addEventListener('change', function () { pushSettings().then(refresh); });

    $('btnStart').addEventListener('click', function () {
      var resumable = view.run.status === S.STATUS.PAUSED || view.run.status === S.STATUS.ATTENTION;
      pushSettings()
        .then(function () { return send(resumable ? S.MSG.RESUME : S.MSG.START); })
        .then(function (result) {
          if (result && result.ok === false && result.error) window.alert(result.error);
          return refresh();
        });
    });

    $('btnPause').addEventListener('click', function () { send(S.MSG.PAUSE).then(refresh); });
    $('btnStop').addEventListener('click', function () { send(S.MSG.STOP).then(refresh); });
    $('btnStopFromBanner').addEventListener('click', function () { send(S.MSG.STOP).then(refresh); });
    $('btnRetry').addEventListener('click', function () { send(S.MSG.RETRY).then(refresh); });
    $('btnSkip').addEventListener('click', function () { send(S.MSG.SKIP_ROW).then(refresh); });
    $('btnOpenFygaro').addEventListener('click', function () { send(S.MSG.OPEN_FYGARO); });

    $('btnResetAll').addEventListener('click', function () {
      var links = view.rows.filter(function (r) { return r.link; }).length;
      var running = view.run.status === S.STATUS.RUNNING;

      var warning = 'Clear everything?\n\n' +
        'This removes the loaded catalog, all ' + links + ' captured link' + (links === 1 ? '' : 's') +
        ', the activity log, the cached copy of your file, and your settings.\n\n' +
        (running ? 'The run in progress will be stopped.\n\n' : '') +
        'Your spreadsheet on disk is not touched. This cannot be undone.';
      if (!window.confirm(warning)) return;

      // Everything this panel holds is dropped straight away, before any async
      // work. Waiting would let the re-render triggered by the cleared storage
      // arrive first and leave stale interface state behind.
      workbook = null;
      fileName = '';
      sheetData = null;
      headerInfo = null;
      choosingFile = false;
      filter = 'all';
      searchTerm = '';
      lastSaveMessage = '';

      $('fileInput').value = '';
      $('search').value = '';
      $('fileSummary').textContent = '';
      $('sheetSelect').innerHTML = '';
      ['mapName', 'mapCode', 'mapPrice', 'mapLink'].forEach(function (id) { $(id).innerHTML = ''; });
      Array.prototype.forEach.call($('filters').children, function (c) {
        c.setAttribute('aria-pressed', String(c.dataset.filter === 'all'));
      });
      $('targetCopy').checked = true;
      $('logBox').open = false;

      setHidden($('fileEmpty'), false);
      setHidden($('fileLoaded'), true);
      setHidden($('btnChangeFile'), true);

      // Forgetting the handle is a convenience and must not hold up the clear.
      forgetHandle();
      send(S.MSG.RESET).then(refresh);
    });

    ['targetCopy', 'targetOriginal'].forEach(function (id) {
      $(id).addEventListener('change', function () { renderExport(view.rows); });
    });

    $('filters').addEventListener('click', function (e) {
      var chip = e.target.closest('.chip');
      if (!chip) return;
      filter = chip.dataset.filter;
      Array.prototype.forEach.call($('filters').children, function (c) {
        c.setAttribute('aria-pressed', String(c === chip));
      });
      renderResults(view.run, view.rows);
    });

    $('search').addEventListener('input', function () {
      searchTerm = U.foldText($('search').value);
      renderResults(view.run, view.rows);
    });

    $('btnExportXlsx').addEventListener('click', exportXlsx);
    $('btnExportCsv').addEventListener('click', exportCsv);
    $('btnCopyLinks').addEventListener('click', copyLinks);

    $('btnCopyLog').addEventListener('click', function () {
      var text = view.log.map(function (e) {
        return U.clockTime(e.ts) + '  [' + e.level + ']  ' + e.message;
      }).join('\n');
      navigator.clipboard.writeText(text);
    });

    chrome.runtime.onMessage.addListener(function (msg) {
      if (msg && msg.type === S.MSG.STATE_CHANGED) refresh();
    });
  }

  /* -------------------------------------------------------------------- boot */

  wire();
  // The file handle comes back first, so the panel knows straight away whether
  // updating the original file is possible for the workbook it is about to load.
  restoreHandle()
    .then(restoreWorkbook)
    .then(function (wb) {
      if (!wb) return null;
      // Restore the sheet and mapping the run was started with.
      return S.read().then(function (data) {
        var file = data.run.file;
        renderSheetChoices(wb, file ? file.sheetName : pickSheet(wb));
        if (file) {
          sheetData = wb.readSheet(file.sheetName);
          headerInfo = X.readHeader(sheetData);
          renderMappingChoices(headerInfo, file.mapping || {});
          fileName = file.name || fileName;
          setHidden($('fileEmpty'), true);
          setHidden($('fileLoaded'), false);
          setHidden($('btnChangeFile'), false);
          $('fileSummary').textContent = 'Loaded from ' + fileName + '.';
        }
        return wb;
      });
    })
    .catch(function () { return null; })
    .then(refresh);
})();
