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

  function pickSheet(wb) {
    var match = wb.findSheet('Logros');
    return match ? match.name : wb.sheets[0].name;
  }

  function renderSheetChoices(wb, chosen) {
    optionList($('sheetSelect'), wb.sheets.map(function (s) {
      return { value: s.name, label: s.name.trim() || '(unnamed)' };
    }), chosen);
  }

  function renderMappingChoices(header, chosen) {
    var choices = [{ value: '', label: 'Not used' }].concat(header.labels.map(function (l) {
      return { value: l.col, label: l.col + '  ' + U.truncate(l.label, 26) };
    }));
    optionList($('mapName'), choices, chosen.name);
    optionList($('mapCode'), choices, chosen.code);
    optionList($('mapPrice'), choices, chosen.price);
    optionList($('mapLink'), choices, chosen.link);
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
      .forEach(function (pair) { $(pair[0]).value = mapping[pair[1]] || ''; });
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
    $('fileSummary').textContent = summary + '.';

    return send(S.MSG.LOAD_CATALOG, {
      file: { name: fileName, sheetName: $('sheetSelect').value, mapping: mapping },
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

    renderMappingChoices(headerInfo, detected);
    setHidden($('fileEmpty'), true);
    setHidden($('fileLoaded'), false);
    setHidden($('btnChangeFile'), false);
    return applyMapping();
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

    ['minDelay', 'maxDelay', 'stepTimeout', 'maxAttempts', 'dryRun'].forEach(function (id) {
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
    if (document.activeElement && document.activeElement.closest('.card') === $('minDelay').closest('.card')) {
      // Do not fight the user while they are typing in the settings card.
      return;
    }
    $('minDelay').value = run.settings.minDelayMs;
    $('maxDelay').value = run.settings.maxDelayMs;
    $('stepTimeout').value = run.settings.stepTimeoutMs;
    $('maxAttempts').value = run.settings.maxAttempts;
    $('dryRun').checked = !!run.settings.dryRun;
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
      word.textContent = BADGE[row.status] || row.status;
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

    if (!ready) {
      $('exportHint').textContent = 'Nothing to export yet. Links appear here as each one is created.';
    } else if (!workbook) {
      $('exportHint').textContent = withLinks + ' links ready. Load the catalog file again to export an ' +
        'updated .xlsx, or use CSV which needs no file.';
    } else {
      $('exportHint').textContent = withLinks + ' link' + (withLinks === 1 ? '' : 's') +
        ' ready. Your original file is never modified: this downloads a new copy.';
    }
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

  function exportXlsx() {
    if (!workbook) return;
    var file = view.run.file || {};
    var column = (file.mapping && file.mapping.link) || 'H';
    var sheetName = file.sheetName || $('sheetSelect').value;

    var updates = view.rows
      .filter(function (r) { return r.link; })
      .map(function (r) { return { row: r.sheetRow, value: r.link }; });

    if (!updates.length) return;

    $('btnExportXlsx').disabled = true;
    X.writeColumn(workbook, sheetName, column, updates)
      .then(function (blob) {
        download(blob, X.exportName(fileName || file.name));
        note('success', 'Exported ' + updates.length + ' links into a new copy of the workbook.');
      })
      .catch(function (err) {
        window.alert('The updated file could not be created.\n\n' + (err && err.message ? err.message : err));
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

  function pushSettings() {
    var min = Math.max(0, parseInt($('minDelay').value, 10) || 0);
    var max = Math.max(min, parseInt($('maxDelay').value, 10) || min);
    if (parseInt($('maxDelay').value, 10) < min) $('maxDelay').value = max;

    return send(S.MSG.UPDATE_SETTINGS, {
      settings: {
        minDelayMs: min,
        maxDelayMs: max,
        stepTimeoutMs: U.clamp(parseInt($('stepTimeout').value, 10) || 20000, 5000, 120000),
        maxAttempts: U.clamp(parseInt($('maxAttempts').value, 10) || 2, 1, 5),
        dryRun: $('dryRun').checked
      }
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
    dropzone.addEventListener('drop', function (e) {
      var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (!file) return;
      if (!confirmDiscardProgress('Loading a different file')) return;
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

    ['minDelay', 'maxDelay', 'stepTimeout', 'maxAttempts'].forEach(function (id) {
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
      if (!window.confirm('Clear the loaded catalog, all captured links and the log?\n\n' +
        'Export first if you still need the links. This cannot be undone.')) return;
      send(S.MSG.RESET).then(function () {
        workbook = null;
        fileName = '';
        sheetData = null;
        headerInfo = null;
        choosingFile = false;
        setHidden($('fileEmpty'), false);
        setHidden($('fileLoaded'), true);
        setHidden($('btnChangeFile'), true);
        $('fileInput').value = '';
        return refresh();
      });
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
  restoreWorkbook()
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
