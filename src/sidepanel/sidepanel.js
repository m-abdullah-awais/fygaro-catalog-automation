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

  /* Confirmation of the last successful save, shown in the Saving card. */
  var lastSaveMessage = '';

  /*
   * Where saves go, when the browser gave us somewhere to write.
   *
   * For "update the original" this is the catalog handle itself. For "save into
   * a new file" it is a file the user picked once, which is what lets the sheet
   * be kept up to date without asking again. Downloading a copy has no handle,
   * so it can only happen at the end.
   */
  var outputHandle = null;

  /* How many links were in the sheet the last time it was written. */
  var savedLinkCount = -1;
  var saving = false;

  /* The run status at the previous render, so leaving a run can be noticed. */
  var lastStatus = '';

  /*
   * How often the sheet is rewritten during a run.
   *
   * Five links is about four minutes of work at the pace this runs at, so that
   * is the most a crash can cost. Writing this workbook takes roughly three
   * seconds, which over a full catalog adds up to about twenty minutes against
   * a run measured in tens of hours: worth paying to never lose more than a
   * handful of links.
   */
  var AUTOSAVE_EVERY = 5;

  /* Columns this sheet does not have and the export therefore has to create,
   * along with the headers to write above them. Empty when the sheet already
   * had a link column, which is the case on every run after the first. */
  var proposed = { link: '', note: '' };

  /* Where the reason goes for a row that finished without a link. */
  var noteColumn = '';

  /* Row to picture map for the loaded sheet, from FYG.xlsx.readImages. */
  var imageIndex = null;

  /*
   * Bumped whenever the rules for sizing a picture change. Pictures are stored
   * once when the catalog is picked and then reused for every later run, so
   * without this a catalog loaded before a rule changed would keep uploading
   * the pictures it stored under the old one for ever.
   */
  var IMAGE_FIT_VERSION = 2;

  var view = { run: S.defaultRun(), rows: [], log: [] };
  var filter = 'all';
  var searchTerm = '';
  /* Set while the user is deliberately picking a different file, so a re-render
   * triggered by the running job cannot snap the picker shut under them. */
  var choosingFile = false;

  /* Set between pressing Clear everything and the storage actually emptying. */
  var clearing = false;

  var $ = function (id) { return document.getElementById(id); };

  function send(type, payload) {
    return chrome.runtime.sendMessage(Object.assign({ type: type }, payload || {}))
      .catch(function () { return null; });
  }

  function setHidden(el, hidden) {
    el.classList.toggle('hidden', !!hidden);
  }

  /* ------------------------------------------------------- workbook storage */

  /*
   * The workbook goes to IndexedDB, not chrome.storage.
   *
   * chrome.storage.local is capped at 10 MB and this catalog is 63.5 MB, which
   * base64 would inflate to about 85 MB. The old code did exactly that and its
   * failure was only a warning, so the panel silently lost the ability to export
   * an xlsx the first time it was closed.
   */
  function storeWorkbook(name, bytes) {
    // A save hands back a Blob and a file pick hands back the raw bytes. Taking
    // either avoids reading a 64 MB Blob into an ArrayBuffer only to wrap it in
    // a Blob again.
    var blob = bytes instanceof Blob ? bytes : new Blob([bytes]);
    return FYG.idb.run('workbook', 'readwrite', function (store) {
      store.put(blob, 'bytes');
      return store.put({ name: name, size: blob.size, savedAt: Date.now() }, 'meta');
    }).catch(function (err) {
      // Not fatal: the run still works, only the xlsx export needs the file back.
      note('warn', 'The catalog file could not be cached: ' + err.message);
    });
  }

  /*
   * The file on disk, when this panel is still allowed to read it.
   *
   * The cached copy is a snapshot taken when the file was picked, and every save
   * since has moved the file on without it. Trusting the snapshot is how a row
   * whose link was written into the sheet hours ago still read as unfinished:
   * it was handed to Fygaro, refused for a code already in use, and the run went
   * off to the payment links to find a link the sheet was already holding.
   *
   * Nothing is asked for here. Permission has to be requested under a live
   * click and a panel opening is not one, so a handle that has gone quiet falls
   * back to the cache rather than throwing a prompt at nobody.
   */
  function readWorkbookFromFile() {
    if (!fileHandle || typeof fileHandle.getFile !== 'function' || !fileHandle.queryPermission) {
      return Promise.resolve(null);
    }
    var file = null;
    return fileHandle.queryPermission({ mode: 'read' })
      .then(function (state) { return state === 'granted' ? fileHandle.getFile() : null; })
      .then(function (found) {
        file = found;
        return file ? file.arrayBuffer() : null;
      })
      .then(function (buffer) {
        if (!buffer) return null;
        var bytes = new Uint8Array(buffer);
        return X.load(bytes).then(function (wb) {
          workbook = wb;
          fileName = file.name || fileName;
          return refreshCacheFrom(file, bytes).then(function () { return wb; });
        });
      })
      .catch(function () { return null; });
  }

  /*
   * Writes the file into the cache, but only when the cache is actually behind.
   *
   * The cache is what keeps exporting alive once the handle is gone, so it is
   * worth keeping current. It is also 66 MB, and rewriting that every time the
   * panel is opened would be a real cost for nothing.
   */
  function refreshCacheFrom(file, bytes) {
    return FYG.idb.run('workbook', 'readonly', function (store) {
      return store.get('meta');
    }).then(function (meta) {
      var current = meta && meta.name === file.name && meta.size === file.size &&
        meta.savedAt >= file.lastModified;
      return current ? null : storeWorkbook(file.name, bytes);
    }).catch(function () { return null; });
  }

  function restoreWorkbook() {
    return readWorkbookFromFile().then(function (fromFile) {
      return fromFile || restoreCachedWorkbook();
    });
  }

  function restoreCachedWorkbook() {
    var meta = null;
    return FYG.idb.run('workbook', 'readonly', function (store) {
      return store.get('meta');
    }).then(function (found) {
      meta = found;
      if (!meta) return null;
      return FYG.idb.run('workbook', 'readonly', function (store) {
        return store.get('bytes');
      });
    }).then(function (blob) {
      if (!blob) return null;
      return blob.arrayBuffer();
    }).then(function (buffer) {
      if (!buffer) return null;
      return X.load(new Uint8Array(buffer)).then(function (wb) {
        workbook = wb;
        fileName = meta.name;
        return wb;
      });
    }).catch(function () { return null; });
  }

  /**
   * Writes the pictures, one transaction each.
   *
   * The store is emptied first. Picture ids are part paths such as
   * xl/media/image8.png, which every workbook has, so a stale picture from a
   * previously loaded catalog could otherwise be attached to the wrong product.
   * Given the whole point of this extension is not creating the wrong record,
   * that is worth a deliberate wipe.
   */
  function storeImages(found) {
    var images = [];
    found.images.forEach(function (img) { images.push(img); });
    if (!images.length) return Promise.resolve(0);

    var reduced = [];

    return FYG.idb.clearStore('images').then(function () {
      return images.reduce(function (chain, img, at) {
        return chain.then(function () {
          $('fileSummary').textContent = 'Preparing image ' + (at + 1) + ' of ' + images.length + '...';

          // Sized here, once per picture, rather than once per row. The same
          // photo is used by as many as 80 rows, so discovering Fygaro's limit
          // during a run measured in tens of hours would waste most of them.
          return FYG.imagefit.fit(img.bytes, img.type, img.name).then(function (fitted) {
            if (fitted.changed || fitted.note) {
              reduced.push({ name: img.name, from: fitted.from, to: fitted.to, note: fitted.note });
            }
            return FYG.idb.run('images', 'readwrite', function (store) {
              return store.put({
                id: img.id, name: fitted.name, type: fitted.type, size: fitted.bytes.length,
                blob: new Blob([fitted.bytes], { type: fitted.type })
              }, img.id);
            });
          });
        });
      }, Promise.resolve());
    }).then(function () {
      return FYG.idb.run('workbook', 'readwrite', function (store) {
        return store.put({
          version: IMAGE_FIT_VERSION,
          target: FYG.imagefit.TARGET_BYTES,
          count: images.length
        }, 'imagesMeta');
      });
    }).then(function () {
      reduced.forEach(function (r) {
        note(r.note ? 'warn' : 'info', 'Picture ' + r.name + ' was ' + Math.round(r.from / 1000) +
          ' KB, larger than Fygaro accepts, so it was re-encoded to ' + Math.round(r.to / 1000) +
          ' KB at its original size. ' + r.note);
      });
      return images.length;
    }).catch(function (err) {
      note('warn', 'The product images could not be stored: ' + err.message +
        ' The run will create products without them.');
      return 0;
    });
  }

  function note(level, message) {
    send(S.MSG.CONTENT_LOG, { level: level, message: message });
  }

  /* ---------------------------------------------------------- file handles */

  function rememberHandle(handle) {
    fileHandle = handle;
    if (!handle) return Promise.resolve();
    return FYG.idb.run('handles', 'readwrite', function (store) {
      return store.put(handle, 'workbook');
    }).catch(function () { /* the run still works, only in place saving needs it */ });
  }

  function forgetHandle() {
    fileHandle = null;
    return FYG.idb.run('handles', 'readwrite', function (store) {
      return store.delete('workbook');
    }).catch(function () {});
  }

  function restoreHandle() {
    return FYG.idb.run('handles', 'readonly', function (store) {
      return store.get('workbook');
    }).then(function (handle) {
      // Only adopt the remembered handle if none has been chosen in the
      // meantime. Reading it back takes a moment, and a user who picks a file
      // straight away must not have it quietly replaced by the previous one.
      if (!fileHandle) fileHandle = handle || null;
      return fileHandle;
    }).catch(function () { return null; });
  }

  function canPickHandles() {
    return typeof window.showOpenFilePicker === 'function';
  }

  /**
   * Whether a handle may be written to, optionally asking for permission.
   *
   * Asking MUST happen while a click is still being handled. Chrome only allows
   * requestPermission during a live user activation, which lasts a few seconds,
   * and this workbook takes longer than that to rebuild. Asking after the file
   * had been prepared is exactly why "Permission to write to the file was not
   * granted" appeared on a file the user had just chosen.
   */
  function canWriteTo(handle, ask) {
    if (!handle || !handle.queryPermission) return Promise.resolve(false);
    return handle.queryPermission({ mode: 'readwrite' }).then(function (state) {
      if (state === 'granted') return true;
      if (!ask || !handle.requestPermission) return false;
      return handle.requestPermission({ mode: 'readwrite' }).then(function (asked) {
        return asked === 'granted';
      });
    }).catch(function () { return false; });
  }

  function rememberOutputHandle(handle) {
    outputHandle = handle;
    if (!handle) {
      return FYG.idb.run('handles', 'readwrite', function (store) {
        return store.delete('output');
      }).catch(function () {});
    }
    return FYG.idb.run('handles', 'readwrite', function (store) {
      return store.put(handle, 'output');
    }).catch(function () {});
  }

  function restoreOutputHandle() {
    return FYG.idb.run('handles', 'readonly', function (store) {
      return store.get('output');
    }).then(function (handle) {
      if (!outputHandle) outputHandle = handle || null;
      return outputHandle;
    }).catch(function () { return null; });
  }

  /** The destination the user picked, and the handle it writes through. */
  function destination() {
    if ($('targetOriginal').checked) return { mode: 'original', handle: fileHandle };
    if ($('targetNewFile').checked) return { mode: 'newfile', handle: outputHandle };
    return { mode: 'copy', handle: null };
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

  /**
   * Reads the pictures for a sheet, and treats any failure as "this sheet has
   * none". A drawing this reader cannot follow must not stop the catalog from
   * loading, because the products still matter more than their photos.
   */
  /**
   * Re-prepares the stored pictures when they were written under older rules.
   *
   * A catalog is picked once and then reused for run after run, so a picture
   * stored before the size limit was understood would otherwise keep being
   * uploaded, and keep being refused, indefinitely.
   */
  function ensureImagesFitted(sheetName) {
    return FYG.idb.run('workbook', 'readonly', function (store) {
      return store.get('imagesMeta');
    }).then(function (meta) {
      if (meta && meta.version === IMAGE_FIT_VERSION && meta.target === FYG.imagefit.TARGET_BYTES) {
        return 0;
      }
      var found = readImagesSafely(sheetName);
      if (!found.images.size) return 0;
      note('info', 'Re-checking the ' + found.images.size + ' product pictures against the size ' +
        'Fygaro accepts, because they were prepared before that limit was known.');
      imageIndex = found;
      return storeImages(found);
    }).catch(function () { return 0; });
  }

  function readImagesSafely(name) {
    try {
      return X.readImages(workbook, name);
    } catch (err) {
      note('warn', 'The product images could not be read: ' + err.message);
      return { byRow: new Map(), images: new Map(), anchors: 0, skipped: [] };
    }
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

  /** The span of sheet rows the run should cover, clamped to what exists. */
  function currentRange() {
    var first = sheetData.rows.length ? sheetData.rows[0].r : 1;
    var last = sheetData.maxRow || first;
    sheetData.rows.forEach(function (r) {
      if (r.r > headerInfo.headerRow && r.r > last) last = r.r;
    });
    if (first <= headerInfo.headerRow) first = headerInfo.headerRow + 1;

    var from = parseInt($('rowFrom').value, 10);
    var to = parseInt($('rowTo').value, 10);
    return {
      from: isNaN(from) ? first : Math.max(first, from),
      to: isNaN(to) ? last : Math.min(last, to),
      first: first,
      last: last
    };
  }

  function currentMapping() {
    return {
      name: $('mapName').value,
      code: $('mapCode').value,
      price: $('mapPrice').value,
      link: $('mapLink').value
    };
  }

  /*
   * The first Fygaro link anywhere in a sheet row, or ''.
   *
   * Only consulted when the sheet arrived without a Link column of its own, in
   * which case the mapping points at a column that does not exist yet and every
   * cell under it reads as empty. A link put there by an earlier run under a
   * heading this build does not recognise, or pasted in by hand somewhere else,
   * would otherwise be invisible, and the row would be sent to Fygaro only to be
   * refused for a duplicate code tens of seconds later.
   *
   * It is deliberately not done when the sheet does have a Link column. There
   * the mapping is the answer, and a link quoted inside a service description
   * must not be allowed to skip a row that still needs doing.
   */
  function linkAnywhereIn(cells) {
    for (var col in cells) {
      if (!Object.prototype.hasOwnProperty.call(cells, col)) continue;
      var found = U.fygaroLink(cells[col]);
      if (found) return found;
    }
    return '';
  }

  /** Turns the chosen sheet and column mapping into rows for the worker. */
  function buildRows(mapping) {
    var rows = [];
    var problems = 0;
    var range = currentRange();
    var searchWholeRow = !!(proposed && proposed.link && mapping.link === proposed.link);

    sheetData.rows.forEach(function (r) {
      if (r.r <= headerInfo.headerRow) return;

      var name = U.normText(mapping.name ? r.cells[mapping.name] : '');
      var code = U.normText(mapping.code ? r.cells[mapping.code] : '');
      var priceRaw = U.normText(mapping.price ? r.cells[mapping.price] : '');
      if (!name && !code) return;

      /*
       * A row counts as done only when the cell really holds a Fygaro link.
       * Any non empty text used to count, which cuts both ways and both ways
       * are wrong: a note, a date or the word "pending" left in that column
       * silently retired a row that never got a link, and the Nota column
       * sitting one place over makes that an easy mapping slip to make.
       */
      var linkCell = U.normText(mapping.link ? r.cells[mapping.link] : '');
      var link = U.fygaroLink(linkCell);
      if (!link && searchWholeRow) link = linkAnywhereIn(r.cells);

      // Something is in the link column, but it is not a link. The row is
      // treated as still to do, and said out loud in the summary rather than
      // quietly retired or quietly retried.
      var linkJunk = !link && !!linkCell;

      var parsed = FYG.price.parse(priceRaw);
      if (!parsed.ok && !link) problems++;

      // Outside the chosen range is not a failure, it is simply not this
      // batch's work, so it reads as skipped rather than as something wrong.
      var outOfRange = !link && (r.r < range.from || r.r > range.to);

      // Fygaro is unlikely to accept a product priced at nothing, and finding
      // out row by row during a run that lasts hours is the worst way to learn.
      var freeOfCharge = parsed.ok && parsed.value === 0;

      var pictures = imageIndex ? imageIndex.byRow.get(r.r) : null;

      rows.push({
        sheetRow: r.r,
        // Only the first. Every one of the 1030 rows that has a picture has
        // exactly one, and keeping this a plain string keeps the row table and
        // every message that carries it small.
        imageId: (pictures && pictures[0]) || '',
        name: name,
        code: code,
        priceRaw: priceRaw,
        priceText: parsed.ok ? parsed.text : '',
        link: link,
        linkJunk: linkJunk,
        // A row the automation cannot possibly complete is flagged now rather
        // than failing halfway through the run.
        outOfRange: outOfRange,
        blocked: !link && !outOfRange && (!name || !code || !parsed.ok || freeOfCharge),
        blockedReason: outOfRange ? 'Fuera del rango elegido (' + range.from + ' a ' + range.to + ').'
          : !name ? 'This row has no service name.'
          : !code ? 'This row has no code.'
          : !parsed.ok ? 'The price "' + priceRaw + '" could not be read.'
          : freeOfCharge ? 'El precio es cero.'
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

  /** Puts the row range fields back to what the run is actually using. */
  function restoreRangeFields() {
    var file = view.run.file;
    if (!file || !file.range) return;
    $('rowFrom').value = file.range.from;
    $('rowTo').value = file.range.to;
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
    var chosen = currentRange();
    var built = buildRows(mapping);

    // Show the range that was actually used, so a bound left blank or typed
    // past the end of the sheet is corrected on screen rather than silently.
    $('rowFrom').value = chosen.from;
    $('rowTo').value = chosen.to;

    var withLink = built.rows.filter(function (r) { return r.link; }).length;
    var blocked = built.rows.filter(function (r) { return r.blocked; }).length;

    var outside = built.rows.filter(function (r) { return r.outOfRange; }).length;

    var junk = built.rows.filter(function (r) { return r.linkJunk; }).length;

    var summary = built.rows.length + ' rows found. ' +
      (built.rows.length - withLink - blocked - outside) + ' to process, ' +
      withLink + ' already have a link';
    if (blocked) summary += ', ' + blocked + ' cannot be processed';
    if (outside) summary += ', ' + outside + ' outside rows ' + chosen.from + ' to ' + chosen.to;
    // Worth a word on screen. These rows are about to be processed, and if that
    // column was meant to hold their links then processing them is wrong.
    if (junk) {
      summary += ', ' + junk + ' hold something in the link column that is not a Fygaro link and will be processed';
    }
    if (proposed.link) {
      summary += '. Links will go into a new column ' + proposed.link;
    }
    $('fileSummary').textContent = summary + '.';

    // One line about the pictures rather than a warning per row. Well over a
    // thousand rows have none, and saying so each time would bury the log.
    var withImages = built.rows.filter(function (r) { return r.imageId; }).length;
    if (imageIndex && imageIndex.images.size) {
      $('fileSummary').textContent += ' ' + withImages + ' row' + (withImages === 1 ? '' : 's') +
        ' have a picture, ' + (built.rows.length - withImages) + ' will be created without one.';
    }

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
        noteHeader: proposed.note && proposed.note === noteColumn ? S.NOTE_HEADER : '',
        range: { from: chosen.from, to: chosen.to }
      },
      rows: built.rows
    }).then(refresh);
  }

  /**
   * Reads a sheet and works out which column holds what.
   *
   * Every context that puts a sheet in front of the panel goes through here, so
   * none of them can quietly skip a piece of it. Restoring a session used to do
   * its own shorter version, which left the note column and the picture index
   * unset, and the next rebuild then created products with no photo and stopped
   * writing the reason a row was passed over.
   *
   * @param {string} name
   * @param {boolean} rereadImages false when only the worksheet can have moved,
   *   which is the case after a save. The drawings are untouched by one.
   * @returns {object} the columns the headers point at
   */
  function readSheetInto(name, rereadImages) {
    sheetData = workbook.readSheet(name);
    headerInfo = X.readHeader(sheetData);
    if (rereadImages || !imageIndex) imageIndex = readImagesSafely(name);

    var detected = {
      name: X.findColumn(headerInfo, S.HEADERS.name),
      code: X.findColumn(headerInfo, S.HEADERS.code),
      price: X.findColumn(headerInfo, S.HEADERS.price),
      link: X.findColumn(headerInfo, S.HEADERS.link)
    };

    /*
     * This catalog arrived with no link column at all, and its last column holds
     * the product images, so one is proposed just past everything the sheet
     * uses. Once the export has written that header, findColumn above finds it
     * and nothing is proposed, which is what lets a restart skip the rows that
     * are already done.
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
    return detected;
  }

  /**
   * Reads the catalog again when the sheet has moved on without it.
   *
   * The rows are worked out once, when the catalog is loaded, and then live in
   * the worker across restarts. So links written into the sheet after that,
   * whether by this extension or by hand, are invisible until something makes
   * the panel rebuild, and until then those rows are handed to Fygaro, refused
   * for a code already in use, and sent looking for a link the sheet has.
   *
   * Only done between runs, and only when the rebuild cannot cost anything: it
   * reads the sheet, so a link captured but not yet written into the sheet would
   * be dropped by it. That is worth far more than the time this saves.
   *
   * @returns {Promise<boolean>} whether the catalog was rebuilt
   */
  function catchUpWithSheet(data, file) {
    var status = data.run.status;
    if (status !== S.STATUS.IDLE && status !== S.STATUS.DONE) return Promise.resolve(false);

    var column = currentMapping().link;
    if (!column || column !== (file.mapping || {}).link) return Promise.resolve(false);

    var inSheet = {};
    sheetData.rows.forEach(function (r) {
      var link = U.fygaroLink(r.cells[column]);
      if (link) inSheet[r.r] = link;
    });

    var unsaved = data.rows.filter(function (r) { return r.link && inSheet[r.sheetRow] !== r.link; });
    if (unsaved.length) {
      note('warn', unsaved.length + ' captured link' + (unsaved.length === 1 ? ' is' : 's are') +
        ' not in the spreadsheet yet, so the catalog was left as it was. Save it first, then load ' +
        'the file again.');
      return Promise.resolve(false);
    }

    var late = data.rows.filter(function (r) { return !r.link && inSheet[r.sheetRow]; });
    if (!late.length) return Promise.resolve(false);

    note('info', late.length + ' row' + (late.length === 1 ? '' : 's') +
      ' already had a link in the spreadsheet, so the catalog was read again and ' +
      (late.length === 1 ? 'it will be' : 'they will be') + ' skipped.');
    return Promise.resolve(applyMapping()).then(function () { return true; });
  }

  function loadSheet(name) {
    var detected = readSheetInto(name, true);

    renderMappingChoices(headerInfo, detected, proposed.link);
    setHidden($('fileEmpty'), true);
    setHidden($('fileLoaded'), false);
    setHidden($('btnChangeFile'), false);

    // Storing the pictures blocks the catalog being ready on purpose. Start
    // needs the first picture immediately, and this is a one time cost on a
    // load that already takes a few seconds.
    return storeImages(imageIndex).then(applyMapping);
  }

  /**
   * Re-reads the sheet out of the workbook, keeping the mapping already chosen.
   *
   * Used after a save has been folded back in, so the columns on screen describe
   * the file as it now stands. It deliberately does not call applyMapping: that
   * rebuilds the row table and hands it to the worker, which mid run would throw
   * away the progress this very save was protecting.
   *
   * The pictures are left alone. Only the worksheet part was rewritten, so the
   * drawings and the media behind them are exactly as they were read.
   */
  function rereadSheet(name) {
    var keep = currentMapping();
    // The save wrote the Link and Nota headings, so columns that were proposals
    // a moment ago are ordinary headers now and must stop being offered as new,
    // or the picker lists each of them twice.
    readSheetInto(name, false);
    renderMappingChoices(headerInfo, keep, proposed.link);
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
      // A branch step is not one of the seven, so numbering it would be a lie.
      $('stepName').textContent = (at === -1 ? '' : 'Step ' + (at + 1) + ' of 7 · ') +
        (S.STEP_LABEL[run.step] || run.step);
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
      // already linked, and both from one simply left for another batch. All
      // three are skipped, but only one of them means the product exists in
      // Fygaro without a link of ours.
      word.textContent = (row.status === S.ROW.SKIPPED && row.reason === S.SKIP.EXISTS) ? 'Exists'
        : (row.status === S.ROW.SKIPPED && row.reason === S.SKIP.HAD_LINK) ? 'Has link'
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

    var canPickOutput = typeof window.showSaveFilePicker === 'function';
    $('targetNewFile').disabled = !canPickOutput;
    if (!canPickOutput && $('targetNewFile').checked) $('targetCopy').checked = true;

    $('targetOriginalNote').textContent = canOverwrite
      ? 'Writes the links straight into ' + U.truncate(fileName, 44) + '. Close it in Excel first, or ' +
        'the write will fail.'
      : (canPickHandles()
        ? 'Not available for this file. Choose it again with the Choose your catalog button, rather than ' +
          'dropping it, so the browser can grant write access.'
        : 'Not available in this browser. Only downloading a copy is possible here.');

    $('targetNewFileNote').textContent = !canPickOutput
      ? 'Not available in this browser. Only downloading a copy is possible here.'
      : (outputHandle
        ? 'Keeping ' + U.truncate(outputHandle.name || 'the file you chose', 44) + ' up to date. Your ' +
          'original file is never touched.'
        : 'Pick where it goes once, and it is kept up to date from then on. Your original file is never ' +
          'touched.');

    var where = destination();
    setHidden($('btnChooseOutput'), where.mode !== 'newfile');
    $('btnChooseOutput').textContent = outputHandle ? 'Choose a different file' : 'Choose where to save';

    /*
     * This line reports the destination as it actually stands, so it must not
     * repeat what the chosen option already says above it. For a download there
     * is nothing left to add, so it says nothing.
     */
    $('saveTargetStatus').textContent = where.handle
      ? 'Saving into ' + U.truncate(where.handle.name || fileName, 40) + ', about every ' +
        AUTOSAVE_EVERY + ' links and whenever the run stops. Leave this panel open so it can.'
      : (where.mode === 'newfile' ? 'Pick a file above before you start.' : '');

    $('btnExportXlsx').textContent = where.mode === 'copy'
      ? 'Download updated .xlsx'
      : 'Save now';

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
    if (clearing) return;
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
      var previous = lastStatus;
      view = data;
      render();
      // Every state change is a chance to have captured more links, which is
      // exactly when the sheet is worth writing again.
      lastStatus = data.run.status;
      maybeAutosave(previous);
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

  /**
   * Puts a fresh copy of the content scripts into a tab.
   *
   * The file list comes from the manifest rather than being written out again
   * here, so the two cannot drift apart.
   */
  function injectContentScripts(tabId) {
    var declared = (chrome.runtime.getManifest().content_scripts || [])[0];
    if (!chrome.scripting || !declared || !declared.js) {
      return Promise.reject(new Error('This browser cannot inject the helper into the page.'));
    }
    return chrome.scripting.executeScript({ target: { tabId: tabId }, files: declared.js });
  }

  /**
   * Sends a message to the page, putting the helper there first if it is not.
   *
   * Reloading the extension leaves the copy already in an open tab orphaned: it
   * keeps running but its link back to the extension is dead, so a message finds
   * nothing listening and Chrome answers "Receiving end does not exist". A tab
   * that was open before the extension was installed never had one at all.
   * Either way the answer is the same, and it is not something to make someone
   * reload a tab over.
   */
  function tellFygaroTab(tabId, message) {
    return chrome.tabs.sendMessage(tabId, message).catch(function () {
      return injectContentScripts(tabId).then(function () {
        return chrome.tabs.sendMessage(tabId, message);
      });
    });
  }

  /**
   * Presses More Results on whichever Fygaro list is open, until it runs out.
   *
   * The tab is messaged directly rather than through the worker. Every other
   * message the worker handles is queued behind one lock and rewrites the run,
   * and this can take minutes on a list of thousands, which would stall the run
   * loop for no reason: nothing about loading a list touches the run.
   */
  function loadAllRows() {
    var button = $('btnLoadAll');
    if (button.disabled) return Promise.resolve();

    var wasSaying = button.textContent;
    button.disabled = true;
    button.textContent = 'Loading...';
    $('loadAllHint').textContent = 'Pressing More Results on the Fygaro tab. Leave it open.';

    return chrome.tabs.query({ url: [S.ORIGIN + '/en/app/*', S.ORIGIN + '/es/app/*'] })
      .then(function (tabs) {
        if (!tabs || !tabs.length) {
          throw new Error('No Fygaro tab is open. Press Open Fygaro first, then go to the products ' +
            'or payment links list.');
        }
        return tellFygaroTab(tabs[0].id, {
          type: S.MSG.LOAD_ALL,
          minDelayMs: view.run.settings.minDelayMs,
          maxDelayMs: view.run.settings.maxDelayMs
        });
      })
      .then(function (result) {
        if (!result) throw new Error('The Fygaro tab did not answer. Reload it and try again.');
        $('loadAllHint').textContent = 'Loaded ' + result.rows + ' ' + result.kind + ' after ' +
          result.clicks + ' click' + (result.clicks === 1 ? '' : 's') + ', ' + result.stopped + '.';
      })
      .catch(function (err) {
        var message = err && err.message ? err.message : String(err);
        // Chrome's own wording for this is "Receiving end does not exist",
        // which says nothing to anyone who has not written an extension.
        if (/Receiving end|Could not establish connection/i.test(message)) {
          message = 'the Fygaro tab could not be reached. Reload that tab and try again.';
        }
        $('loadAllHint').textContent = 'Could not load the list: ' + message;
      })
      .then(function () {
        button.disabled = false;
        button.textContent = wasSaying;
        return refresh();
      });
  }

  /** What the sheet needs written into it, or null when there is nothing yet. */
  function buildBatches() {
    var file = view.run.file || {};
    var column = file.mapping && file.mapping.link;
    if (!column) return null;

    var updates = view.rows
      .filter(function (r) { return r.link; })
      .map(function (r) { return { row: r.sheetRow, value: r.link }; });
    if (!updates.length) return null;

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

    return {
      batches: batches,
      links: updates.length,
      sheetName: file.sheetName || $('sheetSelect').value
    };
  }

  function writeThrough(handle, blob) {
    // createWritable buffers and only commits on close, so a failure part way
    // through does not leave a half written spreadsheet on disk.
    return handle.createWritable().then(function (writable) {
      return writable.write(blob).then(function () { return writable.close(); });
    });
  }

  /**
   * Writes the sheet to wherever the user chose.
   *
   * Permission is never requested from in here. Rebuilding this workbook takes
   * seconds, which is longer than Chrome keeps a click alive, so a prompt raised
   * at this point can only fail. That is exactly why "Permission to write to the
   * file was not granted" appeared on a file that had just been chosen. It is
   * asked for when the destination is picked and again when Start is pressed,
   * both of which are real clicks.
   *
   * @param {boolean} silent an autosave, so nothing may pop up in front of
   *   someone who is not watching
   */
  /**
   * Brings the panel's copy of the workbook up to what was just saved.
   *
   * Without this the panel holds the bytes it first read for as long as it is
   * open, and the cache it restores from holds them for far longer than that.
   * Both then disagree with the file on disk about the one column that decides
   * whether a row still needs doing. The row reads as having no link, so it is
   * sent to Fygaro, where the product already exists, and the run spends its
   * time rediscovering work it finished yesterday.
   *
   * Nothing here is allowed to fail the save. The links are on disk by the time
   * it runs, which was the point.
   */
  function adoptSave(plan, blob) {
    return Promise.resolve()
      .then(function () {
        X.adoptColumns(workbook, plan.sheetName, plan.batches);
        rereadSheet(plan.sheetName);
        return blob ? storeWorkbook(fileName, blob) : null;
      })
      .catch(function (err) {
        note('warn', 'The saved links could not be folded back into the loaded copy: ' +
          (err && err.message ? err.message : err) +
          ' Reload the catalog file before the next run so finished rows are recognised.');
      });
  }

  function saveWorkbook(silent) {
    if (saving || !workbook) return Promise.resolve(false);

    var plan = buildBatches();
    if (!plan) {
      if (!silent) {
        window.alert('There is nothing to save yet.\n\n' +
          'Links appear here as each one is created, and the column they go into is chosen in the ' +
          'catalog card.');
      }
      return Promise.resolve(false);
    }

    var where = destination();
    saving = true;
    $('btnExportXlsx').disabled = true;

    var written = null;

    return X.writeColumns(workbook, plan.sheetName, plan.batches)
      .then(function (blob) {
        written = blob;
        if (!where.handle) {
          download(blob, X.exportName(fileName || 'catalogo.xlsx'));
          note('success', 'Saved ' + plan.links + ' links into a downloaded copy of the workbook.');
          return true;
        }
        return canWriteTo(where.handle, false).then(function (allowed) {
          if (!allowed) {
            throw new Error('Permission to write to that file is no longer granted. Choose the ' +
              'destination again under Saving.');
          }
          return writeThrough(where.handle, blob);
        }).then(function () {
          note('success', 'Wrote ' + plan.links + ' links into ' + (where.handle.name || fileName) + '.');
          return true;
        });
      })
      .then(function (ok) {
        if (!ok) return false;
        savedLinkCount = plan.links;
        // Confirmed in place rather than through a modal, so a long run does not
        // stop dead behind a dialog nobody is there to dismiss.
        lastSaveMessage = 'Saved ' + plan.links + ' link' + (plan.links === 1 ? '' : 's') +
          ' at ' + U.clockTime() + '.';
        return adoptSave(plan, written).then(function () { return true; });
      })
      .catch(function (err) {
        var message = err && err.message ? err.message : String(err);
        note('error', 'The spreadsheet could not be saved: ' + message);
        lastSaveMessage = 'Could not save at ' + U.clockTime() + '. ' + message;
        if (!silent) {
          window.alert((where.mode === 'copy' ? 'The updated file could not be created.'
            : 'The spreadsheet could not be updated.') + '\n\n' + message +
            (where.mode === 'copy' ? '' : '\n\nIf the file is open in Excel, close it and try again.'));
        }
        return false;
      })
      .then(function (ok) {
        saving = false;
        renderExport(view.rows);
        return ok;
      });
  }

  /**
   * Saves without being asked, so a run that stops early keeps its work.
   *
   * Links are held in the panel from the moment they are captured, but nobody
   * watching a run of this length should have to know that, nor be the thing
   * standing between hours of work and a saved file.
   */
  function maybeAutosave(previous) {
    if (saving || !workbook) return;

    var where = destination();
    var links = view.rows.filter(function (r) { return r.link; }).length;
    if (!links || links === savedLinkCount) return;

    /*
     * A run that stops for any reason is a moment worth saving: finished,
     * stopped by hand, paused, or waiting for someone. Stopping is measured as a
     * transition rather than a state, so simply reopening the panel on an old
     * run does not rewrite the file, or worse, download another copy.
     */
    var stopped = previous === S.STATUS.RUNNING && view.run.status !== S.STATUS.RUNNING;

    // A download cannot happen every few rows without burying the user in
    // files, so that destination saves once, when the run stops.
    if (!where.handle) {
      if (stopped) saveWorkbook(true);
      return;
    }

    if (stopped || links - Math.max(savedLinkCount, 0) >= AUTOSAVE_EVERY) saveWorkbook(true);
  }

  /**
   * Confirms the chosen destination can actually be written to.
   *
   * Called straight out of a click, because a live click is the only time Chrome
   * will show the permission prompt.
   */
  function ensureDestinationReady() {
    var where = destination();
    if (where.mode === 'copy') return Promise.resolve(true);

    if (!where.handle) {
      window.alert('Choose where to save first.\n\n' +
        'Under Saving, pick the file to write into, or switch to downloading a copy at the end.');
      return Promise.resolve(false);
    }

    return canWriteTo(where.handle, true).then(function (allowed) {
      if (allowed) return true;
      window.alert('Permission to write to that file was not granted.\n\n' +
        'Choose the destination again under Saving, or switch to downloading a copy at the end.');
      return false;
    });
  }

  /** Picks a file to keep up to date, once, before the run starts. */
  function chooseOutputFile() {
    if (typeof window.showSaveFilePicker !== 'function') {
      window.alert('This browser cannot save into a file you choose.\n\n' +
        'Use "Download a copy at the end" instead.');
      return Promise.resolve(false);
    }

    return window.showSaveFilePicker({
      suggestedName: X.exportName(fileName || 'catalogo.xlsx'),
      types: [{
        description: 'Excel workbook',
        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
      }]
    }).then(function (handle) {
      return rememberOutputHandle(handle).then(function () {
        savedLinkCount = -1;
        note('info', 'The spreadsheet will be kept up to date in ' + handle.name + ' as the run goes along.');
        renderExport(view.rows);
        return true;
      });
    }).catch(function () {
      // Cancelling the picker is an ordinary thing to do, not a failure.
      renderExport(view.rows);
      return false;
    });
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

    // The range decides which rows are in scope, so changing it rebuilds them
    // behind the same guard as the mapping rather than silently discarding work.
    ['rowFrom', 'rowTo'].forEach(function (id) {
      $(id).addEventListener('change', function () {
        if (!confirmDiscardProgress('Changing the row range')) return restoreRangeFields();
        applyMapping();
      });
    });

    ['minDelay', 'maxDelay', 'stepTimeout', 'maxAttempts', 'zoomPercent'].forEach(function (id) {
      $(id).addEventListener('change', pushSettings);
    });
    $('dryRun').addEventListener('change', function () { pushSettings().then(refresh); });

    $('btnStart').addEventListener('click', function () {
      var resumable = view.run.status === S.STATUS.PAUSED || view.run.status === S.STATUS.ATTENTION;

      // Asked for here, in the click, because Chrome only shows the permission
      // prompt while one is still being handled. Finding out hours later that
      // the sheet cannot be written is the failure this exists to prevent.
      ensureDestinationReady().then(function (ready) {
        if (!ready) return null;
        return startRun(resumable);
      });
    });

    function startRun(resumable) {
      return pushSettings()
        .then(function () { return send(resumable ? S.MSG.RESUME : S.MSG.START); })
        .then(function (result) {
          if (result && result.ok === false && result.error) window.alert(result.error);
          return refresh();
        });
    }

    $('btnPause').addEventListener('click', function () { send(S.MSG.PAUSE).then(refresh); });
    $('btnStop').addEventListener('click', function () { send(S.MSG.STOP).then(refresh); });
    $('btnStopFromBanner').addEventListener('click', function () { send(S.MSG.STOP).then(refresh); });
    $('btnRetry').addEventListener('click', function () { send(S.MSG.RETRY).then(refresh); });
    $('btnSkip').addEventListener('click', function () { send(S.MSG.SKIP_ROW).then(refresh); });
    $('btnOpenFygaro').addEventListener('click', function () { send(S.MSG.OPEN_FYGARO); });

    $('btnLoadAll').addEventListener('click', loadAllRows);

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
      /*
       * The remembered state goes too, not just the controls. Storage is
       * cleared asynchronously, and any render that lands in the meantime reads
       * this: with the workbook already gone it would fill the sheet picker
       * back in from the run it still thought was loaded.
       */
      view = { run: S.defaultRun(), rows: [], log: [] };

      /*
       * Nothing may repaint until the storage clear has actually landed.
       * Emptying storage is asynchronous, so a read taken in the meantime still
       * holds the old catalog, and rendering that with the workbook already
       * gone fills the sheet picker back in from the run it just discarded.
       */
      clearing = true;
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
      // The workbook and its pictures are 130 MB between them and live outside
      // chrome.storage, so clearing that alone would leave them on disk while
      // telling the user everything had gone.
      forgetHandle();
      rememberOutputHandle(null);
      savedLinkCount = -1;
      FYG.idb.clearStore('workbook');
      FYG.idb.clearStore('images');
      imageIndex = null;
      send(S.MSG.RESET).then(function () {
        clearing = false;
        return refresh();
      });
    });

    ['targetCopy', 'targetOriginal', 'targetNewFile'].forEach(function (id) {
      $(id).addEventListener('change', function () {
        savedLinkCount = -1;
        renderExport(view.rows);

        // Both of these need a live click, which this is. Asking now means the
        // run cannot get hours in and then discover it has nowhere to write.
        if (id === 'targetNewFile' && !outputHandle) return chooseOutputFile();
        if (id === 'targetOriginal' && fileHandle) return canWriteTo(fileHandle, true);
        return null;
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

    $('btnExportXlsx').addEventListener('click', function () { saveWorkbook(false); });
    $('btnChooseOutput').addEventListener('click', chooseOutputFile);
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

  // Paint from chrome.storage before touching IndexedDB. Reading a 63.5 MB
  // workbook back takes a moment, and the panel should show the run it already
  // knows about rather than sitting blank until a database answers.
  refresh();

  // The file handle comes back first, so the panel knows straight away whether
  // updating the original file is possible for the workbook it is about to load.
  restoreHandle()
    .then(restoreOutputHandle)
    .then(restoreWorkbook)
    .then(function (wb) {
      if (!wb) return null;
      // Restore the sheet and mapping the run was started with.
      return S.read().then(function (data) {
        var file = data.run.file;
        renderSheetChoices(wb, file ? file.sheetName : pickSheet(wb));
        if (file) {
          readSheetInto(file.sheetName, false);
          renderMappingChoices(headerInfo, file.mapping || {}, proposed.link);
          fileName = file.name || fileName;
          setHidden($('fileEmpty'), true);
          setHidden($('fileLoaded'), false);
          setHidden($('btnChangeFile'), false);
          $('fileSummary').textContent = 'Loaded from ' + fileName + '.';
          return ensureImagesFitted(file.sheetName)
            .then(function () { return catchUpWithSheet(data, file); })
            .then(function () { return wb; });
        }
        return wb;
      });
    })
    .catch(function () { return null; })
    .then(refresh);
})();
