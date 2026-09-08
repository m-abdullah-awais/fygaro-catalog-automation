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

  // The trimmed fixture, not the real 63.5 MB catalog: headless Chrome does not
  // finish fetching a file that size before it dumps the page.
  var WORKBOOK_PATH = 'fixtures/catalog-sample.xlsx';
  var WORKBOOK_NAME = 'catalog-sample.xlsx';
  var SHEET = 'Logros ';

  var results = [];
  var recorded = { badges: [], notifications: [], tabUpdates: [], downloads: [], actions: [],
                   alerts: [], confirms: [] };
  var currentZoom = 1;
  var zoomFailuresLeft = 0;

  /* Stands in for a real FileSystemFileHandle, so the in place save path is
   * exercised rather than mocked away. */
  var savedToDisk = null;
  var handlePermission = 'granted';
  var fakeHandle = null;

  // Dialogs are recorded instead of shown. A modal would stall headless Chrome.
  var confirmAnswer = true;
  window.alert = function (text) { recorded.alerts.push(String(text)); };
  window.confirm = function (text) { recorded.confirms.push(String(text)); return confirmAnswer; };

  function assert(condition, message) {
    if (!condition) throw new Error(message);
  }

  /**
   * A PNG of random pixels, which barely compresses, so a modest canvas gives a
   * genuinely large file to test the size limit against.
   */
  function makeNoisyPng(side) {
    var canvas = document.createElement('canvas');
    canvas.width = side;
    canvas.height = side;
    var ctx = canvas.getContext('2d');
    var picture = ctx.createImageData(side, side);
    for (var i = 0; i < picture.data.length; i += 4) {
      picture.data[i] = Math.random() * 256;
      picture.data[i + 1] = Math.random() * 256;
      picture.data[i + 2] = Math.random() * 256;
      picture.data[i + 3] = 255;
    }
    ctx.putImageData(picture, 0, 0);
    return new Promise(function (resolve) {
      canvas.toBlob(function (blob) {
        blob.arrayBuffer().then(function (buffer) { resolve(new Uint8Array(buffer)); });
      }, 'image/png');
    });
  }

  function check(name, fn) {
    // A check that never settles used to stall the whole page, which reported
    // nothing at all and said nothing about which one was stuck. It fails now.
    var stall = new Promise(function (resolve, reject) {
      setTimeout(function () { reject(new Error('this check never finished')); }, 8000);
    });
    return Promise.race([Promise.resolve().then(fn), stall])
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
          // Delivered on its own task, as Chrome does. Calling the listeners
          // straight from this executor meant any IndexedDB transaction a
          // handler opened was created inside the caller's task and never
          // became active, so the read never completed and the page hung.
          setTimeout(function () {
            var wantsAsync = false;
            messageListeners.forEach(function (listener) {
              var result = listener(clone(message), { tab: null }, settle);
              if (result === true) wantsAsync = true;
            });
            // No listener took the message, so resolve like Chrome would.
            if (!wantsAsync) settle(undefined);
          }, 0);
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
        },
        clear: function () {
          var changes = {};
          Object.keys(store).forEach(function (k) {
            changes[k] = { oldValue: clone(store[k]) };
            delete store[k];
          });
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
        // Lets a test make zooming fail the way a still loading tab does.
        if (zoomFailuresLeft > 0) {
          zoomFailuresLeft--;
          return Promise.reject(new Error('Cannot zoom the tab while it is loading'));
        }
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
      return waitUntil(function () { return $('statusPill') && $('btnStart').disabled; },
        'the panel to mount and render')
        .then(function () {
          assert($('statusPill').textContent === 'Idle', 'starts as "' + $('statusPill').textContent + '"');
          assert($('btnExportXlsx').disabled, 'Export should be disabled before any link exists');
          assert($('targetCopy').checked, 'saving into a copy should be the default');
          assert($('targetOriginal').disabled, 'updating the original needs a file handle');
          return 'panel mounted, controls disabled, copy is the default target';
        });
    })

    .then(function () {
      return check('choosing the catalog file loads every row and proposes a link column', function () {
        // Goes through the file picker, which is what hands over a handle the
        // extension can later write back to.
        $('dropzone').click();

        return waitUntil(function () { return $('statTotal').textContent === '30'; }, 'the rows to load', 20000)
          .then(function () {
            assert($('sheetSelect').value === SHEET, 'sheet is "' + $('sheetSelect').value + '"');
            assert($('mapName').value === 'D', 'Name mapped to ' + $('mapName').value);
            assert($('mapCode').value === 'C', 'Code mapped to ' + $('mapCode').value);
            assert($('mapPrice').value === 'E', 'Price mapped to ' + $('mapPrice').value);
            // The sheet has no Link column, so one is proposed past the images.
            assert($('mapLink').value === 'I', 'Link mapped to ' + $('mapLink').value);
            var chosen = $('mapLink').selectedOptions[0];
            assert(/new column/.test(chosen.textContent), 'the option should say it is new: ' + chosen.textContent);
            assert($('statSkipped').textContent === '0', 'expected no pre-skipped rows, saw ' + $('statSkipped').textContent);
            assert(!$('btnStart').disabled, 'Start should now be enabled');
            assert(!$('targetOriginal').disabled,
              'picking the file should make updating the original possible');
            return '30 rows, a new column I proposed, and the original file is writable';
          });
      });
    })

    .then(function () {
      return check('a blank settings field never invents a value different from the default', function () {
        // This is the bug that made the page zoom in instead of out: a blank
        // field fell back to a literal 100, which is not the default of 67, and
        // that wrong value was then saved.
        return readState().then(function (before) {
          assert(before.run.settings.zoomPercent === 67,
            'precondition: the default zoom should be 67, it is ' + before.run.settings.zoomPercent);

          $('zoomPercent').value = '';
          $('zoomPercent').dispatchEvent(new Event('change'));
          $('minDelay').value = '';
          $('minDelay').dispatchEvent(new Event('change'));

          return waitUntil(function () { return $('zoomPercent').value === '67'; },
            'the blank field to be written back');
        }).then(readState).then(function (after) {
          assert(after.run.settings.zoomPercent === 67,
            'a blank zoom field saved ' + after.run.settings.zoomPercent + ' instead of keeping 67');
          assert(after.run.settings.minDelayMs === 1200,
            'a blank min delay saved ' + after.run.settings.minDelayMs + ' instead of keeping 1200');
          assert($('minDelay').value === '1200', 'the field should show the value that will be used');
          return 'blank fields kept 67% and 1200ms, and the fields were corrected on screen';
        });
      });
    })

    .then(function () {
      return check('no row is pre skipped, because none carries a link yet', function () {
        return readState().then(function (data) {
          var skipped = data.rows.filter(function (r) { return r.status === S.ROW.SKIPPED; });
          assert(skipped.length === 0, skipped.length + ' rows were pre skipped, expected none');
          assert(data.rows[0].status === S.ROW.PENDING, 'row 2 should be pending');
          return 'every row is still to do';
        });
      });
    })

    .then(function () {
      return check('a row range leaves the rest of the sheet alone, then restores', function () {
        confirmAnswer = true;
        $('rowFrom').value = '4';
        $('rowTo').value = '6';
        $('rowTo').dispatchEvent(new Event('change'));

        return waitUntil(function () {
          return $('statSkipped').textContent !== '0';
        }, 'the rows outside the range to be set aside')
          .then(readState)
          .then(function (data) {
            var inRange = data.rows.filter(function (r) { return r.status === S.ROW.PENDING; });
            assert(inRange.length === 3, inRange.length + ' rows are still to do, expected 3');
            assert(inRange[0].sheetRow === 4, 'the first is sheet row ' + inRange[0].sheetRow);
            assert(inRange[2].sheetRow === 6, 'the last is sheet row ' + inRange[2].sheetRow);

            var outside = data.rows.filter(function (r) { return r.reason === S.SKIP.OUT_OF_RANGE; });
            assert(outside.length === 27, outside.length + ' rows were set aside, expected 27');
            assert(/Fuera del rango/.test(outside[0].error), 'unhelpful reason: ' + outside[0].error);
            // Not a failure: nothing is wrong with these rows.
            assert(data.run.stats.failed === 0, 'rows outside the range were counted as failures');

            // Put the whole sheet back so the rest of the suite runs as before.
            $('rowFrom').value = '';
            $('rowTo').value = '';
            $('rowTo').dispatchEvent(new Event('change'));
            return waitUntil(function () { return $('statSkipped').textContent === '0'; }, 'the full sheet');
          })
          .then(readState)
          .then(function (data) {
            assert(data.rows.filter(function (r) { return r.status === S.ROW.PENDING; }).length === 30,
              'clearing the range should put all 30 rows back');
            assert($('rowFrom').value === '2' && $('rowTo').value === '31',
              'the fields should show the range actually used, saw ' +
              $('rowFrom').value + ' to ' + $('rowTo').value);
            return '3 rows in range, 27 set aside, then all 30 back';
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
      return check('zooming retries when the tab is not ready, and recovers drift', function () {
        $('btnPause').click();
        return waitUntil(function () { return $('statusPill').dataset.status === 'paused'; }, 'the pause')
          .then(function () {
            return readState();
          })
          .then(function (before) {
            var logLines = before.log.length;
            // The page drifts back to 100%, as if the user pressed Ctrl and zero,
            // and the first two attempts to correct it fail.
            currentZoom = 1;
            zoomFailuresLeft = 2;

            $('btnStart').click();
            return waitUntil(function () { return Math.abs(currentZoom - 0.67) < 0.001; },
              'the zoom to be re-applied after two failures', 15000)
              .then(function () { return readState(); })
              .then(function (after) {
                assert(zoomFailuresLeft === 0, 'the failures were not all consumed');
                assert(after.run.originalZoom === 1,
                  'the original zoom must not be overwritten on resume, it is ' + after.run.originalZoom);
                var warned = after.log.slice(logLines).filter(function (e) {
                  return e.level === 'warn' && /zoom could not be set/.test(e.message);
                });
                assert(warned.length === 0, 'a recovered retry must not warn the user');
                return 'failed twice, then set 67% on the third try';
              });
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
            assert($('progressText').textContent.indexOf('of 30') !== -1,
              'progress reads "' + $('progressText').textContent + '"');
            return $('progressText').textContent;
          });
      });
    })

    .then(function () {
      return check('loading the catalog stores each picture once, not once per row', function () {
        return FYG.idb.run('images', 'readonly', function (store) { return store.getAll(); })
          .then(function (stored) {
            assert(stored.length === 3, 'expected 3 stored pictures, saw ' + stored.length);
            var types = stored.map(function (r) { return r.type; }).sort();
            assert(types.join(',') === 'image/jpeg,image/jpeg,image/png',
              'stored types are ' + types.join(','));
            return stored.length + ' pictures held once each';
          });
      });
    })

    .then(function () {
      return check('every row that has a picture carries its id, the rest carry none', function () {
        return readState().then(function (data) {
          var withImage = data.rows.filter(function (r) { return r.imageId; });
          // Five anchors sit on four rows, one of which carries two pictures.
          assert(withImage.length === 5, withImage.length + ' rows carry a picture, expected 5');
          assert(data.rows[0].imageId, 'the first row should carry one');
          assert(!data.rows[9].imageId, 'a row with no picture should carry an empty id');
          return withImage.length + ' rows carry a picture id';
        });
      });
    })

    .then(function () {
      return check('the worker serves a picture in parts that rebuild it exactly', function () {
        return FYG.idb.run('images', 'readonly', function (store) { return store.getAll(); })
          .then(function (stored) {
            var wanted = stored[0];
            return wanted.blob.arrayBuffer().then(function (buffer) {
              var original = new Uint8Array(buffer);
              var chunks = [];

              function pull(part, total) {
                if (total !== null && part >= total) return Promise.resolve();
                return chrome.runtime.sendMessage({
                  type: S.MSG.REQUEST_IMAGE, id: wanted.id, part: part
                }).then(function (reply) {
                  assert(reply && reply.ok, 'part ' + part + ' failed: ' + (reply && reply.error));
                  chunks.push(reply.data);
                  return pull(part + 1, reply.parts);
                });
              }

              return pull(0, null).then(function () {
                var rebuilt = U.base64ToBytes(chunks.join(''));
                assert(rebuilt.length === original.length,
                  'rebuilt ' + rebuilt.length + ' bytes from ' + original.length);
                // An exact identity check, not a spot check.
                assert(FYG.zip.crc32(rebuilt) === FYG.zip.crc32(original),
                  'the rebuilt picture does not match the stored one');
                return rebuilt.length + ' bytes rebuilt, checksum matches';
              });
            });
          });
      });
    })

    .then(function () {
      return check('a picture stored oversized is shrunk before it reaches the form', function () {
        /*
         * A catalog loaded by an older build holds whatever was written then, so
         * the page checks again rather than trusting what it is handed. Fygaro
         * refusing an upload part way through a run of this length is a very
         * expensive way to discover the limit.
         */
        var id = 'xl/media/oversized.png';
        var realTarget = FYG.imagefit.TARGET_BYTES;

        return makeNoisyPng(200).then(function (big) {
          return FYG.idb.run('images', 'readwrite', function (store) {
            return store.put({
              id: id, name: 'oversized.png', type: 'image/png', size: big.length,
              blob: new Blob([big], { type: 'image/png' })
            }, id);
          }).then(function () {
            // Stand in for the 2.5 MB limit at a size the fixture can reach.
            FYG.imagefit.TARGET_BYTES = Math.floor(big.length / 3);
            FYG.assets.clear();
            return FYG.assets.fetchImage(id);
          }).then(function (file) {
            FYG.imagefit.TARGET_BYTES = realTarget;
            assert(file.size <= Math.floor(big.length / 3),
              'the form was handed ' + file.size + ' bytes, over the limit');
            assert(file.size < big.length, 'it was not actually shrunk');
            assert(file.type === 'image/jpeg', 'it should arrive as JPEG, saw ' + file.type);
            assert(file.name === 'oversized.jpg', 'the name should follow, saw ' + file.name);
            return big.length + ' bytes stored, ' + file.size + ' bytes handed over';
          }).catch(function (err) {
            FYG.imagefit.TARGET_BYTES = realTarget;
            throw err;
          });
        });
      });
    })

    .then(function () {
      return check('an unknown picture is refused with a readable reason', function () {
        return chrome.runtime.sendMessage({
          type: S.MSG.REQUEST_IMAGE, id: 'xl/media/not-in-this-catalog.png', part: 0
        }).then(function (reply) {
          assert(reply && reply.ok === false, 'an unknown id should be refused');
          assert(/not in this catalog/i.test(reply.error), 'unhelpful message: ' + reply.error);
          return reply.error;
        });
      });
    })

    .then(function () {
      return check('rows sharing a picture cost one fetch, not one each', function () {
        // Rows 2 and 3 of the fixture share an image, as 80 rows do in the real
        // catalog. Refetching per row would move gigabytes for no reason.
        return readState().then(function (data) {
          var shared = data.rows[0].imageId;
          assert(shared, 'the first row should carry a picture');

          FYG.assets.clear();
          var asks = 0;
          var realSend = chrome.runtime.sendMessage;
          chrome.runtime.sendMessage = function (message) {
            if (message && message.type === S.MSG.REQUEST_IMAGE && message.part === 0) asks++;
            return realSend.apply(this, arguments);
          };

          var wanted = [];
          for (var i = 0; i < 30; i++) wanted.push(shared);

          return Promise.all(wanted.map(function (id) { return FYG.assets.fetchImage(id); }))
            .then(function (files) {
              chrome.runtime.sendMessage = realSend;
              assert(files.length === 30, 'expected 30 files');
              assert(files[0] instanceof File, 'a real File should come back');
              assert(files[0].size > 0, 'the file is empty');
              assert(asks === 1, 'the picture was fetched ' + asks + ' times, expected once');
              assert(FYG.assets.stats().entries <= FYG.assets.MAX_ENTRIES, 'the cache grew past its cap');
              return '30 rows, 1 fetch, ' + files[0].size + ' bytes';
            })
            .catch(function (err) {
              chrome.runtime.sendMessage = realSend;
              throw err;
            });
        });
      });
    })

    .then(function () {
      return check('a code that already exists skips the row instead of stopping the run', function () {
        var complaint = 'This code is already in use by another product or version';
        var notificationsBefore = recorded.notifications.length;

        return readState().then(function (before) {
          var at = before.run.cursor;
          return askForJob('dashboard')
            .then(function (job) { return finishStep(job.step); })
            .then(function () { return askForJob('productList'); })
            .then(function (job) { return finishStep(job.step); })
            .then(function () { return askForJob('productAdd'); })
            .then(function (job) {
              assert(job.act === 'run', 'expected the product form step, got "' + job.act + '"');
              // The form step is the only one that needs a picture, so this is
              // where the id has to have arrived.
              assert(job.row.imageId === before.rows[at].imageId,
                'the job carries "' + job.row.imageId + '" but the row has "' + before.rows[at].imageId + '"');
              return finishStep(job.step, { duplicate: complaint });
            })
            .then(readState)
            .then(function (after) {
              var row = after.rows[at];
              assert(row.status === S.ROW.SKIPPED, 'the row is "' + row.status + '", expected skipped');
              assert(row.reason === S.SKIP.EXISTS, 'the reason is "' + row.reason + '"');
              assert(/Ya existe en Fygaro/.test(row.error), 'unhelpful reason: ' + row.error);
              assert(row.link === '', 'a skipped row must carry no link');

              // The run has to keep going, not stop and wait for a person.
              assert(after.run.status === S.STATUS.RUNNING,
                'the run went to "' + after.run.status + '" instead of carrying on');
              assert(after.run.pending === null, 'an attention prompt was raised');
              assert(after.run.attempt === 0, 'the attempt counter moved to ' + after.run.attempt);
              assert(after.run.cursor > at, 'the cursor stayed on the skipped row');
              assert(after.run.step === S.STEP.NAV_TO_PRODUCTS,
                'the next row should start from the dashboard, not ' + after.run.step);
              assert(recorded.tabUpdates[recorded.tabUpdates.length - 1] === S.APP_URL,
                'the browser was not sent back to the dashboard');

              assert(after.run.stats.skipped === before.run.stats.skipped + 1,
                'the skipped count did not move');
              assert(after.run.stats.failed === before.run.stats.failed,
                'a skipped row must not be counted as a failure');
              assert(recorded.notifications.length === notificationsBefore,
                'a duplicate must not raise a desktop notification');
              return 'row skipped as "exists", run continued, no retry and no notification';
            });
        });
      });
    })

    .then(function () {
      return check('the panel labels an existing row differently from a plain skip', function () {
        // The panel re-renders on its own when the worker broadcasts, so this
        // waits for that rather than reaching into the panel to force it.
        return waitUntil(function () {
          return Array.prototype.map.call(
            $('results').querySelectorAll('.badge span'),
            function (el) { return el.textContent; }).indexOf('Exists') !== -1;
        }, 'the row to be labelled Exists').then(function () {
          var words = Array.prototype.map.call(
            $('results').querySelectorAll('.badge span'), function (el) { return el.textContent; });
          assert(words.indexOf('Failed') === -1, 'an existing row was shown as a failure');
          return 'shown as Exists, not Skipped and not Failed';
        });
      });
    })

    .then(function () {
      return check('a second row runs and the cursor moves on', function () {
        var link = 'https://www.fygaro.com/en/pb/aaaaaaaa-2222-2222-2222-222222222222/';
        return playRow('aaaaaaaa-0000-0000-0000-000000000002', link).then(function () {
          return readState();
        }).then(function (data) {
          assert(data.run.stats.done === 2, 'done count is ' + data.run.stats.done);
          assert(data.rows[0].status === S.ROW.DONE, 'the first row should be done');
          // The second row was the one Fygaro already had, so the run moved past
          // it and finished the third instead.
          assert(data.rows[1].status === S.ROW.SKIPPED, 'the skipped row changed to ' + data.rows[1].status);
          assert(data.rows[2].status === S.ROW.DONE, 'the third row should be done');
          assert(data.run.cursor === 3, 'the cursor is at ' + data.run.cursor + ', expected the fourth row');
          return 'two rows done either side of the skipped one';
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
          assert(after.rows.filter(function (r) { return r.link; }).length === 2,
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
                assert(byRow[row.sheetRow].I === row.link,
                  'sheet row ' + row.sheetRow + ' has "' + byRow[row.sheetRow].I + '" instead of its link');
              });
              // The new column needs its header, or the links below it mean
              // nothing and a reload cannot find them again.
              assert(byRow[1].I === 'Link', 'the Link header is "' + byRow[1].I + '"');
              assert(byRow[1].H === 'Columna 1', 'the images header changed to "' + byRow[1].H + '"');

              // Everything else must be exactly as it was.
              assert(byRow[2].D === 'Cita de Ingreso Presencial con Psicóloga Clínica del Equipo',
                'row 2 name changed');
              assert(wb.sheets.length === 6, 'sheets were lost');

              // Column H carries the product images and must never be written to.
              data.rows.forEach(function (r) {
                assert(!byRow[r.sheetRow].H,
                  'sheet row ' + r.sheetRow + ' column H was written into: "' + byRow[r.sheetRow].H + '"');
              });

              // Rows the run never reached must still have no link.
              var touched = done.map(function (r) { return r.sheetRow; });
              var untouched = data.rows.filter(function (r) { return touched.indexOf(r.sheetRow) === -1; });
              assert(untouched.length > 20, 'expected most rows to be untouched');
              untouched.forEach(function (r) {
                assert(!byRow[r.sheetRow].I,
                  'sheet row ' + r.sheetRow + ' should have no link, saw "' + byRow[r.sheetRow].I + '"');
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
      return check('the sheet saves itself as the run goes, with nothing clicked', function () {
        /*
         * The whole point. Links exist in Fygaro the moment they are created,
         * and a run of this length that ends without writing them anywhere is
         * hours of work sitting in a panel nobody can read.
         */
        savedToDisk = null;
        handlePermission = 'granted';
        $('targetOriginal').checked = true;
        $('targetOriginal').dispatchEvent(new Event('change'));

        return readState().then(function (before) {
          var links = before.rows.filter(function (r) { return r.link; }).length;
          assert(links > 0, 'the run should have captured links by now');

          // Run, then stop, because a run leaving the running state is when a
          // save is owed however few links arrived since the last one.
          return chrome.runtime.sendMessage({ type: S.MSG.START })
            .then(function () { return waitUntil(function () {
              return store[S.KEY_RUN] && store[S.KEY_RUN].status === S.STATUS.RUNNING;
            }, 'the run to start'); })
            .then(function () { return chrome.runtime.sendMessage({ type: S.MSG.STOP }); })
            .then(function () { return waitUntil(function () { return savedToDisk; },
              'the sheet to be written without being asked', 20000); })
            .then(function () { return savedToDisk.arrayBuffer(); })
            .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
            .then(function (wb) {
              var byRow = {};
              wb.readSheet(SHEET).rows.forEach(function (r) { byRow[r.r] = r.cells; });
              before.rows.filter(function (r) { return r.link; }).forEach(function (row) {
                assert(byRow[row.sheetRow].I === row.link,
                  'sheet row ' + row.sheetRow + ' holds "' + byRow[row.sheetRow].I + '"');
              });
              assert(byRow[1].I === 'Link', 'the header was not written');
              return links + ' links written with no button pressed';
            });
        });
      });
    })

    .then(function () {
      return check('Start refuses to run when the destination cannot be written to', function () {
        /*
         * Asked for up front, while the click is still live, because Chrome only
         * shows the permission prompt then. Discovering it hours later, with the
         * products already created, is the failure this prevents.
         */
        handlePermission = 'denied';
        recorded.alerts.length = 0;
        $('targetOriginal').checked = true;

        return Promise.resolve()
          .then(function () { $('btnStart').click(); })
          .then(function () { return waitUntil(function () { return recorded.alerts.length > 0; },
            'the refusal to be reported'); })
          .then(readState)
          .then(function (data) {
            var said = recorded.alerts[recorded.alerts.length - 1];
            assert(/Permission to write/.test(said), 'unhelpful message: ' + said);
            assert(/Saving/.test(said), 'it should point at where to fix it: ' + said);
            assert(data.run.status !== S.STATUS.RUNNING,
              'the run started anyway, with nowhere to save to');
            handlePermission = 'granted';
            return 'refused before creating anything';
          });
      });
    })

    .then(function () {
      return check('updating the original writes to the file itself, with no download', function () {
        savedToDisk = null;
        recorded.downloads.length = 0;
        handlePermission = 'granted';
        confirmAnswer = true;

        $('targetOriginal').checked = true;
        $('targetOriginal').dispatchEvent(new Event('change'));
        assert($('btnExportXlsx').textContent === 'Save now',
          'the button should say what it will do, it says "' + $('btnExportXlsx').textContent + '"');
        assert(/every \d+ links and whenever the run stops/.test($('saveTargetStatus').textContent),
          'the card should say it saves as it goes: ' + $('saveTargetStatus').textContent);

        return readState().then(function (data) {
          var done = data.rows.filter(function (r) { return r.status === S.ROW.DONE; });
          $('btnExportXlsx').click();

          return waitUntil(function () { return savedToDisk; }, 'the file to be written', 20000)
            .then(function () {
              assert(recorded.downloads.length === 0,
                'updating the original must not also download a copy');
              return savedToDisk.arrayBuffer();
            })
            .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
            .then(function (wb) {
              var byRow = {};
              wb.readSheet(SHEET).rows.forEach(function (r) { byRow[r.r] = r.cells; });
              done.forEach(function (row) {
                assert(byRow[row.sheetRow].I === row.link,
                  'sheet row ' + row.sheetRow + ' was written as "' + byRow[row.sheetRow].I + '"');
              });
              assert(byRow[1].I === 'Link', 'the in place save lost the Link header');
              assert(wb.sheets.length === 6, 'the other sheets were lost');
              assert(byRow[2].D === 'Cita de Ingreso Presencial con Psicóloga Clínica del Equipo',
                'row 2 was damaged');
              return done.length + ' links written into the file on disk, nothing downloaded';
            });
        });
      });
    })

    .then(function () {
      return check('a refused permission fails loudly and writes nothing', function () {
        savedToDisk = null;
        handlePermission = 'denied';
        recorded.alerts.length = 0;

        $('btnExportXlsx').click();
        return waitUntil(function () { return recorded.alerts.length > 0; }, 'the failure to be reported')
          .then(function () {
            assert(savedToDisk === null, 'nothing may be written without permission');
            var said = recorded.alerts[recorded.alerts.length - 1];
            assert(/could not be updated/.test(said), 'unhelpful message: ' + said);
            assert(/Excel/.test(said), 'the message should mention closing the file in Excel');
            handlePermission = 'granted';
            return 'refused cleanly, the file was left alone';
          });
      });
    })

    .then(function () {
      return check('saving into a copy still downloads and leaves the file alone', function () {
        savedToDisk = null;
        recorded.downloads.length = 0;
        $('targetCopy').checked = true;
        $('targetCopy').dispatchEvent(new Event('change'));
        assert($('btnExportXlsx').textContent === 'Download updated .xlsx',
          'the button should change back, it says "' + $('btnExportXlsx').textContent + '"');
        // A download has nothing to add beyond what the option itself says.
        assert($('saveTargetStatus').textContent === '',
          'the status line should stay quiet for a download: ' + $('saveTargetStatus').textContent);

        $('btnExportXlsx').click();
        return waitUntil(function () { return recorded.downloads.length > 0; }, 'the copy to download', 20000)
          .then(function () {
            assert(savedToDisk === null, 'saving into a copy must never touch the original');
            return 'downloaded a copy, original untouched';
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
      return check('Clear everything empties storage and resets the whole panel', function () {
        // Leave some interface state behind, so the reset has something to undo.
        $('search').value = 'CT-ING';
        $('search').dispatchEvent(new Event('input'));
        $('filters').querySelector('[data-filter="done"]').click();
        store.fyg_legacy_key = { left: 'behind by an older version' };
        confirmAnswer = true;
        recorded.confirms.length = 0;

        $('btnResetAll').click();

        return waitUntil(function () { return Object.keys(store).length === 0; }, 'storage to empty')
          .then(function () {
            var asked = recorded.confirms[recorded.confirms.length - 1];
            assert(/Clear everything/.test(asked), 'the prompt should say what it does: ' + asked);
            assert(/captured link/.test(asked), 'the prompt should name the cost: ' + asked);

            assert(!(S.KEY_RUN in store), 'the run was left behind');
            assert(!(S.KEY_ROWS in store), 'the rows were left behind');
            assert(!(S.KEY_LOG in store), 'the log was left behind');
            assert(!(S.KEY_FILE in store), 'the cached workbook was left behind');
            assert(!('fyg_legacy_key' in store), 'a key from an older version survived');

            return waitUntil(function () { return $('statTotal').textContent === '0'; }, 'the panel to reset');
          })
          .then(function () {
            assert($('search').value === '', 'the search box was not cleared');
            assert($('filters').querySelector('[data-filter="all"]').getAttribute('aria-pressed') === 'true',
              'the filter was not put back to All');
            assert($('targetCopy').checked, 'the save target was not put back to a copy');
            assert($('targetOriginal').disabled, 'the file handle was not forgotten');
            assert($('fileEmpty').classList.contains('hidden') === false, 'the file picker is not shown again');
            assert($('fileLoaded').classList.contains('hidden'), 'the loaded file panel is still showing');
            assert($('resetSummary').textContent === 'Nothing is stored yet.',
              'the summary still reads "' + $('resetSummary').textContent + '"');
            assert($('btnStart').disabled, 'Start should be disabled again');
            assert($('sheetSelect').options.length === 0, 'the sheet list was not cleared');
            return 'storage empty, panel back to its first run state';
          });
      });
    })

    .then(report);
  }

  function report() {
    var failed = results.filter(function (r) { return !r.ok; });
    var summary = document.getElementById('reportSummary');
    summary.className = failed.length ? 'fail' : 'pass';
    summary.textContent = failed.length
      ? failed.length + ' of ' + results.length + ' checks FAILED'
      : 'All ' + results.length + ' checks passed';
    document.title = (failed.length ? 'FAIL ' + failed.length + '/' : 'PASS 0/') + results.length;
    // Report back to the test runner. The pages are driven in real time rather
    // than under a virtual clock, because Chrome's virtual time fast forwards
    // past IndexedDB completion callbacks whenever the page looks idle, which
    // left transactions hanging for ever.
    if (location.protocol.indexOf('http') === 0) {
      fetch('/__result', {
        method: 'POST',
        body: JSON.stringify({
          title: document.title,
          failures: results.filter(function (r) { return !r.ok; })
            .map(function (r) { return r.name + ' -- ' + r.detail; })
        })
      }).catch(function () {});
    }


    var list = document.getElementById('reportResults');
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

      fakeHandle = {
        kind: 'file',
        name: WORKBOOK_NAME,
        queryPermission: function () { return Promise.resolve(handlePermission); },
        requestPermission: function () { return Promise.resolve(handlePermission); },
        getFile: function () {
          return Promise.resolve(new File([workbookBytes], WORKBOOK_NAME,
            { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        },
        createWritable: function () {
          if (handlePermission !== 'granted') return Promise.reject(new Error('not allowed'));
          var chunks = [];
          return Promise.resolve({
            write: function (blob) { chunks.push(blob); return Promise.resolve(); },
            // Only a close that is reached commits, mirroring the real API.
            close: function () { savedToDisk = new Blob(chunks); return Promise.resolve(); }
          });
        }
      };
      window.showOpenFilePicker = function () { return Promise.resolve([fakeHandle]); };

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
