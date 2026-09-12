/*
 * Fygaro Catalog Automation
 * Presses "More Results" until a list has nothing left to add.
 *
 * Fygaro paginates its products and its payment links behind the same button at
 * the foot of the page, so a list of two thousand rows takes a hundred clicks to
 * see in full. This does that, and nothing else: it does not read the rows, does
 * not touch the run, and leaves the page exactly as a person clicking the button
 * themselves would have left it.
 *
 * Growth is measured by counting rows rather than by watching the button. The
 * button stays on the page while a page of results is being fetched, and it is
 * removed only once the server says there is no more, so counting is the only
 * signal that says the click actually did something.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var L = FYG.loadall || (FYG.loadall = {});
  var D = FYG.dom;
  var U = FYG.util;
  var S = FYG.state;

  /*
   * A ceiling on clicks, so a button that reappears without ever adding a row
   * cannot spin for ever. At roughly twenty rows a page this is far more than
   * the largest catalog needs.
   */
  L.MAX_CLICKS = 600;

  /* How long one page of results is given to arrive before it counts as the end. */
  L.PAGE_TIMEOUT_MS = 15000;

  /** True while a load is in progress, so two cannot overlap. */
  var running = false;

  function locate() {
    return FYG.steps.locate;
  }

  /**
   * Loads every page of whichever list is on screen.
   *
   * @param {{minDelayMs: number, maxDelayMs: number, onProgress: function}} [opts]
   * @returns {Promise<{ok, kind, rows, added, clicks, stopped}>}
   */
  L.run = function (opts) {
    var options = opts || {};
    var pace = {
      minDelayMs: options.minDelayMs == null ? 400 : options.minDelayMs,
      maxDelayMs: options.maxDelayMs == null ? 900 : options.maxDelayMs
    };
    var announce = options.onProgress || function () {};

    if (running) {
      return Promise.resolve({
        ok: false, kind: '', rows: 0, added: 0, clicks: 0,
        stopped: 'A load is already running on this page.'
      });
    }

    var kind = S.LIST_KIND[S.routeOf(location.pathname)] || 'rows';
    var startedWith = locate().listRowCount();

    if (!locate().moreResultsButton()) {
      // Either everything is already on screen, or this is not a list at all.
      // Both are answered the same way: there is nothing here to load.
      return Promise.resolve({
        ok: true, kind: kind, rows: startedWith, added: 0, clicks: 0,
        stopped: startedWith ? 'everything was already loaded' : 'this page has no list to load'
      });
    }

    running = true;
    var clicks = 0;
    var stopped = '';

    function step() {
      if (clicks >= L.MAX_CLICKS) {
        stopped = 'stopped at the ' + L.MAX_CLICKS + ' click limit';
        return Promise.resolve();
      }

      var button = locate().moreResultsButton();
      if (!button) {
        stopped = 'the list ran out of pages';
        return Promise.resolve();
      }

      var before = locate().listRowCount();

      // Bring the foot of the page into view first. The button is the last thing
      // on a very long page, and a list that grows by twenty rows a click moves
      // it down the document every time.
      window.scrollTo(0, document.body.scrollHeight);
      D.reveal(button);
      D.click(button);
      clicks++;

      return D.waitFor(function () {
        if (locate().listRowCount() > before) return 'grew';
        // The button going away is the other way a click can finish: it means
        // that was the last page.
        if (!locate().moreResultsButton()) return 'ended';
        return null;
      }, L.PAGE_TIMEOUT_MS, 'the next page of ' + kind)
        .then(function (how) {
          announce({ kind: kind, rows: locate().listRowCount(), clicks: clicks });
          if (how === 'ended') {
            stopped = 'the list ran out of pages';
            return null;
          }
          // Paced like the rest of the automation, so a hundred clicks do not
          // arrive as a hundred requests at machine speed.
          return U.randomDelay(pace.minDelayMs, pace.maxDelayMs).then(step);
        })
        .catch(function () {
          stopped = 'no more rows arrived, so it stopped there';
          return null;
        });
    }

    return step().then(function () {
      running = false;
      var rows = locate().listRowCount();
      return {
        ok: true,
        kind: kind,
        rows: rows,
        added: rows - startedWith,
        clicks: clicks,
        stopped: stopped
      };
    }, function (err) {
      running = false;
      return {
        ok: false, kind: kind, rows: locate().listRowCount(), added: 0, clicks: clicks,
        stopped: String(err && err.message ? err.message : err)
      };
    });
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
