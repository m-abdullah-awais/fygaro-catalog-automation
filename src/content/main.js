/*
 * Fygaro Catalog Automation
 * Content script driver.
 *
 * Fygaro is a single page app, so the script is injected once and then has to
 * notice route changes itself. A content script runs in an isolated world and
 * cannot patch the page's own history calls, so the location is polled instead,
 * which is cheap and cannot be defeated by however the app routes.
 *
 * Nothing here decides what to do next. It reports the current route, asks the
 * worker for a job, runs exactly that one step, and reports the outcome.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG;
  var S = FYG.state;
  var U = FYG.util;

  var HEARTBEAT_MS = 2000;
  var LOCATION_POLL_MS = 400;
  var REPEAT_GUARD_MS = 4000;

  var busy = false;
  var requesting = false;
  var lastHref = location.href;
  var lastJobKey = '';
  var lastJobAt = 0;

  function send(type, payload) {
    var message = Object.assign({ type: type }, payload || {});
    return chrome.runtime.sendMessage(message).catch(function () {
      // The worker may be restarting. The next heartbeat will pick things up.
      return null;
    });
  }

  function report(level, message) {
    send(S.MSG.CONTENT_LOG, { level: level, message: message });
  }

  /* ---------------------------------------------------------------- overlay */

  var hud = null;

  function buildHud() {
    if (hud || !document.body) return;

    var host = document.createElement('div');
    host.id = 'fygaro-automation-hud';
    host.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;';
    // A shadow root keeps Fygaro's stylesheet and this widget completely apart.
    var shadow = host.attachShadow({ mode: 'open' });
    shadow.innerHTML =
      '<style>' +
      ':host{all:initial}' +
      '.card{display:none;align-items:center;gap:10px;font:500 12px/1.35 -apple-system,"Segoe UI",Roboto,Arial,sans-serif;' +
      'color:#1A1F26;background:#FFFFFF;border:1px solid #E3E6EA;border-radius:10px;padding:10px 12px;' +
      'box-shadow:0 1px 2px rgba(16,24,40,.05),0 6px 16px rgba(16,24,40,.10);max-width:320px}' +
      '.card.on{display:flex}' +
      '.dot{width:8px;height:8px;border-radius:50%;background:#2457D6;flex:none}' +
      '.dot.run{animation:p 1.2s ease-in-out infinite}' +
      '.dot.warn{background:#B45309;animation:none}' +
      '.dot.bad{background:#C02626;animation:none}' +
      '.dot.ok{background:#17803D;animation:none}' +
      '@keyframes p{0%,100%{opacity:1}50%{opacity:.25}}' +
      '.txt{min-width:0}' +
      '.t1{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}' +
      '.t2{color:#5B6673;font-weight:400;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}' +
      'button{font:600 11px/1 inherit;color:#5B6673;background:#F6F7F9;border:1px solid #CBD2D9;' +
      'border-radius:6px;padding:6px 9px;cursor:pointer;flex:none}' +
      'button:hover{background:#EEF3FF;border-color:#2457D6;color:#2457D6}' +
      'button:focus-visible{outline:2px solid #2457D6;outline-offset:2px}' +
      '</style>' +
      '<div class="card" role="status" aria-live="polite">' +
      '<span class="dot"></span>' +
      '<span class="txt"><span class="t1"></span><br><span class="t2"></span></span>' +
      '<button type="button">Pause</button>' +
      '</div>';

    document.body.appendChild(host);
    hud = {
      card: shadow.querySelector('.card'),
      dot: shadow.querySelector('.dot'),
      t1: shadow.querySelector('.t1'),
      t2: shadow.querySelector('.t2'),
      button: shadow.querySelector('button')
    };
    hud.button.addEventListener('click', function () {
      send(hud.button.dataset.act === 'resume' ? S.MSG.RESUME : S.MSG.PAUSE);
    });
  }

  function paintHud(info) {
    buildHud();
    if (!hud) return;

    if (!info || info.status === S.STATUS.IDLE) {
      hud.card.classList.remove('on');
      return;
    }

    hud.card.classList.add('on');
    hud.dot.className = 'dot' +
      (info.status === S.STATUS.RUNNING ? ' run' :
       info.status === S.STATUS.PAUSED ? ' warn' :
       info.status === S.STATUS.ATTENTION ? ' bad' :
       info.status === S.STATUS.DONE ? ' ok' : '');

    var progress = info.done + ' of ' + info.total;
    if (info.status === S.STATUS.RUNNING) {
      hud.t1.textContent = 'Row ' + info.sheetRow + ' · ' + U.truncate(info.code || '', 28);
      hud.t2.textContent = info.stepLabel + ' · ' + progress;
    } else if (info.status === S.STATUS.ATTENTION) {
      hud.t1.textContent = 'Needs attention';
      hud.t2.textContent = U.truncate(info.message || 'A step could not be completed.', 44);
    } else if (info.status === S.STATUS.PAUSED) {
      hud.t1.textContent = 'Paused';
      hud.t2.textContent = progress + ' done';
    } else {
      hud.t1.textContent = 'Finished';
      hud.t2.textContent = progress + ' done';
    }

    var resumable = info.status === S.STATUS.PAUSED || info.status === S.STATUS.ATTENTION;
    hud.button.dataset.act = resumable ? 'resume' : 'pause';
    hud.button.textContent = resumable ? 'Resume' : 'Pause';
    hud.button.style.display = info.status === S.STATUS.DONE ? 'none' : '';
  }

  /* ------------------------------------------------------------------ steps */

  function execute(job) {
    var runner = FYG.steps[job.step];
    if (!runner) {
      return send(S.MSG.STEP_FAILED, {
        step: job.step,
        message: 'Unknown step "' + job.step + '".',
        url: location.href
      });
    }

    busy = true;
    report('info', 'Row ' + job.row.sheetRow + ': ' + (S.STEP_LABEL[job.step] || job.step) + '.');

    return Promise.resolve()
      .then(function () { return runner(job); })
      .then(function (result) {
        result = result || {};
        return send(S.MSG.STEP_DONE, {
          step: job.step,
          url: location.href,
          link: result.link || '',
          productUuid: result.productUuid || '',
          landedRoute: result.landedRoute || ''
        });
      })
      .catch(function (err) {
        return send(S.MSG.STEP_FAILED, {
          step: job.step,
          url: location.href,
          message: String(err && err.message ? err.message : err)
        });
      })
      .then(function () {
        busy = false;
        // The step usually navigated, so look immediately for the next job.
        setTimeout(tick, 250);
      });
  }

  function tick() {
    if (busy || requesting) return;

    var route = S.routeOf(location.pathname);
    requesting = true;

    send(S.MSG.REQUEST_JOB, { route: route, url: location.href })
      .then(function (job) {
        requesting = false;
        if (!job) return;

        paintHud(job.hud);
        if (job.act !== 'run' || busy) return;

        // Guards against the same job running twice if two ticks overlap around
        // a re-render. A retry carries a new attempt number and is allowed.
        var key = [job.step, location.href, job.row.sheetRow, job.attempt].join('|');
        var now = Date.now();
        if (key === lastJobKey && now - lastJobAt < REPEAT_GUARD_MS) return;
        lastJobKey = key;
        lastJobAt = now;

        execute(job);
      })
      .catch(function () {
        requesting = false;
      });
  }

  /* --------------------------------------------------------------- watchers */

  setInterval(function () {
    if (location.href === lastHref) return;
    lastHref = location.href;
    // A route change resets the guard so the next step can start at once.
    lastJobKey = '';
    tick();
  }, LOCATION_POLL_MS);

  setInterval(tick, HEARTBEAT_MS);

  // Reacts the instant the side panel starts, pauses or resumes a run.
  chrome.storage.onChanged.addListener(function (changes, area) {
    if (area === 'local' && changes[S.KEY_RUN]) tick();
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { buildHud(); tick(); });
  } else {
    buildHud();
    tick();
  }
})(typeof globalThis !== 'undefined' ? globalThis : self);
