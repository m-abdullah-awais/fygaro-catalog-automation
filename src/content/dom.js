/*
 * Fygaro Catalog Automation
 * DOM helpers for driving the Fygaro app.
 *
 * Two things shape this file. Fygaro is a React app, so writing to `value`
 * directly is invisible to it and every write has to go through the native
 * setter plus real events. And its class names are content hashed, so nothing
 * here may depend on them: locators use `name`, `href` and visible text.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var D = FYG.dom || (FYG.dom = {});
  var U = FYG.util;

  /** True when the element is actually rendered and on screen. */
  D.isVisible = function (el) {
    if (!el || !el.getClientRects) return false;
    if (el.disabled) return false;
    if (el.getClientRects().length === 0) return false;
    var style = getComputedStyle(el);
    return style.visibility !== 'hidden' && style.display !== 'none' && style.opacity !== '0';
  };

  /**
   * The product form renders its Currency, Type of Item and Show In Website
   * controls twice, once per responsive layout column. This returns every match
   * so all copies can be written, with the visible ones first.
   */
  D.all = function (selector, scope) {
    var list = Array.prototype.slice.call((scope || document).querySelectorAll(selector));
    return list.sort(function (a, b) {
      return (D.isVisible(b) ? 1 : 0) - (D.isVisible(a) ? 1 : 0);
    });
  };

  D.visible = function (selector, scope) {
    return D.all(selector, scope).filter(D.isVisible);
  };

  D.first = function (selector, scope) {
    return D.visible(selector, scope)[0] || null;
  };

  /**
   * Finds elements by their visible text. Comparison folds case and accents so
   * the English and Spanish interfaces both resolve.
   * @param {string} selector
   * @param {string[]} wanted
   * @param {{mode?: 'equals'|'includes', scope?: Element}} [opts]
   */
  D.byText = function (selector, wanted, opts) {
    opts = opts || {};
    var mode = opts.mode || 'equals';
    var needles = wanted.map(U.foldText);
    return D.all(selector, opts.scope).filter(function (el) {
      var text = U.foldText(el.textContent);
      if (!text) return false;
      for (var i = 0; i < needles.length; i++) {
        if (mode === 'equals' ? text === needles[i] : text.indexOf(needles[i]) !== -1) return true;
      }
      return false;
    });
  };

  D.firstByText = function (selector, wanted, opts) {
    var hits = D.byText(selector, wanted, opts).filter(D.isVisible);
    return hits[0] || null;
  };

  /**
   * Waits until `probe` returns something truthy.
   * Backed by a MutationObserver so it reacts the instant React re renders, with
   * an interval as a safety net for changes that do not touch the DOM tree.
   * @returns {Promise<*>} whatever the probe returned
   */
  D.waitFor = function (probe, timeoutMs, label) {
    var timeout = timeoutMs || 20000;
    return new Promise(function (resolve, reject) {
      var settled = false;
      var startedAt = Date.now();
      var observer = null;
      var timer = null;

      function stop() {
        settled = true;
        if (observer) observer.disconnect();
        if (timer) clearInterval(timer);
      }

      function check() {
        if (settled) return;
        var value = null;
        try { value = probe(); } catch (e) { value = null; }
        if (value) { stop(); resolve(value); return; }
        if (Date.now() - startedAt >= timeout) {
          stop();
          reject(new Error('Timed out after ' + Math.round(timeout / 1000) + 's waiting for ' + (label || 'the page')));
        }
      }

      observer = new MutationObserver(check);
      observer.observe(document.documentElement, {
        childList: true, subtree: true, attributes: true, characterData: true
      });
      timer = setInterval(check, 250);
      check();
    });
  };

  /** Waits for an element matching a selector, and returns it. */
  D.waitForSelector = function (selector, timeoutMs, label) {
    return D.waitFor(function () { return D.first(selector); }, timeoutMs, label || selector);
  };

  /**
   * Writes a value the way a user would, so React's state updates with it.
   * Assigning `el.value` alone is swallowed by React's own value tracker.
   */
  D.setNativeValue = function (el, value) {
    var proto;
    if (el instanceof HTMLTextAreaElement) proto = HTMLTextAreaElement.prototype;
    else if (el instanceof HTMLSelectElement) proto = HTMLSelectElement.prototype;
    else proto = HTMLInputElement.prototype;

    var setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
    setter.call(el, String(value));
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  };

  /** Fills a text input and confirms the value stuck. */
  D.fillText = function (el, value) {
    el.focus();
    D.setNativeValue(el, value);
    el.dispatchEvent(new Event('blur', { bubbles: true }));
    return el.value === String(value);
  };

  /** Selects an option by its value attribute, verifying the option exists. */
  D.selectValue = function (el, value) {
    var found = Array.prototype.some.call(el.options, function (o) { return o.value === String(value); });
    if (!found) return false;
    D.setNativeValue(el, value);
    return el.value === String(value);
  };

  /** Ticks or unticks a checkbox through a real click, so React sees it. */
  D.setChecked = function (el, wanted) {
    if (!!el.checked === !!wanted) return true;
    el.click();
    return !!el.checked === !!wanted;
  };

  /** Clicks an element the way a user reaches it. */
  D.click = function (el) {
    if (!el) return false;
    try { el.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (e) { /* older engines */ }
    if (typeof el.focus === 'function') { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
    el.click();
    return true;
  };

  /** A short human readable description of an element, for error messages. */
  D.describe = function (el) {
    if (!el) return 'nothing';
    var bits = [el.tagName.toLowerCase()];
    if (el.name) bits.push('name="' + el.name + '"');
    if (el.getAttribute && el.getAttribute('href')) bits.push('href="' + el.getAttribute('href') + '"');
    var text = U.truncate(U.textOf(el), 40);
    if (text) bits.push('"' + text + '"');
    return bits.join(' ');
  };

  /**
   * Collects anything on the page that looks like a validation complaint, so a
   * rejected save reports why instead of just saying it did not navigate.
   */
  D.collectErrors = function () {
    var found = [];
    var nodes = D.all('[role="alert"], [class*="error" i], [class*="invalid" i]');
    nodes.forEach(function (el) {
      if (!D.isVisible(el)) return;
      var text = U.normText(el.textContent);
      if (text && text.length < 240 && found.indexOf(text) === -1) found.push(text);
    });
    D.all('input, select, textarea').forEach(function (el) {
      if (el.validationMessage && !el.validity.valid) {
        var msg = (el.name ? el.name + ': ' : '') + el.validationMessage;
        if (found.indexOf(msg) === -1) found.push(msg);
      }
    });
    return found.slice(0, 4);
  };

  /** Current path, without the query string or hash. */
  D.path = function () {
    return location.pathname;
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
