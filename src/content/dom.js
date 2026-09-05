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
   * Finds a form control, preferring one that is on screen but accepting one
   * that is not.
   *
   * Fygaro styles its checkboxes by hiding the real input and painting a div
   * beside it, so the element that has to be read and clicked has no boxes at
   * all. Requiring visibility here silently found nothing on the live site even
   * though the control was right there, so anything that reads or writes a form
   * value must use this rather than D.first.
   *
   * D.all already sorts visible matches first, so a genuinely duplicated control
   * still resolves to the copy the user can see.
   */
  D.control = function (selector, scope) {
    return D.all(selector, scope)[0] || null;
  };

  /** Waits for a form control to exist, whether or not it is painted. */
  D.waitForControl = function (selector, timeoutMs, label) {
    return D.waitFor(function () { return D.control(selector); }, timeoutMs, label || selector);
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

  /**
   * Ticks or unticks a checkbox through a real click, so React sees it.
   *
   * The input itself is tried first. If that does not take, the wrapping label
   * is clicked instead, which is literally what a user does when the real input
   * is hidden behind a painted replacement.
   */
  D.setChecked = function (el, wanted) {
    if (!el) return false;
    if (!!el.checked === !!wanted) return true;

    D.reveal(el);
    el.click();
    if (!!el.checked === !!wanted) return true;

    var label = el.closest ? el.closest('label') : null;
    if (label) {
      label.click();
      if (!!el.checked === !!wanted) return true;
    }
    return false;
  };

  /** Scrolls a control into view, using its label when the control is hidden. */
  D.reveal = function (el) {
    var anchor = D.isVisible(el) ? el : ((el.closest && el.closest('label')) || el.parentElement || el);
    try { anchor.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (e) { /* older engines */ }
  };

  /** Clicks an element the way a user reaches it. */
  D.click = function (el) {
    if (!el) return false;
    D.reveal(el);
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
   * Validation text that a control's own label renders after it.
   *
   * Fygaro gives that message a content hashed class with no "error" or
   * "invalid" in it, puts no role="alert" and no aria-invalid anywhere on the
   * page, and leaves the input natively valid, so every query in D.collectErrors
   * below finds nothing at all. The message is located by association instead:
   * an element child of the label wrapping the control, positioned after the
   * child that contains the control.
   *
   * The ordering rule is what keeps this honest. Decorative captions such as the
   * "Add Image" span sit before their control, complaints sit after it.
   *
   * Presence is the signal, not visibility. Demanding visibility is what once
   * made a locator pass against the snapshots and find nothing on the live site.
   *
   * @param {Element} control
   * @returns {string[]} messages in document order
   */
  D.messagesFor = function (control) {
    if (!control || !control.closest) return [];
    var label = control.closest('label');
    if (!label) return [];

    var kids = Array.prototype.slice.call(label.children);
    var at = -1;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i] === control || kids[i].contains(control)) { at = i; break; }
    }
    if (at === -1) return [];

    var out = [];
    for (var j = at + 1; j < kids.length; j++) {
      var text = U.normText(U.textOf(kids[j]));
      if (text && text.length < 240 && out.indexOf(text) === -1) out.push(text);
    }
    return out;
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
    // Fygaro's own field complaints match none of the queries above, so without
    // this a rejected save reports "did not lead anywhere" and says nothing
    // about the reason the page gave.
    D.all('input, select, textarea').forEach(function (el) {
      D.messagesFor(el).forEach(function (text) {
        var line = (el.name ? el.name + ': ' : '') + text;
        if (found.indexOf(line) === -1) found.push(line);
      });
    });
    return found.slice(0, 4);
  };

  /** Current path, without the query string or hash. */
  D.path = function () {
    return location.pathname;
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
