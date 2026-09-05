/*
 * Fygaro Catalog Automation
 * The seven steps that turn one catalog row into one Fygaro link.
 *
 * Finding things and doing things are kept apart. Everything in `STEPS.locate`
 * is a pure lookup against a scope, which is what tests/selectors.test.html
 * exercises using the saved page snapshots. The step functions below hold only
 * the flow: pause, act, confirm the navigation actually happened.
 *
 * No locator may depend on a class name. Fygaro's classes are content hashed
 * and change on every deploy, so `name`, `href` and visible text are used
 * instead, and text matching folds case and accents to cover EN and ES.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var STEPS = FYG.steps || (FYG.steps = {});
  var D = FYG.dom;
  var U = FYG.util;
  var S = FYG.state;

  var PATH = {
    products: /^\/(?:en|es)\/app\/products\/?$/,
    productAdd: /^\/(?:en|es)\/app\/products\/add\/?$/,
    productDetail: /^\/(?:en|es)\/app\/products\/[0-9a-f-]{36}\/permalink\/product\/?$/i,
    linkAdd: /^\/(?:en|es)\/app\/payment-buttons\/payments\/payment-buttons\/add\/?$/,
    dashboard: /^\/(?:en|es)\/app\/dashboard\/?$/
  };

  var LINK_PATTERN = /^https?:\/\/[^\s]*\/pb\/[0-9a-f-]{36}\/?$/i;

  /*
   * Wording Fygaro uses when a product Code is already taken, folded so case and
   * accents do not matter.
   *
   * The English string is verbatim from the captured page. The Spanish ones are
   * the likely translations and have NOT been seen live, so an unrecognised
   * message deliberately falls through to the normal failure path with its exact
   * text reported, rather than a row being skipped on a guess.
   */
  var DUPLICATE_CODE_TEXT = [
    'already in use',
    'already exists',
    'ya esta en uso',
    'ya se esta usando',
    'ya existe',
    'codigo duplicado',
    'duplicate code'
  ];

  /* --------------------------------------------------------------- locators */

  /** Visible anchors whose path matches, in document order. */
  function linksMatching(pattern, scope) {
    return D.visible('a[href]', scope).filter(function (a) {
      var path;
      try { path = new URL(a.getAttribute('href'), location.origin).pathname; } catch (e) { return false; }
      return pattern.test(path);
    });
  }

  /** The product name on a list row, without the leading "View" label. */
  function listItemName(anchor) {
    var h3 = anchor.querySelector('h3');
    if (!h3) return '';
    var clone = h3.cloneNode(true);
    Array.prototype.forEach.call(clone.querySelectorAll('span'), function (s) { s.remove(); });
    return U.normText(clone.textContent);
  }

  var locate = STEPS.locate = {};

  /** Sidebar entry that opens the products area. */
  locate.productsSidebarLink = function (scope) {
    return linksMatching(PATH.products, scope)[0] || null;
  };

  /**
   * The Create control in the header, never the centred one the empty state
   * renders. The header is identified by the Refresh control beside it, with the
   * inline centring style as a second signal.
   */
  locate.createItemLink = function (scope) {
    var candidates = linksMatching(PATH.productAdd, scope);
    if (!candidates.length) return null;

    var inToolbar = candidates.filter(function (a) {
      var box = a.parentElement;
      for (var depth = 0; box && depth < 3; depth++, box = box.parentElement) {
        if (box.querySelector('i[class*="fa-refresh"]')) return true;
      }
      return false;
    })[0];
    if (inToolbar) return inToolbar;

    var notCentred = candidates.filter(function (a) {
      return !a.closest('[style*="text-align"]');
    })[0];
    return notCentred || candidates[0];
  };

  /** The centred Create link, which exists only so the tests can reject it. */
  locate.emptyStateCreateLink = function (scope) {
    return linksMatching(PATH.productAdd, scope).filter(function (a) {
      return !!a.closest('[style*="text-align"]');
    })[0] || null;
  };

  /** A products list row matching both the name and the code. */
  locate.productListItem = function (scope, name, code) {
    var wantName = U.foldText(name);
    var wantCode = U.foldText(code || '');
    var anchors = linksMatching(PATH.productDetail, scope);
    for (var i = 0; i < anchors.length; i++) {
      if (U.foldText(listItemName(anchors[i])) !== wantName) continue;
      if (wantCode) {
        var detail = U.foldText(U.textOf(anchors[i].querySelector('p')));
        if (detail.indexOf(wantCode) === -1) continue;
      }
      return anchors[i];
    }
    return null;
  };

  /** The "Create Fygaro Link for Product" action on a product page. */
  locate.createLinkAction = function (scope) {
    return linksMatching(PATH.linkAdd, scope)[0] ||
      D.firstByText('a, button', ['Create Fygaro Link for Product', 'Crear Fygaro Link para el Producto'],
        { mode: 'includes', scope: scope });
  };

  /** The Show or Hide Advanced Options toggle on the link form. */
  locate.advancedOptionsToggle = function (scope) {
    return D.firstByText('button', ['Advanced Options', 'Opciones Avanzadas'],
      { mode: 'includes', scope: scope });
  };

  /**
   * A checkout requirement checkbox.
   * Type qualified deliberately: `max_successful_payments` exists twice on this
   * form, once as a checkbox and once as a number input.
   */
  locate.requirementCheckbox = function (scope, name) {
    // D.control, not D.first: Fygaro hides the real input and paints a div in
    // its place, so requiring visibility finds nothing on the live site.
    return D.control('input[type="checkbox"][name="' + name + '"]', scope);
  };

  locate.saveButton = function (scope) {
    return D.firstByText('button[type="submit"]', ['Save', 'Guardar'], { scope: scope }) ||
      D.byText('button[type="submit"]', ['Save', 'Guardar'], { scope: scope })[0] ||
      D.control('button[type="submit"]', scope);
  };

  /**
   * The complaint Fygaro renders beside the Code field, or ''.
   *
   * No class name is used. The message span's class is content hashed and holds
   * neither "error" nor "invalid", the page carries no role="alert" and no
   * aria-invalid, and the input stays natively valid, so it is found by its
   * position relative to the Code input instead.
   *
   * Every copy of the field is tried, because this form renders some of its
   * controls twice, once per responsive layout.
   */
  locate.codeFieldMessage = function (scope) {
    var inputs = D.all('input[name="code"]', scope);
    for (var i = 0; i < inputs.length; i++) {
      var messages = D.messagesFor(inputs[i]);
      if (messages.length) return messages[0];
    }
    return '';
  };

  /**
   * That complaint, but only when it says the code is already taken.
   * @returns {string} the message, or '' when the code was not refused as a duplicate
   */
  locate.duplicateCodeError = function (scope) {
    var text = locate.codeFieldMessage(scope);
    if (!text) return '';
    var folded = U.foldText(text);
    for (var i = 0; i < DUPLICATE_CODE_TEXT.length; i++) {
      if (folded.indexOf(DUPLICATE_CODE_TEXT[i]) !== -1) return text;
    }
    return '';
  };

  locate.backHomeLink = function (scope) {
    return linksMatching(PATH.dashboard, scope)[0] || null;
  };

  locate.copyLinkButton = function (scope) {
    return D.firstByText('button', ['Copy Link', 'Copiar Link', 'Copiar Enlace'],
      { mode: 'includes', scope: scope });
  };

  /** The generated Fygaro link, read from the textarea or the Go To Link href. */
  locate.generatedLink = function (scope) {
    var areas = D.all('textarea', scope);
    for (var i = 0; i < areas.length; i++) {
      var value = U.normText(areas[i].value || areas[i].textContent);
      if (LINK_PATTERN.test(value)) return value;
    }
    var anchors = D.all('a[href*="/pb/"]', scope);
    for (var j = 0; j < anchors.length; j++) {
      var href = U.normText(anchors[j].getAttribute('href'));
      if (LINK_PATTERN.test(href)) return href;
    }
    return null;
  };

  /* --------------------------------------------------------------- helpers */

  function pause(settings) {
    return U.randomDelay(settings.minDelayMs, settings.maxDelayMs);
  }

  function fail(message) {
    throw new Error(message);
  }

  /** Waits until the app has actually landed on one of the given routes. */
  function waitForRoute(routes, timeoutMs) {
    var wanted = [].concat(routes);
    return D.waitFor(function () {
      var here = S.routeOf(location.pathname);
      return wanted.indexOf(here) !== -1 ? here : null;
    }, timeoutMs, 'the ' + wanted.join(' or ') + ' page');
  }

  /**
   * Waits for a save to resolve one of the two ways it can: the app navigates,
   * or the Code field complains that the code is already taken.
   *
   * One probe rather than two racing promises, because D.waitFor is already
   * driven by a MutationObserver and so sees the message the moment React
   * inserts it. A taken code therefore costs about a second instead of a full
   * step timeout, then a retry, then another full step timeout, then a human.
   *
   * The route is checked first, so a save that really did navigate can never be
   * misread as a rejection.
   */
  function waitForSaveOutcome(routes, timeoutMs, action) {
    var wanted = [].concat(routes);
    return D.waitFor(function () {
      var here = S.routeOf(location.pathname);
      if (wanted.indexOf(here) !== -1) return { landedRoute: here };
      var duplicate = STEPS.locate.duplicateCodeError();
      return duplicate ? { duplicate: duplicate } : null;
    }, timeoutMs, 'the ' + wanted.join(' or ') + ' page').catch(function () {
      var problems = D.collectErrors();
      fail(problems.length
        ? action + ' was rejected by the page: ' + problems.join(' | ')
        : action + ' did not lead anywhere. The page stayed on ' + location.pathname + '.');
    });
  }

  /**
   * Waits for a route, and if it never arrives reports what the page complained
   * about rather than only saying the navigation timed out.
   */
  function waitForRouteOrExplain(routes, timeoutMs, action) {
    return waitForRoute(routes, timeoutMs).catch(function () {
      var problems = D.collectErrors();
      fail(problems.length
        ? action + ' was rejected by the page: ' + problems.join(' | ')
        : action + ' did not lead anywhere. The page stayed on ' + location.pathname + '.');
    });
  }

  /**
   * Picks an option by its visible label, falling back to a known value.
   * Reading the label first survives Fygaro renumbering its option values, and
   * the fallback covers a relabelled interface.
   */
  function chooseOption(el, labels, fallbackValue) {
    var wanted = labels.map(U.foldText);
    for (var i = 0; i < el.options.length; i++) {
      if (wanted.indexOf(U.foldText(el.options[i].textContent)) !== -1) {
        return D.selectValue(el, el.options[i].value);
      }
    }
    return D.selectValue(el, fallbackValue);
  }

  /**
   * Sets every copy of a select. The product form renders the same controls
   * twice for its two responsive layouts, and only one copy is on screen.
   */
  function setAllSelects(name, labels, fallbackValue, humanName) {
    var all = D.all('select[name="' + name + '"]');
    if (!all.length) fail('The "' + humanName + '" dropdown was not found on the form.');

    var anySet = false;
    all.forEach(function (el) { if (chooseOption(el, labels, fallbackValue)) anySet = true; });
    if (!anySet) fail('The "' + humanName + '" dropdown has no option for ' + labels.join(' or ') + '.');

    var onScreen = all.filter(D.isVisible);
    var check = onScreen.length ? onScreen : all;
    var wanted = labels.map(U.foldText);
    var ok = check.every(function (el) {
      var opt = el.options[el.selectedIndex];
      return opt && (wanted.indexOf(U.foldText(opt.textContent)) !== -1 || opt.value === fallbackValue);
    });
    if (!ok) fail('The "' + humanName + '" dropdown did not keep the value that was selected.');
    return true;
  }

  /* ----------------------------------------------------------------- steps */

  /** 1. On the dashboard, open Products from the sidebar. */
  STEPS[S.STEP.NAV_TO_PRODUCTS] = function (job) {
    var link = locate.productsSidebarLink();
    if (!link) fail('The Products link was not found in the sidebar.');

    return pause(job.settings).then(function () {
      D.click(link);
      return waitForRouteOrExplain('productList', job.settings.stepTimeoutMs, 'Opening Products');
    }).then(function () { return {}; });
  };

  /** 2. On the products list, use the Create control in the top right. */
  STEPS[S.STEP.CLICK_CREATE] = function (job) {
    var target = locate.createItemLink();
    if (!target) fail('No Create control was found on the products page.');

    return pause(job.settings).then(function () {
      D.click(target);
      return waitForRouteOrExplain('productAdd', job.settings.stepTimeoutMs, 'Opening the new item form');
    }).then(function () { return {}; });
  };

  /** 3. Fill the product form and save it. */
  STEPS[S.STEP.FILL_PRODUCT] = function (job) {
    var row = job.row;
    var settings = job.settings;

    return D.waitForControl('input[name="name"]', settings.stepTimeoutMs, 'the new item form')
      .then(function () { return pause(settings); })
      .then(function () {
        var nameInput = D.control('input[name="name"]');
        var codeInput = D.control('input[name="code"]');
        var priceInput = D.control('input[name="price"]');

        if (!nameInput) fail('The Name field was not found on the item form.');
        if (!codeInput) fail('The Code field was not found on the item form.');
        if (!priceInput) fail('The Price field was not found on the item form.');

        if (!D.fillText(nameInput, row.name)) fail('The Name field did not keep the value that was typed.');
        if (!D.fillText(codeInput, row.code)) fail('The Code field did not keep the value that was typed.');
        if (!D.fillText(priceInput, row.priceText)) fail('The Price field did not keep the value that was typed.');

        setAllSelects('currency', ['USD'], '2', 'Currency');
        setAllSelects('product_type', ['Service', 'Servicio'], '2', 'Type of Item');
        setAllSelects('show_in_website', ['No'], 'false', 'Show In Website');

        return pause(settings);
      })
      .then(function () {
        // A dry run stops with the form filled, so nothing is ever created. The
        // Code is still inspected, because Fygaro may have checked it when the
        // field lost focus, and knowing early is useful. It is only reported.
        if (settings.dryRun) return { dryRun: true, duplicate: locate.duplicateCodeError() };

        // Fygaro can refuse the Code before Save is ever pressed. Nothing has
        // been created in that case, so the row is reported without saving.
        var already = locate.duplicateCodeError();
        if (already) return { duplicate: already };

        var save = locate.saveButton();
        if (!save) fail('The Save button was not found on the item form.');

        D.click(save);
        return waitForSaveOutcome(['productList', 'productDetail'], settings.stepTimeoutMs, 'Saving the item')
          .then(function (outcome) {
            if (outcome.duplicate) return { duplicate: outcome.duplicate };
            return {
              landedRoute: outcome.landedRoute,
              productUuid: outcome.landedRoute === 'productDetail' ? U.extractUuid(location.pathname) : ''
            };
          });
      });
  };

  /**
   * 4. Open the product that was just created.
   * Matched on name plus code rather than on being newest, so a catalog with
   * repeated names cannot open the wrong record.
   */
  STEPS[S.STEP.OPEN_PRODUCT] = function (job) {
    var row = job.row;

    return D.waitFor(function () { return locate.productListItem(null, row.name, row.code); },
      job.settings.stepTimeoutMs, 'the product "' + U.truncate(row.name, 60) + '" in the list')
      .catch(function () {
        fail('The product "' + U.truncate(row.name, 80) + '" (code ' + row.code +
          ') did not appear in the products list, so it may not have saved.');
      })
      .then(function (anchor) {
        var uuid = U.extractUuid(anchor.getAttribute('href'));
        return pause(job.settings).then(function () {
          D.click(anchor);
          return waitForRouteOrExplain('productDetail', job.settings.stepTimeoutMs, 'Opening the product');
        }).then(function () {
          return { productUuid: uuid };
        });
      });
  };

  /** 5. Confirm the right product is open, then start its Fygaro Link. */
  STEPS[S.STEP.CLICK_CREATE_LINK] = function (job) {
    var row = job.row;

    return D.waitForControl('input[name="name"]', job.settings.stepTimeoutMs, 'the product page')
      .then(function () {
        if (row.productUuid) {
          var here = U.extractUuid(location.pathname);
          if (here && here !== row.productUuid) {
            fail('The open product is ' + here + ' but row ' + row.sheetRow + ' created ' + row.productUuid + '.');
          }
        }

        var nameInput = D.control('input[name="name"]');
        var codeInput = D.control('input[name="code"]');

        if (U.foldText(nameInput ? nameInput.value : '') !== U.foldText(row.name)) {
          fail('The open product is named "' + U.truncate(nameInput ? nameInput.value : '', 60) +
            '" but row ' + row.sheetRow + ' expects "' + U.truncate(row.name, 60) + '".');
        }
        if (row.code && U.foldText(codeInput ? codeInput.value : '') !== U.foldText(row.code)) {
          fail('The open product has code "' + (codeInput ? codeInput.value : '') +
            '" but row ' + row.sheetRow + ' expects "' + row.code + '".');
        }

        var target = locate.createLinkAction();
        if (!target) fail('The "Create Fygaro Link for Product" button was not found.');

        return pause(job.settings).then(function () {
          D.click(target);
          return waitForRouteOrExplain('linkAdd', job.settings.stepTimeoutMs, 'Starting the Fygaro Link');
        });
      })
      .then(function () { return {}; });
  };

  /** 6. Name the link after the code, tick the required checkout fields, save. */
  STEPS[S.STEP.FILL_LINK] = function (job) {
    var row = job.row;
    var settings = job.settings;

    var REQUIRED = [
      { name: 'require_phone', label: 'Phone' },
      { name: 'require_legal_id', label: 'Legal ID' },
      { name: 'require_billing_address', label: 'Billing Address' }
    ];

    return D.waitForControl('input[name="name"]', settings.stepTimeoutMs, 'the new link form')
      .then(function () { return pause(settings); })
      .then(function () {
        // Fygaro attaches the product itself. Confirm it attached the right one.
        if (row.code && U.foldText(document.body.textContent).indexOf(U.foldText(row.code)) === -1) {
          fail('The link form does not show product code ' + row.code + ', so the wrong product may be attached.');
        }

        var nameInput = D.control('input[name="name"]');
        if (!nameInput) fail('The Name field was not found on the link form.');
        if (!D.fillText(nameInput, row.code)) fail('The link Name field did not keep the value that was typed.');

        setAllSelects('currency', ['USD'], '2', 'Currency');
        return pause(settings);
      })
      .then(function () {
        // The toggle flips between Show and Hide, so decide from the panel
        // itself rather than from the button label. Presence is the signal, not
        // visibility: these inputs are hidden behind a painted replacement.
        if (locate.requirementCheckbox(null, 'require_phone')) return null;

        var toggle = locate.advancedOptionsToggle();
        if (!toggle) fail('The "Show Advanced Options" button was not found on the link form.');
        D.click(toggle);

        return D.waitFor(function () { return locate.requirementCheckbox(null, 'require_phone'); },
          settings.stepTimeoutMs, 'the Advanced Options panel')
          .catch(function () {
            // Name what is actually on the form, so a future change to it is
            // obvious from the error rather than needing a live debug session.
            var names = D.all('input[type="checkbox"]').map(function (el) { return el.name; })
              .filter(Boolean);
            fail('The Advanced Options panel did not open. The checkboxes on the form are: ' +
              (names.length ? names.join(', ') : 'none') + '.');
          });
      })
      .then(function () { return pause(settings); })
      .then(function () {
        for (var i = 0; i < REQUIRED.length; i++) {
          var box = locate.requirementCheckbox(null, REQUIRED[i].name);
          if (!box) {
            fail('The "' + REQUIRED[i].label + '" checkbox (' + REQUIRED[i].name +
              ') was not found in Advanced Options.');
          }
          if (!D.setChecked(box, true)) {
            fail('The "' + REQUIRED[i].label + '" checkbox would not tick. It was clicked directly and ' +
              'through its label, and stayed ' + (box.checked ? 'ticked' : 'unticked') + '.');
          }
        }
        return pause(settings);
      })
      .then(function () {
        var save = locate.saveButton();
        if (!save) fail('The Save button was not found on the link form.');

        D.click(save);
        return waitForRouteOrExplain('linkDone', settings.stepTimeoutMs, 'Saving the Fygaro Link');
      })
      .then(function () { return {}; });
  };

  /** 7. Read the finished link, then go back to the dashboard for the next row. */
  STEPS[S.STEP.CAPTURE_LINK] = function (job) {
    var settings = job.settings;

    return D.waitFor(function () { return locate.generatedLink(); }, settings.stepTimeoutMs,
      'the generated Fygaro link')
      .catch(function () { fail('The Fygaro link did not appear on the confirmation page.'); })
      .then(function (link) {
        // The link is read from the page, not from the clipboard. Copy Link is
        // still pressed so the page ends in the state a user would expect.
        var copy = locate.copyLinkButton();
        if (copy) { try { D.click(copy); } catch (e) { /* purely cosmetic */ } }

        return pause(settings).then(function () {
          var home = locate.backHomeLink();
          if (!home) fail('The "Back to Home" link was not found in the sidebar.');
          D.click(home);
          return waitForRouteOrExplain('dashboard', settings.stepTimeoutMs, 'Going back to the dashboard');
        }).then(function () {
          return { link: link };
        });
      });
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
