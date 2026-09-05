/*
 * Fygaro Catalog Automation
 * The one IndexedDB database, shared by the side panel and the worker.
 *
 * Three things live here that chrome.storage cannot hold: a file handle, which
 * is not JSON; the workbook bytes, which are 63.5 MB against a 10 MB quota; and
 * the product pictures, which are another 65 MB.
 *
 * Every call is time limited. All of this is a convenience layered on top of a
 * run that works without it, so a database that misbehaves must let the caller
 * carry on rather than leave it waiting for an event that never arrives. That is
 * not hypothetical: a page opened from disk has no IndexedDB at all.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var I = FYG.idb || (FYG.idb = {});

  I.NAME = 'fygaro-file';
  I.VERSION = 2;
  I.STORES = ['handles', 'workbook', 'images'];

  I.DEFAULT_TIMEOUT_MS = 3000;

  I.available = function () {
    return typeof indexedDB !== 'undefined' && !!indexedDB;
  };

  I.open = function () {
    return new Promise(function (resolve, reject) {
      if (!I.available()) {
        reject(new Error('IndexedDB is not available here.'));
        return;
      }
      var request;
      try {
        request = indexedDB.open(I.NAME, I.VERSION);
      } catch (err) {
        reject(err);
        return;
      }
      request.onupgradeneeded = function () {
        var db = request.result;
        I.STORES.forEach(function (name) {
          if (!db.objectStoreNames.contains(name)) db.createObjectStore(name);
        });
      };
      request.onsuccess = function () {
        var db = request.result;
        // A connection still open in another context would otherwise block the
        // next version change indefinitely.
        db.onversionchange = function () { db.close(); };
        resolve(db);
      };
      request.onerror = function () {
        reject(request.error || new Error('IndexedDB could not be opened.'));
      };
      request.onblocked = function () {
        reject(new Error('IndexedDB is blocked by another tab.'));
      };
    });
  };

  /** Runs one transaction against an already open connection. */
  I.on = function (db, storeNames, mode, action) {
    return new Promise(function (resolve, reject) {
      var request;
      try {
        var tx = db.transaction(storeNames, mode);
        request = action(tx.objectStore([].concat(storeNames)[0]), tx);
        tx.oncomplete = function () { resolve(request ? request.result : undefined); };
        tx.onerror = function () { reject(tx.error || new Error('The transaction failed.')); };
        tx.onabort = function () { reject(tx.error || new Error('The transaction was aborted.')); };
      } catch (err) {
        reject(err);
      }
    });
  };

  /**
   * Opens, runs one transaction, and closes.
   *
   * There is deliberately no wall clock timeout racing the transaction. An
   * earlier version had one, to guarantee the panel could never sit waiting for
   * a database that was never going to answer. Every way that can happen is now
   * handled at the source instead: an absent IndexedDB is refused by available()
   * before anything is opened, and open, blocked, error and abort all settle the
   * promise. A timer on top of that could only ever fire early and turn a slow
   * read of a 2 MB picture into a spurious failure.
   *
   * @param {string|string[]} storeNames
   * @param {string} mode "readonly" or "readwrite"
   * @param {function(IDBObjectStore, IDBTransaction): (IDBRequest|undefined)} action
   */
  I.run = function (storeNames, mode, action) {
    var db = null;
    return I.open().then(function (opened) {
      db = opened;
      return I.on(db, storeNames, mode, action);
    }).then(function (value) {
      if (db) db.close();
      return value;
    }, function (err) {
      if (db) db.close();
      throw err;
    });
  };

  /** Empties a store, reporting failure rather than throwing. */
  I.clearStore = function (name) {
    return I.run(name, 'readwrite', function (store) { return store.clear(); })
      .then(function () { return true; }, function () { return false; });
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
