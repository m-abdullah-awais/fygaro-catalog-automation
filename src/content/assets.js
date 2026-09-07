/*
 * Fygaro Catalog Automation
 * Fetches product pictures from the worker, and remembers a few of them.
 *
 * The catalog reuses 35 pictures across 1030 rows, in tight runs of three or
 * four consecutive rows sharing one image. A small cache therefore removes
 * nearly all the traffic: walking the real sheet in order, one entry still costs
 * 932 fetches, two costs 771, and four costs 39. Six is used here, which is
 * about 12 MB held in a page the user is also working in, and within a handful
 * of the 35 fetches that perfect caching would need.
 *
 * Caching the promise rather than the file is deliberate: several callers asking
 * for the same picture at once share one round trip.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var A = FYG.assets || (FYG.assets = {});
  var U = FYG.util;
  var S = FYG.state;

  A.MAX_ENTRIES = 6;

  var cache = new Map();
  var hits = 0;
  var misses = 0;

  function ask(id, part) {
    return chrome.runtime.sendMessage({ type: S.MSG.REQUEST_IMAGE, id: id, part: part })
      .then(function (reply) {
        if (!reply) throw new Error('The extension did not answer for ' + id + '.');
        if (!reply.ok) throw new Error(reply.error || 'The picture could not be read.');
        return reply;
      });
  }

  /** Pulls every chunk in order and rebuilds the file. */
  function fetchWhole(id) {
    return ask(id, 0).then(function (first) {
      var chunks = [first.data];

      function next(part) {
        if (part >= first.parts) return Promise.resolve();
        // One at a time, so a slow disk cannot leave several megabyte sized
        // messages in flight at once.
        return ask(id, part).then(function (reply) {
          chunks.push(reply.data);
          return next(part + 1);
        });
      }

      return next(1).then(function () {
        var bytes = U.base64ToBytes(chunks.join(''));
        if (first.size && bytes.length !== first.size) {
          throw new Error('The picture arrived incomplete: ' + bytes.length + ' of ' + first.size + ' bytes.');
        }

        /*
         * A last check before it reaches the form. The panel sizes pictures when
         * the catalog is loaded, but a catalog stored by an older build holds
         * whatever was written then, and Fygaro refusing the upload halfway
         * through a run measured in hours is an expensive way to find out.
         * Almost always this does nothing, because the bytes are already small.
         */
        return FYG.imagefit.fit(bytes, first.type, first.name).then(function (fitted) {
          return new File([fitted.bytes], fitted.name, { type: fitted.type });
        });
      });
    });
  }

  /**
   * The picture for one catalog row, as a File ready for a file input.
   * @returns {Promise<File>}
   */
  A.fetchImage = function (id) {
    if (!id) return Promise.reject(new Error('No picture was asked for.'));

    if (cache.has(id)) {
      hits++;
      return cache.get(id);
    }
    misses++;

    var pending = fetchWhole(id).catch(function (err) {
      // A failure must not be remembered, or one hiccup would poison every
      // later row that shares this picture.
      cache.delete(id);
      throw err;
    });

    cache.set(id, pending);
    while (cache.size > A.MAX_ENTRIES) {
      cache.delete(cache.keys().next().value);
    }
    return pending;
  };

  A.clear = function () {
    cache.clear();
    hits = 0;
    misses = 0;
  };

  /** Exposed so the tests can prove the cache is doing its job. */
  A.stats = function () {
    return { hits: hits, misses: misses, entries: cache.size };
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
