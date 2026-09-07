/*
 * Fygaro Catalog Automation
 * Brings an oversized product picture under Fygaro's upload limit.
 *
 * Fygaro refuses anything over 2.5 MB with "Image size exceeds 2.5MB", counting
 * a megabyte as 1,000,000 bytes rather than 1,048,576. One picture in this
 * catalog is 2,525,525 bytes, which is over that line and under the binary one,
 * which is exactly why it failed while everything else went through.
 *
 * Only pictures above the target are touched. Re-encoding costs quality, the
 * originals were chosen deliberately, and 32 of the 35 in this catalog are
 * comfortably inside the limit, so they are stored byte for byte as they are.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var F = FYG.imagefit || (FYG.imagefit = {});

  /** What Fygaro refuses. Decimal megabytes, as its own message counts them. */
  F.LIMIT_BYTES = 2500000;

  /*
   * What we aim for. The margin matters because the server may be measuring the
   * encoded upload rather than the file, and being told a picture is too large
   * after a product has been created is far worse than a slightly smaller photo.
   */
  F.TARGET_BYTES = 2250000;

  /* Tried in order. Quality first, because scaling loses detail permanently. */
  var QUALITIES = [0.92, 0.85, 0.78, 0.7, 0.6];
  var SCALES = [1, 0.8, 0.65, 0.5, 0.4];

  function canDecode() {
    return typeof createImageBitmap === 'function' &&
      (typeof OffscreenCanvas === 'function' || typeof document !== 'undefined');
  }

  function draw(bitmap, scale) {
    var width = Math.max(1, Math.round(bitmap.width * scale));
    var height = Math.max(1, Math.round(bitmap.height * scale));

    var canvas;
    if (typeof OffscreenCanvas === 'function') {
      canvas = new OffscreenCanvas(width, height);
    } else {
      canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
    }

    var ctx = canvas.getContext('2d');
    // JPEG has no alpha. Without this a transparent PNG comes out on black.
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, width, height);
    ctx.drawImage(bitmap, 0, 0, width, height);
    return canvas;
  }

  function encode(canvas, quality) {
    if (canvas.convertToBlob) {
      return canvas.convertToBlob({ type: 'image/jpeg', quality: quality });
    }
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (blob) {
        if (blob) resolve(blob);
        else reject(new Error('The picture could not be re-encoded.'));
      }, 'image/jpeg', quality);
    });
  }

  function jpegName(name) {
    return String(name || 'image').replace(/\.[^.]+$/, '') + '.jpg';
  }

  /**
   * Returns a picture guaranteed to be within the target, re-encoding it only if
   * it has to.
   *
   * @param {Uint8Array} bytes
   * @param {string} type
   * @param {string} name
   * @returns {Promise<{bytes, type, name, changed: boolean, from: number, to: number, note: string}>}
   */
  F.fit = function (bytes, type, name) {
    var unchanged = {
      bytes: bytes, type: type, name: name, changed: false,
      from: bytes.length, to: bytes.length, note: ''
    };

    if (bytes.length <= F.TARGET_BYTES) return Promise.resolve(unchanged);

    if (!canDecode()) {
      unchanged.note = 'This browser cannot re-encode pictures, so it was left at its original size.';
      return Promise.resolve(unchanged);
    }

    var bitmap = null;

    return createImageBitmap(new Blob([bytes], { type: type || 'image/png' }))
      .then(function (decoded) {
        bitmap = decoded;

        // Widen the search only as far as it has to go: every quality at one
        // scale before losing any more pixels.
        var attempts = [];
        SCALES.forEach(function (scale) {
          QUALITIES.forEach(function (quality) { attempts.push({ scale: scale, quality: quality }); });
        });

        var smallest = null;

        function tryNext(at) {
          if (at >= attempts.length) return Promise.resolve(smallest);
          var attempt = attempts[at];
          return encode(draw(bitmap, attempt.scale), attempt.quality).then(function (blob) {
            if (!smallest || blob.size < smallest.size) smallest = blob;
            if (blob.size <= F.TARGET_BYTES) return blob;
            return tryNext(at + 1);
          });
        }

        return tryNext(0);
      })
      .then(function (blob) {
        if (!blob) throw new Error('No re-encoding of the picture came out small enough.');
        return blob.arrayBuffer().then(function (buffer) {
          var out = new Uint8Array(buffer);

          // A re-encode that made it bigger is not worth the quality it cost.
          if (out.length >= bytes.length) return unchanged;

          return {
            bytes: out,
            type: 'image/jpeg',
            name: jpegName(name),
            changed: true,
            from: bytes.length,
            to: out.length,
            note: out.length <= F.TARGET_BYTES ? ''
              : 'It is still ' + Math.round(out.length / 1000) + ' KB, which Fygaro may refuse.'
          };
        });
      })
      .catch(function (err) {
        // Better to try the original and be told no than to lose the picture.
        unchanged.note = 'It could not be resized (' + (err && err.message ? err.message : err) +
          '), so it was left at its original size.';
        return unchanged;
      })
      .then(function (result) {
        if (bitmap && bitmap.close) bitmap.close();
        return result;
      });
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
