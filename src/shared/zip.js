/*
 * Fygaro Catalog Automation
 * Minimal ZIP reader and writer, which is all an xlsx really needs.
 *
 * Deflate is handled by the browser's native CompressionStream and
 * DecompressionStream, so this file has no dependencies at all.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var ZIP = FYG.zip || (FYG.zip = {});

  var SIG_LOCAL = 0x04034b50;
  var SIG_CENTRAL = 0x02014b50;
  var SIG_EOCD = 0x06054b50;

  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    return table;
  })();

  ZIP.crc32 = function (bytes) {
    var c = 0xffffffff;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };

  function inflateRaw(bytes) {
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Response(stream).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }

  function deflateRaw(bytes) {
    var stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
    return new Response(stream).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }

  ZIP.inflateRaw = inflateRaw;
  ZIP.deflateRaw = deflateRaw;

  /**
   * Reads a zip archive.
   * @param {ArrayBuffer|Uint8Array} input
   * @returns {Promise<{order: string[], files: Map<string, Uint8Array>}>}
   *   `order` preserves the original entry order so the rewritten archive keeps
   *   the same layout, which some readers are sensitive to.
   */
  ZIP.read = function (input) {
    var bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

    // The end of central directory record sits at the tail, after an optional
    // comment of up to 64 KB, so scan backwards for its signature.
    var eocd = -1;
    var minStart = Math.max(0, bytes.length - 22 - 0xffff);
    for (var i = bytes.length - 22; i >= minStart; i--) {
      if (view.getUint32(i, true) === SIG_EOCD) { eocd = i; break; }
    }
    if (eocd < 0) return Promise.reject(new Error('Not a zip file: no end of central directory record found.'));

    var count = view.getUint16(eocd + 10, true);
    var cdOffset = view.getUint32(eocd + 16, true);

    var order = [];
    var jobs = [];
    var files = new Map();
    var p = cdOffset;
    var decoder = new TextDecoder('utf-8');

    function readEntry(name, method, raw) {
      if (method === 0) { files.set(name, raw.slice()); return Promise.resolve(); }
      if (method !== 8) return Promise.reject(new Error('Unsupported compression method ' + method + ' for ' + name + '.'));
      return inflateRaw(raw).then(function (out) { files.set(name, out); });
    }

    for (var n = 0; n < count; n++) {
      if (view.getUint32(p, true) !== SIG_CENTRAL) {
        return Promise.reject(new Error('Corrupt zip: bad central directory entry at ' + p + '.'));
      }
      var method = view.getUint16(p + 10, true);
      var compSize = view.getUint32(p + 20, true);
      var nameLen = view.getUint16(p + 28, true);
      var extraLen = view.getUint16(p + 30, true);
      var commentLen = view.getUint16(p + 32, true);
      var localOffset = view.getUint32(p + 42, true);
      var name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen));

      if (view.getUint32(localOffset, true) !== SIG_LOCAL) {
        return Promise.reject(new Error('Corrupt zip: bad local header for ' + name + '.'));
      }
      var lNameLen = view.getUint16(localOffset + 26, true);
      var lExtraLen = view.getUint16(localOffset + 28, true);
      var dataStart = localOffset + 30 + lNameLen + lExtraLen;
      var raw = bytes.subarray(dataStart, dataStart + compSize);

      order.push(name);
      jobs.push(readEntry(name, method, raw));
      p += 46 + nameLen + extraLen + commentLen;
    }

    return Promise.all(jobs).then(function () { return { order: order, files: files }; });
  };

  /**
   * Writes a zip archive.
   * @param {string[]} order entry names, in the order they should be stored
   * @param {Map<string, Uint8Array>} files
   * @returns {Promise<Blob>}
   */
  ZIP.write = function (order, files) {
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    var encoder = new TextEncoder();

    var prepared = order.map(function (name) {
      var data = files.get(name);
      if (!data) throw new Error('Missing zip entry content for ' + name + '.');
      var nameBytes = encoder.encode(name);
      return deflateRaw(data).then(function (deflated) {
        // Fall back to storing when deflate does not actually help.
        var useDeflate = deflated.length < data.length;
        return {
          nameBytes: nameBytes,
          // Bit 11 marks the file name as UTF-8, which matters for accented names.
          flags: 0x0800,
          method: useDeflate ? 8 : 0,
          crc: ZIP.crc32(data),
          body: useDeflate ? deflated : data,
          size: data.length
        };
      });
    });

    return Promise.all(prepared).then(function (entries) {
      var localSize = 0;
      var centralSize = 0;
      entries.forEach(function (e) {
        localSize += 30 + e.nameBytes.length + e.body.length;
        centralSize += 46 + e.nameBytes.length;
      });

      var out = new Uint8Array(localSize + centralSize + 22);
      var dv = new DataView(out.buffer);
      var offset = 0;
      var offsets = [];

      entries.forEach(function (e) {
        offsets.push(offset);
        dv.setUint32(offset, SIG_LOCAL, true);
        dv.setUint16(offset + 4, 20, true);
        dv.setUint16(offset + 6, e.flags, true);
        dv.setUint16(offset + 8, e.method, true);
        dv.setUint16(offset + 10, dosTime, true);
        dv.setUint16(offset + 12, dosDate, true);
        dv.setUint32(offset + 14, e.crc, true);
        dv.setUint32(offset + 18, e.body.length, true);
        dv.setUint32(offset + 22, e.size, true);
        dv.setUint16(offset + 26, e.nameBytes.length, true);
        dv.setUint16(offset + 28, 0, true);
        out.set(e.nameBytes, offset + 30);
        out.set(e.body, offset + 30 + e.nameBytes.length);
        offset += 30 + e.nameBytes.length + e.body.length;
      });

      var cdStart = offset;
      entries.forEach(function (e, i) {
        dv.setUint32(offset, SIG_CENTRAL, true);
        dv.setUint16(offset + 4, 20, true);
        dv.setUint16(offset + 6, 20, true);
        dv.setUint16(offset + 8, e.flags, true);
        dv.setUint16(offset + 10, e.method, true);
        dv.setUint16(offset + 12, dosTime, true);
        dv.setUint16(offset + 14, dosDate, true);
        dv.setUint32(offset + 16, e.crc, true);
        dv.setUint32(offset + 20, e.body.length, true);
        dv.setUint32(offset + 24, e.size, true);
        dv.setUint16(offset + 28, e.nameBytes.length, true);
        dv.setUint16(offset + 30, 0, true);
        dv.setUint16(offset + 32, 0, true);
        dv.setUint16(offset + 34, 0, true);
        dv.setUint16(offset + 36, 0, true);
        dv.setUint32(offset + 38, 0, true);
        dv.setUint32(offset + 42, offsets[i], true);
        out.set(e.nameBytes, offset + 46);
        offset += 46 + e.nameBytes.length;
      });

      dv.setUint32(offset, SIG_EOCD, true);
      dv.setUint16(offset + 4, 0, true);
      dv.setUint16(offset + 6, 0, true);
      dv.setUint16(offset + 8, entries.length, true);
      dv.setUint16(offset + 10, entries.length, true);
      dv.setUint32(offset + 12, offset - cdStart, true);
      dv.setUint32(offset + 16, cdStart, true);
      dv.setUint16(offset + 20, 0, true);

      return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    });
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
