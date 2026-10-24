// SK Hub Office の ZIP（window.OfficeZip）。Office のファイル（.docx・.xlsx・.pptx）は、中身が XML の ZIP なので、その読み書きに使う
//   ・OfficeZip.read(blob か ArrayBuffer) → Promise<{ "word/document.xml": Uint8Array, … }>
//   ・OfficeZip.write({ "名前": 文字列か Uint8Array, … }) → Promise<Blob>
//   ・OfficeZip.text(Uint8Array) → 文字列（UTF-8）
//   縮める・戻すのは、ブラウザの CompressionStream("deflate-raw")。使えないブラウザでは、縮めずに入れる（読むほうは使えないと読めない）
(function () {
  "use strict";
  var enc = new TextEncoder(), dec = new TextDecoder("utf-8");
  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(u8) { var c = 0xffffffff; for (var i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
  var canDeflate = typeof window.CompressionStream === "function";
  var canInflate = typeof window.DecompressionStream === "function";
  function pipe(u8, stream) {
    return new Response(new Blob([u8]).stream().pipeThrough(stream)).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }
  function deflate(u8) { return canDeflate ? pipe(u8, new CompressionStream("deflate-raw")).catch(function () { return null; }) : Promise.resolve(null); }
  function inflate(u8) { return pipe(u8, new DecompressionStream("deflate-raw")); }

  function read(src) {
    var buf = src instanceof ArrayBuffer ? Promise.resolve(src) : src.arrayBuffer();
    return buf.then(function (ab) {
      var u8 = new Uint8Array(ab), dv = new DataView(ab), i, out = {}, jobs = [];
      // 終わりの目録（End of central directory）を後ろから探す
      for (i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) break;
      if (i < 0 || dv.getUint32(i, true) !== 0x06054b50) throw new Error("not-zip");
      var count = dv.getUint16(i + 10, true), p = dv.getUint32(i + 16, true);
      for (var n = 0; n < count; n++) {
        if (dv.getUint32(p, true) !== 0x02014b50) break;
        var method = dv.getUint16(p + 10, true), size = dv.getUint32(p + 20, true);
        var nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
        var local = dv.getUint32(p + 42, true);
        var name = dec.decode(u8.subarray(p + 46, p + 46 + nameLen));
        var start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
        var data = u8.subarray(start, start + size);
        if (name.slice(-1) !== "/") {
          if (method === 0) out[name] = data;
          else if (method === 8) {
            if (!canInflate) throw new Error("no-inflate");
            jobs.push((function (nm, d) { return inflate(d).then(function (r) { out[nm] = r; }); })(name, data));
          }
        }
        p += 46 + nameLen + extraLen + commentLen;
      }
      return Promise.all(jobs).then(function () { return out; });
    });
  }

  function write(files) {
    var names = Object.keys(files);
    return Promise.all(names.map(function (nm) {
      var raw = typeof files[nm] === "string" ? enc.encode(files[nm]) : files[nm];
      return deflate(raw).then(function (z) {
        var use = z && z.length < raw.length;
        return { name: enc.encode(nm), raw: raw, data: use ? z : raw, method: use ? 8 : 0, crc: crc32(raw) };
      });
    })).then(function (list) {
      var parts = [], central = [], offset = 0;
      var d = new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
      var date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
      list.forEach(function (f) {
        var h = new DataView(new ArrayBuffer(30));
        h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, f.method, true);
        h.setUint16(10, time, true); h.setUint16(12, date, true); h.setUint32(14, f.crc, true);
        h.setUint32(18, f.data.length, true); h.setUint32(22, f.raw.length, true); h.setUint16(26, f.name.length, true);
        parts.push(new Uint8Array(h.buffer), f.name, f.data);
        var c = new DataView(new ArrayBuffer(46));
        c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, f.method, true);
        c.setUint16(12, time, true); c.setUint16(14, date, true); c.setUint32(16, f.crc, true);
        c.setUint32(20, f.data.length, true); c.setUint32(24, f.raw.length, true); c.setUint16(28, f.name.length, true);
        c.setUint32(42, offset, true);
        central.push(new Uint8Array(c.buffer), f.name);
        offset += 30 + f.name.length + f.data.length;
      });
      var size = central.reduce(function (a, b) { return a + b.length; }, 0);
      var e = new DataView(new ArrayBuffer(22));
      e.setUint32(0, 0x06054b50, true); e.setUint16(8, list.length, true); e.setUint16(10, list.length, true);
      e.setUint32(12, size, true); e.setUint32(16, offset, true);
      return new Blob(parts.concat(central, [new Uint8Array(e.buffer)]), { type: "application/zip" });
    });
  }

  window.OfficeZip = {
    read: read, write: write, canRead: canInflate,
    text: function (u8) { return u8 ? dec.decode(u8) : ""; },
    xml: function (u8) { return u8 ? new DOMParser().parseFromString(dec.decode(u8), "application/xml") : null; },
    esc: function (s) { return String(s === undefined || s === null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ""); }
  };
})();
