// SK Hub Office のファイルの保存（window.OfficeStore）
//   ・PC のフォルダに 1 ファイルずつ保存する（File System Access API。Chrome・Edge のパソコン）
//       最初に 1 回だけ、保存するフォルダを選んでもらう（showDirectoryPicker）。選んだフォルダは IndexedDB の sk_office に覚えておく
//       次に開いたときにブラウザが許可を聞き直すことがあるので、そのときは「許可する」を押してもらう（allow()）
//       ファイル名は「名前.skdoc」「名前.sksheet」「名前.skslide」。中身は JSON：
//         { format: "sk-hub-office", version: 1, app, tpl, created, data }（data の形は templates.js の説明のとおり）
//       ファイルの ID はフォルダの中のファイル名
//   ・フォルダに保存できないブラウザ（Safari・Firefox・スマホなど）は、ブラウザの中（localStorage の sk_office_files）に保存する
//       [{ id, app, name, tpl, created, updated, data }]。ファイルとして書き出す（exportFile）・読み込む（importFile）もできる
//   ・前にブラウザの中に保存したファイルは、フォルダを選んだあとに migrate() でフォルダに移せる
//   どれも Promise を返す（state()・folderName()・supported だけはすぐ返す）
(function () {
  "use strict";
  var EXT = { doc: ".skdoc", sheet: ".sksheet", slide: ".skslide" };
  var FORMAT = "sk-hub-office";
  var LS_KEY = "sk_office_files";
  var supported = typeof window.showDirectoryPicker === "function" && window.isSecureContext;
  var folder = null, state = supported ? "none" : "unsupported";

  function appOf(fileName) {
    for (var a in EXT) if (fileName.toLowerCase().slice(-EXT[a].length) === EXT[a]) return a;
    return null;
  }
  function baseName(fileName) { var a = appOf(fileName); return a ? fileName.slice(0, -EXT[a].length) : fileName; }
  function payload(f) { return JSON.stringify({ format: FORMAT, version: 1, app: f.app, tpl: f.tpl || "", created: f.created || Date.now(), data: f.data }); }
  function parse(text, fallbackApp) {
    var o = JSON.parse(text);
    if (!o || o.format !== FORMAT || !o.data) throw new Error("not-office");
    return { app: o.app || fallbackApp, tpl: o.tpl || "", created: o.created || Date.now(), data: o.data };
  }

  // ---------- IndexedDB（選んだフォルダを覚えておく） ----------
  function idb(mode, fn) {
    return new Promise(function (resolve, reject) {
      if (!window.indexedDB) return reject(new Error("no-idb"));
      var req = indexedDB.open("sk_office", 1);
      req.onupgradeneeded = function () { req.result.createObjectStore("kv"); };
      req.onerror = function () { reject(req.error); };
      req.onsuccess = function () {
        var db = req.result, tx = db.transaction("kv", mode), r = fn(tx.objectStore("kv"));
        tx.oncomplete = function () { db.close(); resolve(r && r.result); };
        tx.onerror = function () { db.close(); reject(tx.error); };
      };
    });
  }
  function idbGet(k) { return idb("readonly", function (s) { return s.get(k); }); }
  function idbSet(k, v) { return idb("readwrite", function (s) { return s.put(v, k); }); }

  // ---------- ブラウザの中（localStorage） ----------
  function lsAll() { try { var a = JSON.parse(localStorage.getItem(LS_KEY) || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function lsSave(list) {
    try { localStorage.setItem(LS_KEY, JSON.stringify(list)); }
    catch (e) { throw new Error("full"); }
  }
  function newId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  // ---------- フォルダ ----------
  function ready() { return state === "ready" && folder; }
  function clean(name) {
    var s = String(name || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").replace(/^[.\s]+|[.\s]+$/g, "").slice(0, 100);
    return s || "無題";
  }
  function exists(fileName) {
    return folder.getFileHandle(fileName).then(function () { return true; }, function () { return false; });
  }
  // 同じ名前があるときは「名前 (2)」のようにする（except は自分自身の名前。名前を変えないときに使う）
  function uniqueName(name, app, except) {
    var base = clean(name), ext = EXT[app], i = 1;
    function next() {
      var cand = (i === 1 ? base : base + " (" + i + ")") + ext;
      if (cand === except) return Promise.resolve(cand);
      return exists(cand).then(function (yes) { if (!yes) return cand; i++; return next(); });
    }
    return next();
  }
  function write(fileName, text) {
    return folder.getFileHandle(fileName, { create: true }).then(function (h) {
      return h.createWritable().then(function (w) { return w.write(text).then(function () { return w.close(); }); });
    });
  }
  // 書き込みが重ならないように、順番に行う
  var queue = Promise.resolve();
  function inOrder(fn) { var p = queue.then(fn, fn); queue = p.catch(function () {}); return p; }

  var Store = {
    EXT: EXT,
    supported: supported,
    appOf: appOf,
    state: function () { return state; },
    folderName: function () { return folder ? folder.name : ""; },

    // 前に選んだフォルダを読み出して、許可があるか確かめる
    init: function () {
      if (!supported) return Promise.resolve(state);
      return idbGet("folder").then(function (h) {
        if (!h) { state = "none"; return state; }
        folder = h;
        return h.queryPermission({ mode: "readwrite" }).then(function (p) { state = p === "granted" ? "ready" : "needs-permission"; return state; });
      }).catch(function () { state = "none"; return state; });
    },
    // フォルダを選ぶ（ボタンを押したときに呼ぶ）。やめたら false
    pick: function () {
      return window.showDirectoryPicker({ id: "sk-hub-office", mode: "readwrite", startIn: "documents" }).then(function (h) {
        folder = h; state = "ready";
        return idbSet("folder", h).catch(function () {}).then(function () { return true; });
      }, function () { return false; });
    },
    // 前に選んだフォルダを使う許可をもらう（ボタンを押したときに呼ぶ）
    allow: function () {
      if (!folder) return Promise.resolve(false);
      return folder.requestPermission({ mode: "readwrite" }).then(function (p) { state = p === "granted" ? "ready" : "needs-permission"; return state === "ready"; }, function () { return false; });
    },

    // ファイルの一覧（中身は読まない）: [{ id, app, name, updated, size }]
    list: function () {
      if (!supported) return Promise.resolve(lsAll().map(function (f) { return { id: f.id, app: f.app, name: f.name, updated: f.updated }; }));
      if (!ready()) return Promise.resolve([]);
      var out = [], it = folder.values();
      function step() {
        return it.next().then(function (r) {
          if (r.done) return out;
          var h = r.value, app = h.kind === "file" ? appOf(h.name) : null;
          if (!app) return step();
          return h.getFile().then(function (file) {
            out.push({ id: h.name, app: app, name: baseName(h.name), updated: file.lastModified, size: file.size });
            return step();
          }, step);
        });
      }
      return step();
    },
    get: function (id) {
      if (!supported) return Promise.resolve(lsAll().filter(function (f) { return f.id === id; })[0] || null);
      if (!ready()) return Promise.resolve(null);
      return folder.getFileHandle(id).then(function (h) { return h.getFile(); }).then(function (file) {
        return file.text().then(function (text) {
          var o = parse(text, appOf(id));
          return { id: id, app: o.app, name: baseName(id), tpl: o.tpl, created: o.created, updated: file.lastModified, data: o.data };
        });
      }).catch(function () { return null; });
    },
    create: function (app, name, tpl, data) {
      var now = Date.now();
      var f = { app: app, name: name, tpl: tpl || "", created: now, updated: now, data: data };
      if (!supported) {
        f.id = newId();
        return Promise.resolve().then(function () { var list = lsAll(); list.push(f); lsSave(list); return f; });
      }
      return inOrder(function () {
        return uniqueName(name, app).then(function (fileName) {
          f.id = fileName; f.name = baseName(fileName);
          return write(fileName, payload(f)).then(function () { return f; });
        });
      });
    },
    save: function (f) {
      f.updated = Date.now();
      if (!supported) {
        return Promise.resolve().then(function () { var list = lsAll().filter(function (x) { return x.id !== f.id; }); list.push(f); lsSave(list); return f; });
      }
      var text = payload(f), id = f.id;
      return inOrder(function () { return write(id, text).then(function () { return f; }); });
    },
    // 名前を変える（フォルダでは、新しい名前で書いて、前のファイルを消す）。f.id が変わることがある
    rename: function (f, name) {
      if (!supported) { f.name = name; return Store.save(f); }
      return inOrder(function () {
        return uniqueName(name, f.app, f.id).then(function (fileName) {
          if (fileName === f.id) return f;
          var old = f.id;
          return write(fileName, payload(f)).then(function () { return folder.removeEntry(old); }).then(function () {
            f.id = fileName; f.name = baseName(fileName); f.updated = Date.now();
            return f;
          });
        });
      });
    },
    // フォルダからファイルを消す（ごみ箱には入らない）
    remove: function (id) {
      if (!supported) { lsSave(lsAll().filter(function (f) { return f.id !== id; })); return Promise.resolve(); }
      return inOrder(function () { return folder.removeEntry(id); });
    },
    copy: function (f, name) { return Store.create(f.app, name, f.tpl, JSON.parse(JSON.stringify(f.data))); },

    // 前にブラウザの中に保存したファイル（フォルダに保存できるブラウザで、フォルダを選んだあとに移す）
    legacy: function () { return supported ? lsAll() : []; },
    migrate: function () {
      var list = lsAll(), done = 0;
      return list.reduce(function (p, f) {
        return p.then(function () { return Store.create(f.app, f.name, f.tpl, f.data); }).then(function () { done++; });
      }, Promise.resolve()).then(function () {
        try { localStorage.removeItem(LS_KEY); } catch (e) { /* 消せなくても、移したものはフォルダにある */ }
        return done;
      });
    },

    // ほかの場所のファイル（.skdoc など）を読み込んで、保存先に入れる
    importFile: function (file) {
      var app = appOf(file.name);
      return file.text().then(function (text) {
        var o = parse(text, app);
        if (!EXT[o.app]) throw new Error("not-office");
        return Store.create(o.app, baseName(file.name), o.tpl, o.data);
      });
    },
    // ファイルとして書き出す（フォルダに保存できるブラウザは、保存する場所を選ぶ。それ以外はダウンロード）
    exportFile: function (f) {
      var text = payload(f), fileName = clean(f.name) + EXT[f.app];
      if (typeof window.showSaveFilePicker === "function") {
        var accept = {}; accept["application/json"] = [EXT[f.app]];
        return window.showSaveFilePicker({ suggestedName: fileName, types: [{ description: "SK Hub Office", accept: accept }] }).then(function (h) {
          return h.createWritable().then(function (w) { return w.write(text).then(function () { return w.close(); }); }).then(function () { return true; });
        }, function () { return false; });
      }
      var a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([text], { type: "application/json" }));
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      return Promise.resolve(true);
    }
  };
  window.OfficeStore = Store;
})();
