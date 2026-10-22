// SK Hub Office のファイルの保存（window.OfficeStore）
//   ・PC のフォルダに 1 ファイルずつ保存する（File System Access API。Chrome・Edge のパソコン）
//       最初に 1 回だけ、保存するフォルダを選んでもらう（showDirectoryPicker）。選んだフォルダは IndexedDB の sk_office に覚えておく
//       次に開いたときにブラウザが許可を聞き直すことがあるので、そのときは「許可する」を押してもらう（allow()）
//       ファイル名は「名前.skdoc」「名前.sksheet」「名前.skslide」。中身は JSON：
//         { format: "sk-hub-office", version: 2, app, tpl, created, data }（data の形は templates.js の説明のとおり。version 1 も読める）
//       ファイルの ID はフォルダの中のファイル名
//   ・フォルダに保存できないブラウザ（Safari・Firefox・スマホなど）は、ブラウザの中（localStorage の sk_office_files）に保存する
//       [{ id, app, name, tpl, created, updated, data }]。ファイルとして書き出す（exportFile）・読み込む（importFile）もできる
//   ・前にブラウザの中に保存したファイルは、フォルダを選んだあとに migrate() でフォルダに移せる
//   ・アプリとして入れたあと、エクスプローラー・Finder で .skdoc などをダブルクリックしたときは openHandle() で、そのファイルに直接保存する
//   ・版の履歴（versions・addVersion）は、このブラウザの IndexedDB に残す（ほかの PC には移らない）
//   ・Google ドライブ（drive.js の window.OfficeDrive）にも保存できる。ファイルの ID が「g:」で始まるものは Google ドライブのファイル。
//     新しく作るファイルをどこに入れるかは place()（"local": この PC のフォルダかブラウザの中・"drive": Google ドライブ）。setPlace() で変える
//     list() は両方のファイルを返す（どちらのファイルかは place）。get・save・rename・remove は ID で、どちらに行くかを決める
//   どれも Promise を返す（state()・folderName()・supported・place()・driveState() だけはすぐ返す）
(function () {
  "use strict";
  var Drive = window.OfficeDrive || null;
  var PLACE_KEY = "sk_office_place";
  var EXT = { doc: ".skdoc", sheet: ".sksheet", slide: ".skslide" };
  var FORMAT = "sk-hub-office";
  var FILE_VERSION = 2; // 1 は 0.1。2 は書式・複数シート・ノートなどを足した形（1 も読める）
  var LS_KEY = "sk_office_files";
  var supported = typeof window.showDirectoryPicker === "function" && window.isSecureContext;
  var folder = null, state = supported ? "none" : "unsupported";

  function appOf(fileName) {
    for (var a in EXT) if (fileName.toLowerCase().slice(-EXT[a].length) === EXT[a]) return a;
    return null;
  }
  function baseName(fileName) { var a = appOf(fileName); return a ? fileName.slice(0, -EXT[a].length) : fileName; }
  function payload(f) { return JSON.stringify({ format: FORMAT, version: FILE_VERSION, app: f.app, tpl: f.tpl || "", created: f.created || Date.now(), data: f.data }); }
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
  // フォルダの外のファイル（エクスプローラー・Finder でダブルクリックして開いたもの）。ID は「@番号/ファイル名」。開いているあいだだけ覚える
  var handles = {}, handleSeq = 0;
  function isHandleId(id) { return id.charAt(0) === "@"; }
  function writeHandle(h, text) {
    var ask = typeof h.requestPermission === "function" ? h.requestPermission({ mode: "readwrite" }) : Promise.resolve("granted");
    return ask.then(function (p) {
      if (p !== "granted") throw new Error("no-permission");
      return h.createWritable().then(function (w) { return w.write(text).then(function () { return w.close(); }); });
    });
  }

  // ---------- Google ドライブ ----------
  function isDrive(id) { return String(id).slice(0, 2) === "g:"; }
  function gid(id) { return String(id).slice(2); }
  function driveState() { return Drive ? Drive.state() : "unconfigured"; }
  function driveFile(m, extra) {
    var app = appOf(m.name) || m.app;
    return Object.assign({ id: "g:" + m.id, app: app, name: baseName(m.name), updated: m.updated, created: m.created, size: m.size, place: "drive", mine: m.mine, owner: m.owner, canEdit: m.canEdit }, extra || {});
  }
  // 同じファイルの保存が続いたら、いま送っている分が終わってから、いちばん新しい中身だけを送る（自動保存で何度も送らない）
  var driveQ = {};
  function driveSave(id, text) {
    var x = driveQ[id] || (driveQ[id] = { running: null, pending: null, next: null });
    x.pending = text;
    if (!x.running) return start();
    return x.next || (x.next = x.running.catch(function () {}).then(start));
    function start() {
      var body = x.pending;
      x.pending = null; x.next = null;
      x.running = Drive.save(id, body).then(function (m) { x.running = null; return m; }, function (e) { x.running = null; throw e; });
      return x.running;
    }
  }

  var Store = {
    EXT: EXT,
    supported: supported,
    appOf: appOf,
    state: function () { return state; },
    folderName: function () { return folder ? folder.name : ""; },
    isDrive: isDrive,
    driveState: driveState,
    drive: Drive,
    // 新しいファイルを入れる場所（"local"・"drive"）。Google ドライブを使えないときは "local"
    place: function () { var p = ""; try { p = localStorage.getItem(PLACE_KEY) || ""; } catch (e) { /* なし */ } return p === "drive" && driveState() !== "unconfigured" && driveState() !== "off" ? "drive" : "local"; },
    setPlace: function (p) { try { localStorage.setItem(PLACE_KEY, p === "drive" ? "drive" : "local"); } catch (e) { /* このページだけ */ } },

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

    // ファイルの一覧（中身は読まない）: [{ id, app, name, updated, size, place }]。Google ドライブを読めなかったときは driveError に理由を入れる
    list: function () {
      Store.driveError = null;
      var drive = driveState() === "ready"
        ? Drive.list().then(function (all) { return all.map(function (m) { return driveFile(m); }).filter(function (f) { return EXT[f.app]; }); }, function (e) { Store.driveError = e; return []; })
        : Promise.resolve([]);
      // この PC のフォルダを読めなくても、Google ドライブのファイルは出す（両方だめなときだけ失敗にする）
      var local = Store.listLocal().catch(function (e) { if (driveState() !== "ready") throw e; Store.localError = e; return []; });
      return Promise.all([local, drive]).then(function (r) { return r[0].concat(r[1]); });
    },
    listLocal: function () {
      if (!supported) return Promise.resolve(lsAll().map(function (f) { return { id: f.id, app: f.app, name: f.name, updated: f.updated, place: "local" }; }));
      if (!ready()) return Promise.resolve([]);
      var out = [], it = folder.values();
      function step() {
        return it.next().then(function (r) {
          if (r.done) return out;
          var h = r.value, app = h.kind === "file" ? appOf(h.name) : null;
          if (!app) return step();
          return h.getFile().then(function (file) {
            out.push({ id: h.name, app: app, name: baseName(h.name), updated: file.lastModified, size: file.size, place: "local" });
            return step();
          }, step);
        });
      }
      return step();
    },
    // ファイルを読む。見つからない・読めないときは null（Google ドライブにつなぎなおしが要るときは、Error("auth") で失敗する）
    get: function (id) {
      if (isDrive(id)) {
        return Drive.get(gid(id)).then(function (m) {
          var o = parse(m.text, appOf(m.name) || m.app);
          return driveFile(m, { app: o.app, tpl: o.tpl, created: o.created, data: o.data });
        }, function (e) { if (e && e.message === "auth") throw e; return null; });
      }
      if (isHandleId(id)) {
        var hd = handles[id];
        if (!hd) return Promise.resolve(null);
        return hd.getFile().then(function (file) {
          return file.text().then(function (text) {
            var o = parse(text, appOf(hd.name));
            return { id: id, app: o.app, name: baseName(hd.name), tpl: o.tpl, created: o.created, updated: file.lastModified, data: o.data, outside: true };
          });
        }).catch(function () { return null; });
      }
      if (!supported) return Promise.resolve(lsAll().filter(function (f) { return f.id === id; })[0] || null);
      if (!ready()) return Promise.resolve(null);
      return folder.getFileHandle(id).then(function (h) { return h.getFile(); }).then(function (file) {
        return file.text().then(function (text) {
          var o = parse(text, appOf(id));
          return { id: id, app: o.app, name: baseName(id), tpl: o.tpl, created: o.created, updated: file.lastModified, data: o.data };
        });
      }).catch(function () { return null; });
    },
    // where: "local"・"drive"（書かないと place()）
    create: function (app, name, tpl, data, where) {
      var now = Date.now();
      var f = { app: app, name: name, tpl: tpl || "", created: now, updated: now, data: data };
      if ((where || Store.place()) === "drive") {
        return Drive.create(clean(name) + EXT[app], app, payload(f)).then(function (m) { return driveFile(m, { tpl: f.tpl, created: now, data: data }); });
      }
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
      if (isDrive(f.id)) return driveSave(gid(f.id), payload(f)).then(function () { return f; });
      if (isHandleId(f.id)) {
        var hd = handles[f.id], body = payload(f);
        if (!hd) return Promise.reject(new Error("closed"));
        return inOrder(function () { return writeHandle(hd, body).then(function () { return f; }); });
      }
      if (!supported) {
        return Promise.resolve().then(function () { var list = lsAll().filter(function (x) { return x.id !== f.id; }); list.push(f); lsSave(list); return f; });
      }
      var text = payload(f), id = f.id;
      return inOrder(function () { return write(id, text).then(function () { return f; }); });
    },
    // 名前を変える（フォルダでは、新しい名前で書いて、前のファイルを消す）。f.id が変わることがある
    rename: function (f, name) {
      if (isDrive(f.id)) return Drive.rename(gid(f.id), clean(name) + EXT[f.app]).then(function (m) { f.name = baseName(m.name); f.updated = m.updated; return f; });
      if (isHandleId(f.id)) return Promise.reject(new Error("outside"));
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
    // フォルダからファイルを消す（ごみ箱には入らない）。Google ドライブのファイルは、ドライブのごみ箱に入れる（30 日は戻せる）
    remove: function (id) {
      if (isDrive(id)) return Drive.trash(gid(id));
      if (!supported) { lsSave(lsAll().filter(function (f) { return f.id !== id; })); return Promise.resolve(); }
      return inOrder(function () { return folder.removeEntry(id); });
    },
    // コピーは、もとのファイルと同じ場所に作る（共有されたファイルのコピーは、自分のドライブに作る）
    copy: function (f, name) { return Store.create(f.app, name, f.tpl, JSON.parse(JSON.stringify(f.data)), isDrive(f.id) ? "drive" : "local"); },

    // 前にブラウザの中に保存したファイル（フォルダに保存できるブラウザで、フォルダを選んだあとに移す）
    legacy: function () { return supported ? lsAll() : []; },
    migrate: function () {
      var list = lsAll(), done = 0;
      return list.reduce(function (p, f) {
        return p.then(function () { return Store.create(f.app, f.name, f.tpl, f.data, "local"); }).then(function () { done++; });
      }, Promise.resolve()).then(function () {
        try { localStorage.removeItem(LS_KEY); } catch (e) { /* 消せなくても、移したものはフォルダにある */ }
        return done;
      });
    },

    // エクスプローラー・Finder でダブルクリックしたファイル（PWA の file_handlers。launchQueue で来る）を開けるようにして、ID を返す
    //   選んだフォルダの中のファイルなら、ふつうの ID（ファイル名）。外のファイルなら「@番号/ファイル名」で、そのファイルに直接保存する
    openHandle: function (h) {
      var inFolder = ready() && typeof folder.resolve === "function"
        ? folder.resolve(h).then(function (path) { return path && path.length === 1 ? path[0] : null; }, function () { return null; })
        : Promise.resolve(null);
      return inFolder.then(function (name) {
        if (name) return name;
        var id = "@" + (++handleSeq) + "/" + h.name;
        handles[id] = h;
        return id;
      });
    },
    isOutside: isHandleId,

    // 版の履歴（前の版を、このブラウザの IndexedDB の kv に「v:ファイルの ID」で 20 個まで残す。ファイルそのものには入れない）
    //   [{ ts, size, data（JSON の文字列） }]。新しいものが先
    versions: function (id) {
      if (!id || isHandleId(id)) return Promise.resolve([]);
      return idbGet("v:" + id).then(function (a) { return Array.isArray(a) ? a : []; }, function () { return []; });
    },
    addVersion: function (id, data) {
      if (!id || isHandleId(id)) return Promise.resolve(false);
      var text = JSON.stringify(data);
      return Store.versions(id).then(function (list) {
        if (list.length && list[0].data === text) return false;
        list.unshift({ ts: Date.now(), size: text.length, data: text });
        // 大きなファイル（画像の多いもの）は、残す数を減らす
        var total = 0;
        list = list.filter(function (v, i) { total += v.size; return i < 20 && (i < 3 || total < 30000000); });
        return idbSet("v:" + id, list).then(function () { return true; });
      }).catch(function () { return false; });
    },
    moveVersions: function (from, to) {
      if (!from || !to || from === to) return Promise.resolve();
      return Store.versions(from).then(function (list) {
        if (!list.length) return;
        return idbSet("v:" + to, list).then(function () { return idb("readwrite", function (s) { return s.delete("v:" + from); }); });
      }).catch(function () {});
    },
    dropVersions: function (id) {
      return idb("readwrite", function (s) { return s.delete("v:" + id); }).catch(function () {});
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
