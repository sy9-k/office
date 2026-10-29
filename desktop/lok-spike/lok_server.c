// LOKit のローカルサーバー（WebView の試作用）
//   使い方: lok_server.exe <文書> <LibreOffice の program フォルダ> <プロファイルのフォルダ> <web フォルダ> <ポート>
//   127.0.0.1 だけで待ち受ける。LibreOffice の呼び出しは、すべてこのスレッドで行う（LOKit はスレッドを気にするため）
//   エンドポイント:
//     GET  /               web/index.html を返す（そのほかのファイルは web フォルダから返す）
//     GET  /size           文書の大きさ（twips）を JSON で返す
//     GET  /tile?y=&w=&h=  縦の位置 y（twips）から、幅 w・高さ h（px、96dpi）を描いた BMP を返す
//     POST /key?type=&cp=  キー入力（type: 0 = 押す、1 = 離す。cp は 16 進のコードポイント）
//     POST /ime?type=&text=  変換中（0）・確定（2）の文字列（text は 16 進のコードポイントを「,」区切り）
#define LOK_USE_UNSTABLE_API
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <stdbool.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "LibreOfficeKit.h"

#pragma comment(lib, "ws2_32.lib")

typedef LibreOfficeKit *(*HookFn)(const char *install_path, const char *user_profile_url);

static LibreOfficeKit *g_kit;
static LibreOfficeKitDocument *g_doc;
static char g_web_dir[MAX_PATH * 2];
static char g_doc_abs[MAX_PATH * 2];   // 開いた文書のパス（保存先は、ここに「.saved.docx」を付けたもの）
static char g_saved_abs[MAX_PATH * 2]; // 直近の保存先（書き出したあと、/saved.docx で返す）
static int g_saved_ok;

// LOKit のコールバックの種類（LibreOfficeKitEnums.h の値。C から読むので、ここに写す）
#define LOK_CB_INVALIDATE_TILES 0          // 描画が変わった
#define LOK_CB_INVALIDATE_VISIBLE_CURSOR 1 // カーソルの位置（payload: "x, y, w, h"・twips）
#define LOK_CB_TEXT_SELECTION 2            // 選択の範囲（payload: "x, y, w, h; x, y, w, h; …"・twips。選択がないと空）

// カーソルの位置（twips）と、描画が変わった回数。コールバックで更新する
static int g_cursor_valid;
static long g_cursor_x, g_cursor_y, g_cursor_w, g_cursor_h;
static unsigned long g_dirty;
static char g_selection[4096]; // 選択の範囲（そのまま JSON に入れる。数字・「,」「;」「空白」だけ）

// LOKit の処理ループ（命令の実行や、画面の更新を回す）。専用のスレッドで回す。
// 回しておかないと、postUnoCommand などの命令が実行されない
static int lok_poll(void *data, int timeoutUs) {
    (void)data;
    int ms = timeoutUs / 1000;
    Sleep(ms < 1 ? 1 : (ms > 5 ? 5 : ms));
    return 0;
}

static void lok_wake(void *data) {
    (void)data;
}

static DWORD WINAPI lok_loop_thread(LPVOID arg) {
    (void)arg;
    g_kit->pClass->runLoop(g_kit, lok_poll, lok_wake, NULL);
    return 0;
}

static void on_lok_callback(int type, const char *payload, void *data) {
    (void)data;
    if (type == LOK_CB_INVALIDATE_TILES) {
        g_dirty++;
    } else if (type == LOK_CB_INVALIDATE_VISIBLE_CURSOR && payload) {
        if (sscanf(payload, "%ld, %ld, %ld, %ld", &g_cursor_x, &g_cursor_y, &g_cursor_w, &g_cursor_h) == 4) {
            g_cursor_valid = 1;
        }
    } else if (type == LOK_CB_TEXT_SELECTION) {
        snprintf(g_selection, sizeof g_selection, "%s", payload ? payload : "");
    }
}

// フォルダのパスを file:/// の URL にする（空白は %20）
static void to_file_url(const char *path, char *out, size_t n) {
    char tmp[MAX_PATH * 4];
    size_t j = 0;
    for (size_t i = 0; path[i] && j + 4 < sizeof tmp; i++) {
        char c = path[i] == '\\' ? '/' : path[i];
        if (c == ' ') { memcpy(tmp + j, "%20", 3); j += 3; }
        else tmp[j++] = c;
    }
    tmp[j] = 0;
    snprintf(out, n, "file:///%s", tmp);
}

// 16 進のコードポイント列（「,」区切り）を UTF-8 にする
static void hex_to_utf8(const char *hex, char *out, size_t n) {
    size_t j = 0;
    char tmp[1024];
    snprintf(tmp, sizeof tmp, "%s", hex);
    char *save = NULL;
    for (char *tok = strtok_s(tmp, ",", &save); tok; tok = strtok_s(NULL, ",", &save)) {
        unsigned int cp = (unsigned int)strtoul(tok, NULL, 16);
        if (j + 4 >= n) break;
        if (cp < 0x80) out[j++] = (char)cp;
        else if (cp < 0x800) { out[j++] = (char)(0xC0 | (cp >> 6)); out[j++] = (char)(0x80 | (cp & 0x3F)); }
        else { out[j++] = (char)(0xE0 | (cp >> 12)); out[j++] = (char)(0x80 | ((cp >> 6) & 0x3F)); out[j++] = (char)(0x80 | (cp & 0x3F)); }
    }
    out[j] = 0;
}

// 「?a=1&b=2」から値を取り出す（値は URL エンコードしていない前提）
static bool get_param(const char *query, const char *name, char *out, size_t n) {
    size_t nl = strlen(name);
    const char *p = query;
    while (p && *p) {
        if (strncmp(p, name, nl) == 0 && p[nl] == '=') {
            const char *v = p + nl + 1;
            size_t k = 0;
            while (v[k] && v[k] != '&' && k + 1 < n) { out[k] = v[k]; k++; }
            out[k] = 0;
            return true;
        }
        p = strchr(p, '&');
        if (p) p++;
    }
    out[0] = 0;
    return false;
}

static void send_all(SOCKET s, const void *data, size_t len) {
    const char *p = data;
    while (len > 0) {
        int n = send(s, p, (int)(len > 1 << 20 ? 1 << 20 : len), 0);
        if (n <= 0) return;
        p += n;
        len -= (size_t)n;
    }
}

static void send_response(SOCKET s, const char *status, const char *type, const void *body, size_t len) {
    char head[512];
    int hn = snprintf(head, sizeof head,
        "HTTP/1.1 %s\r\nContent-Type: %s\r\nContent-Length: %zu\r\nConnection: close\r\nCache-Control: no-store\r\nAccess-Control-Allow-Origin: *\r\n\r\n",
        status, type, len);
    send_all(s, head, (size_t)hn);
    if (len) send_all(s, body, len);
}

static void send_text(SOCKET s, const char *status, const char *text) {
    send_response(s, status, "text/plain; charset=utf-8", text, strlen(text));
}

// BMP（32 ビット、上から下）を作る。呼び出し側が free する
static unsigned char *make_bmp(const unsigned char *px, int w, int h, size_t *out_len) {
    size_t img = (size_t)w * (size_t)h * 4;
    size_t total = 54 + img;
    unsigned char *b = malloc(total);
    if (!b) return NULL;
    memset(b, 0, 54);
    b[0] = 'B'; b[1] = 'M';
    *(unsigned int *)(b + 2) = (unsigned int)total;
    *(unsigned int *)(b + 10) = 54;
    *(unsigned int *)(b + 14) = 40;
    *(int *)(b + 18) = w;
    *(int *)(b + 22) = -h;
    *(unsigned short *)(b + 26) = 1;
    *(unsigned short *)(b + 28) = 32;
    *(unsigned int *)(b + 34) = (unsigned int)img;
    memcpy(b + 54, px, img);
    *out_len = total;
    return b;
}

static void handle_request(SOCKET s, const char *method, const char *target) {
    char path[512] = { 0 };
    const char *q = strchr(target, '?');
    size_t plen = q ? (size_t)(q - target) : strlen(target);
    if (plen >= sizeof path) plen = sizeof path - 1;
    memcpy(path, target, plen);
    path[plen] = 0;
    const char *query = q ? q + 1 : "";
    char p1[256], p2[256], p3[256];

    if (strcmp(path, "/size") == 0) {
        long w = 0, h = 0;
        g_doc->pClass->getDocumentSize(g_doc, &w, &h);
        char json[128];
        snprintf(json, sizeof json, "{\"w\":%ld,\"h\":%ld}", w, h);
        send_response(s, "200 OK", "application/json", json, strlen(json));
        return;
    }
    if (strcmp(path, "/tile") == 0) {
        get_param(query, "y", p1, sizeof p1);
        get_param(query, "w", p2, sizeof p2);
        get_param(query, "h", p3, sizeof p3);
        int y = atoi(p1), w = atoi(p2), h = atoi(p3);
        if (w <= 0 || h <= 0 || w > 4000 || h > 4000) { send_text(s, "400 Bad Request", "bad size"); return; }
        unsigned char *buf = malloc((size_t)w * h * 4);
        if (!buf) { send_text(s, "500 Internal Server Error", "no memory"); return; }
        // 1 px = 15 twips（96dpi）
        g_doc->pClass->paintTile(g_doc, buf, w, h, 0, y, w * 15, h * 15);
        size_t len = 0;
        unsigned char *bmp = make_bmp(buf, w, h, &len);
        free(buf);
        if (!bmp) { send_text(s, "500 Internal Server Error", "no memory"); return; }
        send_response(s, "200 OK", "image/bmp", bmp, len);
        free(bmp);
        return;
    }
    if (strcmp(path, "/cursor") == 0) {
        char json[4608];
        snprintf(json, sizeof json,
            "{\"valid\":%d,\"x\":%ld,\"y\":%ld,\"w\":%ld,\"h\":%ld,\"dirty\":%lu,\"sel\":\"%s\"}",
            g_cursor_valid, g_cursor_x, g_cursor_y, g_cursor_w, g_cursor_h, g_dirty, g_selection);
        send_response(s, "200 OK", "application/json", json, strlen(json));
        return;
    }
    if (strcmp(path, "/mouse") == 0 && strcmp(method, "POST") == 0) {
        // type: 0 = 押す、1 = 離す、2 = 動かす。x・y は twips（文書の位置）
        get_param(query, "type", p1, sizeof p1);
        get_param(query, "x", p2, sizeof p2);
        get_param(query, "y", p3, sizeof p3);
        g_doc->pClass->postMouseEvent(g_doc, atoi(p1), atoi(p2), atoi(p3), 1, 1, 0);
        send_text(s, "200 OK", "ok");
        return;
    }
    if (strcmp(path, "/save") == 0 && strcmp(method, "POST") == 0) {
        // 今の内容を docx に書き出す（LibreOffice の書き出しを使う）。保存先は文書の隣の「.saved.docx」
        char url[MAX_PATH * 4];
        snprintf(g_saved_abs, sizeof g_saved_abs, "%s.saved.docx", g_doc_abs);
        to_file_url(g_saved_abs, url, sizeof url);
        // 形式は NULL にして、保存先の拡張子（.docx）から LibreOffice に決めさせる
        int ok = g_doc->pClass->saveAs(g_doc, url, NULL, NULL);
        g_saved_ok = ok ? 1 : 0;
        send_text(s, ok ? "200 OK" : "500 Internal Server Error", ok ? "ok" : "save failed");
        return;
    }
    if (strcmp(path, "/saved.docx") == 0) {
        FILE *f = g_saved_ok ? fopen(g_saved_abs, "rb") : NULL;
        if (!f) { send_text(s, "404 Not Found", "not saved yet"); return; }
        fseek(f, 0, SEEK_END);
        long n = ftell(f);
        fseek(f, 0, SEEK_SET);
        unsigned char *body = malloc((size_t)n);
        if (body && fread(body, 1, (size_t)n, f) == (size_t)n) {
            send_response(s, "200 OK", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", body, (size_t)n);
        } else {
            send_text(s, "500 Internal Server Error", "read failed");
        }
        free(body);
        fclose(f);
        return;
    }
    if (strcmp(path, "/key") == 0 && strcmp(method, "POST") == 0) {
        // cp: 文字のコードポイント（16 進）。kc: 特殊キーの番号（css::awt::Key。Enter は 1280 など。0 なら使わない）
        get_param(query, "type", p1, sizeof p1);
        get_param(query, "cp", p2, sizeof p2);
        get_param(query, "kc", p3, sizeof p3);
        int type = atoi(p1);
        int cp = (int)strtoul(p2, NULL, 16);
        int kc = atoi(p3);
        g_doc->pClass->postKeyEvent(g_doc, type, cp, kc);
        send_text(s, "200 OK", "ok");
        return;
    }
    if (strcmp(path, "/ime") == 0 && strcmp(method, "POST") == 0) {
        get_param(query, "type", p1, sizeof p1);
        static char text_hex[1024];
        get_param(query, "text", text_hex, sizeof text_hex);
        char utf8[1024];
        hex_to_utf8(text_hex, utf8, sizeof utf8);
        g_doc->pClass->postWindowExtTextInputEvent(g_doc, 0, atoi(p1), utf8);
        send_text(s, "200 OK", "ok");
        return;
    }
    if (strcmp(path, "/") == 0) {
        strcpy_s(path, sizeof path, "/index.html");
    }
    if (strstr(path, "..") == NULL && strcmp(method, "GET") == 0) {
        char file[MAX_PATH * 3];
        snprintf(file, sizeof file, "%s%s", g_web_dir, path);
        FILE *f = fopen(file, "rb");
        if (f) {
            fseek(f, 0, SEEK_END);
            long n = ftell(f);
            fseek(f, 0, SEEK_SET);
            unsigned char *body = malloc((size_t)n);
            if (body && fread(body, 1, (size_t)n, f) == (size_t)n) {
                const char *type = strstr(path, ".js") ? "text/javascript; charset=utf-8" :
                                   strstr(path, ".css") ? "text/css; charset=utf-8" :
                                   "text/html; charset=utf-8";
                send_response(s, "200 OK", type, body, (size_t)n);
            } else {
                send_text(s, "500 Internal Server Error", "read failed");
            }
            free(body);
            fclose(f);
            return;
        }
    }
    send_text(s, "404 Not Found", "not found");
}

static void serve_client(SOCKET c) {
    static char req[16384];
    int got = 0;
    while (got < (int)sizeof req - 1) {
        int n = recv(c, req + got, (int)sizeof req - 1 - got, 0);
        if (n <= 0) break;
        got += n;
        req[got] = 0;
        if (strstr(req, "\r\n\r\n")) break;
    }
    if (got <= 0) return;
    char method[16] = { 0 }, target[1024] = { 0 };
    if (sscanf(req, "%15s %1023s", method, target) != 2) return;
    handle_request(c, method, target);
}

int main(int argc, char **argv) {
    if (argc < 6) {
        fprintf(stderr, "usage: lok_server <file> <program_dir> <profile_dir> <web_dir> <port>\n");
        return 2;
    }
    const char *doc_path = argv[1], *prog = argv[2], *profile = argv[3];
    snprintf(g_web_dir, sizeof g_web_dir, "%s", argv[4]);
    int port = atoi(argv[5]);
    setvbuf(stdout, NULL, _IONBF, 0);

    char ini_path[MAX_PATH * 2], ini_url[MAX_PATH * 4], profile_url[MAX_PATH * 4], doc_abs[MAX_PATH * 2], doc_url[MAX_PATH * 4];
    snprintf(ini_path, sizeof ini_path, "%s\\fundamental.ini", prog);
    to_file_url(ini_path, ini_url, sizeof ini_url);
    SetEnvironmentVariableA("URE_BOOTSTRAP", ini_url);
    SetEnvironmentVariableA("UNO_PATH", prog);
    char path_env[32768];
    DWORD got_path = GetEnvironmentVariableA("PATH", path_env, sizeof path_env);
    char new_path[32768 + MAX_PATH];
    snprintf(new_path, sizeof new_path, "%s;%s", prog, got_path ? path_env : "");
    SetEnvironmentVariableA("PATH", new_path);
    SetDllDirectoryA(prog);

    char dll[MAX_PATH * 2];
    snprintf(dll, sizeof dll, "%s\\mergedlo.dll", prog);
    HMODULE mod = LoadLibraryA(dll);
    if (!mod) { fprintf(stderr, "LoadLibrary failed: %lu\n", GetLastError()); return 1; }
    HookFn hook = (HookFn)GetProcAddress(mod, "libreofficekit_hook_2");
    if (!hook) { fprintf(stderr, "hook not found\n"); return 1; }

    to_file_url(profile, profile_url, sizeof profile_url);
    GetFullPathNameA(doc_path, sizeof doc_abs, doc_abs, NULL);
    snprintf(g_doc_abs, sizeof g_doc_abs, "%s", doc_abs);
    to_file_url(doc_abs, doc_url, sizeof doc_url);

    // 前の起動で残ったロックファイル（.~lock.名前#）があると、LibreOffice はその文書を読み取り専用で開く。
    // このファイルは、このサーバーが作った一時ファイルなので、起動の前に消す
    {
        char dir[MAX_PATH * 2], base[MAX_PATH], lock[MAX_PATH * 2 + 64];
        snprintf(dir, sizeof dir, "%s", doc_abs);
        char *slash = strrchr(dir, '\\');
        if (slash) {
            snprintf(base, sizeof base, "%s", slash + 1);
            *slash = 0;
            snprintf(lock, sizeof lock, "%s\\.~lock.%s#", dir, base);
            DeleteFileA(lock);
        }
    }
    g_kit = hook(prog, profile_url);
    if (!g_kit) { fprintf(stderr, "LOKit init failed\n"); return 1; }
    g_doc = g_kit->pClass->documentLoad(g_kit, doc_url);
    if (!g_doc) {
        char *e = g_kit->pClass->getError(g_kit);
        fprintf(stderr, "load failed: %s\n", e ? e : "(no message)");
        return 1;
    }
    g_doc->pClass->initializeForRendering(g_doc, NULL);
    g_doc->pClass->registerCallback(g_doc, on_lok_callback, NULL);
    CreateThread(NULL, 0, lok_loop_thread, NULL, 0, NULL);
    printf("loaded %s\n", doc_abs);

    WSADATA wsa;
    WSAStartup(MAKEWORD(2, 2), &wsa);
    SOCKET ls = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    struct sockaddr_in addr = { 0 };
    addr.sin_family = AF_INET;
    addr.sin_port = htons((unsigned short)port);
    addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK); // 127.0.0.1 だけ
    int one = 1;
    setsockopt(ls, SOL_SOCKET, SO_REUSEADDR, (const char *)&one, sizeof one);
    if (bind(ls, (struct sockaddr *)&addr, sizeof addr) != 0 || listen(ls, 8) != 0) {
        fprintf(stderr, "listen failed on port %d\n", port);
        return 1;
    }
    printf("listening on http://127.0.0.1:%d/\n", port);

    for (;;) {
        SOCKET c = accept(ls, NULL, NULL);
        if (c == INVALID_SOCKET) continue;
        serve_client(c);
        closesocket(c);
    }
}
