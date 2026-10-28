// LOKit の 1 ページ描画の試作（Doc の案 A の確認用）
//   使い方: lok_spike.exe <文書> <出力.bmp> <LibreOffice の program フォルダ> <プロファイルのフォルダ>
//   mergedlo.dll を直接読み込み、libreofficekit_hook_2 で LOKit を起動する。文書を開いて、最初のページを 96dpi で 1 枚描いて BMP に書く
//   画素は 32 ビット（BGRA と想定）。色が入れ替わって見えたら、ここの並びを確かめる
// 描画（paintTile など）は LOK の「不安定」区分にあるので、このマクロで有効にする
#define LOK_USE_UNSTABLE_API
#include <windows.h>
#include <stdbool.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "LibreOfficeKit.h"

typedef LibreOfficeKit *(*HookFn)(const char *install_path, const char *user_profile_url);

static double now_ms(void) {
    LARGE_INTEGER f, c;
    QueryPerformanceFrequency(&f);
    QueryPerformanceCounter(&c);
    return c.QuadPart * 1000.0 / f.QuadPart;
}

// フォルダのパスを file:/// の URL にする（空白は %20）
static void to_file_url(const char *path, char *out, size_t n) {
    char tmp[MAX_PATH * 2];
    size_t j = 0;
    for (size_t i = 0; path[i] && j + 4 < sizeof tmp; i++) {
        char c = path[i] == '\\' ? '/' : path[i];
        if (c == ' ') { memcpy(tmp + j, "%20", 3); j += 3; }
        else tmp[j++] = c;
    }
    tmp[j] = 0;
    snprintf(out, n, "file:///%s", tmp);
}

static int write_bmp(const char *path, const unsigned char *px, int w, int h) {
    BITMAPFILEHEADER fh = { 0 };
    BITMAPINFOHEADER ih = { 0 };
    DWORD img = (DWORD)w * (DWORD)h * 4;
    fh.bfType = 0x4D42;
    fh.bfOffBits = sizeof fh + sizeof ih;
    fh.bfSize = fh.bfOffBits + img;
    ih.biSize = sizeof ih;
    ih.biWidth = w;
    ih.biHeight = -h; // 負の値は上から下へ並べる
    ih.biPlanes = 1;
    ih.biBitCount = 32;
    ih.biCompression = BI_RGB;
    ih.biSizeImage = img;
    FILE *f = fopen(path, "wb");
    if (!f) return 0;
    fwrite(&fh, sizeof fh, 1, f);
    fwrite(&ih, sizeof ih, 1, f);
    fwrite(px, 1, img, f);
    fclose(f);
    return 1;
}

int main(int argc, char **argv) {
    if (argc < 5) {
        fprintf(stderr, "usage: lok_spike <file> <out.bmp> <program_dir> <profile_dir>\n");
        return 2;
    }
    const char *doc_path = argv[1], *out_path = argv[2], *prog = argv[3], *profile = argv[4];
    setvbuf(stdout, NULL, _IONBF, 0); // 途中経過が止まっても見えるように、出力はバッファしない
    printf("start\n");

    // DLL の依存を見つけられるように、program フォルダを先頭に入れる
    char path_env[32768];
    DWORD got = GetEnvironmentVariableA("PATH", path_env, sizeof path_env);
    char new_path[32768 + MAX_PATH];
    snprintf(new_path, sizeof new_path, "%s;%s", prog, got ? path_env : "");
    SetEnvironmentVariableA("PATH", new_path);
    // LibreOffice の起動に必要な設定（soffice と同じく、program の fundamental.ini を指す）
    char ini_url[MAX_PATH * 2];
    char ini_path[MAX_PATH * 2];
    snprintf(ini_path, sizeof ini_path, "%s\\fundamental.ini", prog);
    to_file_url(ini_path, ini_url, sizeof ini_url);
    SetEnvironmentVariableA("URE_BOOTSTRAP", ini_url);
    SetEnvironmentVariableA("UNO_PATH", prog);

    SetDllDirectoryA(prog);
    char dll[MAX_PATH * 2];
    snprintf(dll, sizeof dll, "%s\\mergedlo.dll", prog);
    HMODULE mod = LoadLibraryA(dll);
    if (!mod) { fprintf(stderr, "LoadLibrary failed: %lu\n", GetLastError()); return 1; }

    HookFn hook = (HookFn)GetProcAddress(mod, "libreofficekit_hook_2");
    if (!hook) { fprintf(stderr, "libreofficekit_hook_2 not found\n"); return 1; }

    char profile_url[MAX_PATH * 2], doc_url[MAX_PATH * 2];
    // 相対パスは URL にできないので、絶対パスにしてから変換する
    char doc_abs[MAX_PATH * 2];
    GetFullPathNameA(doc_path, sizeof doc_abs, doc_abs, NULL);
    to_file_url(profile, profile_url, sizeof profile_url);
    to_file_url(doc_abs, doc_url, sizeof doc_url);

    double t0 = now_ms();
    LibreOfficeKit *kit = hook(prog, profile_url);
    if (!kit) { fprintf(stderr, "LOKit init failed\n"); return 1; }
    printf("init: %.0f ms\n", now_ms() - t0);

    t0 = now_ms();
    LibreOfficeKitDocument *doc = kit->pClass->documentLoad(kit, doc_url);
    if (!doc) {
        char *e = kit->pClass->getError(kit);
        fprintf(stderr, "load failed: %s\n", e ? e : "(no message)");
        return 1;
    }
    printf("load: %.0f ms\n", now_ms() - t0);

    doc->pClass->initializeForRendering(doc, NULL);
    long dw = 0, dh = 0;
    doc->pClass->getDocumentSize(doc, &dw, &dh);
    printf("document size (twips): %ld x %ld\n", dw, dh);

    // A4 を 96dpi で描く。1 px は 15 twips（1440 twips = 1 inch = 96 px）
    const int cw = 794, ch = 1123;
    unsigned char *buf = malloc((size_t)cw * ch * 4);
    memset(buf, 0xFF, (size_t)cw * ch * 4);

    t0 = now_ms();
    doc->pClass->paintTile(doc, buf, cw, ch, 0, 0, cw * 15, ch * 15);
    printf("paint 1 page: %.1f ms\n", now_ms() - t0);

    if (!write_bmp(out_path, buf, cw, ch)) fprintf(stderr, "cannot write %s\n", out_path);
    else printf("wrote %s\n", out_path);

    // 任意: 5 番目の引数に 16 進のコードポイントを「,」区切りで渡すと、文字を入れて再描画し、1 文字ずつの時間を測る
    //   例: 41,3042  → 「A」と「あ」
    if (argc >= 6 && strcmp(argv[5], "-") != 0) {
        char list[256];
        snprintf(list, sizeof list, "%s", argv[5]);
        char *tok = strtok(list, ",");
        while (tok) {
            unsigned int cp = (unsigned int)strtoul(tok, NULL, 16);
            double tk = now_ms();
            // LOK のキー種別: 0 = 押す（KEYINPUT）、1 = 離す（KEYUP）
            doc->pClass->postKeyEvent(doc, 0, (int)cp, 0);
            doc->pClass->postKeyEvent(doc, 1, (int)cp, 0);
            double tp = now_ms();
            doc->pClass->paintTile(doc, buf, cw, ch, 0, 0, cw * 15, ch * 15);
            printf("type U+%04X: key %.1f ms, repaint %.1f ms\n", cp, tp - tk, now_ms() - tp);
            tok = strtok(NULL, ",");
        }
        // 入力の反映が非同期なら、少し待ってから描き直すと最後の文字が出る
        Sleep(300);
        double tw = now_ms();
        doc->pClass->paintTile(doc, buf, cw, ch, 0, 0, cw * 15, ch * 15);
        printf("final repaint after 300 ms: %.1f ms\n", now_ms() - tw);
        char typed[MAX_PATH * 2];
        snprintf(typed, sizeof typed, "%s.typed.bmp", out_path);
        if (write_bmp(typed, buf, cw, ch)) printf("wrote %s\n", typed);
    }

    // IME の確認: argv[6] が "ime" のとき、変換中の文字を 3 段階で送り、最後に確定する
    //   LOK_EXT_TEXTINPUT（変換中）= 0、LOK_EXT_TEXTINPUT_END（確定）= 2（LibreOfficeKitEnums.h）
    //   送る文字列は UTF-8。コードポイントの並びは hex で書く
    if (argc >= 7 && strcmp(argv[6], "ime") == 0) {
        static const char *stages[][2] = {
            { "0", "304B" },              // か
            { "0", "304B,3093" },         // かん
            { "0", "304B,3093,3058" },    // かんじ
            { "0", "6F22,5B57" },         // 漢字（変換の候補を選んだ状態）
            { "2", "6F22,5B57" },         // 確定
        };
        const int nstage = (int)(sizeof stages / sizeof stages[0]);
        for (int i = 0; i < nstage; i++) {
            char utf8[64] = { 0 };
            size_t j = 0;
            char tmp[128];
            snprintf(tmp, sizeof tmp, "%s", stages[i][1]);
            char *tok = strtok(tmp, ",");
            while (tok) {
                unsigned int cp = (unsigned int)strtoul(tok, NULL, 16);
                if (cp < 0x80) { utf8[j++] = (char)cp; }
                else if (cp < 0x800) { utf8[j++] = (char)(0xC0 | (cp >> 6)); utf8[j++] = (char)(0x80 | (cp & 0x3F)); }
                else { utf8[j++] = (char)(0xE0 | (cp >> 12)); utf8[j++] = (char)(0x80 | ((cp >> 6) & 0x3F)); utf8[j++] = (char)(0x80 | (cp & 0x3F)); }
                tok = strtok(NULL, ",");
            }
            int type = atoi(stages[i][0]);
            double ti = now_ms();
            doc->pClass->postWindowExtTextInputEvent(doc, 0, type, utf8);
            double tp = now_ms();
            Sleep(300); // 入力は非同期に反映されるので、少し待ってから描く（待たないと 1 手前の状態が出る）
            doc->pClass->paintTile(doc, buf, cw, ch, 0, 0, cw * 15, ch * 15);
            printf("ime stage %d (type %d): post %.1f ms, repaint %.1f ms\n", i, type, tp - ti, now_ms() - tp);
            char name[MAX_PATH * 2];
            snprintf(name, sizeof name, "%s.ime%d.bmp", out_path, i);
            write_bmp(name, buf, cw, ch);
        }
    }

    free(buf);
    doc->pClass->destroy(doc);
    kit->pClass->destroy(kit);
    return 0;
}
