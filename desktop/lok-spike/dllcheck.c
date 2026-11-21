// LibreOffice の program フォルダの DLL を 1 つずつ読み込み、失敗したものとエラー番号を出す（依存の確認用）
#include <windows.h>
#include <stdio.h>

int main(int argc, char **argv) {
    const char *dir = argc > 1 ? argv[1] : "C:\\Program Files\\LibreOffice\\program";
    char pattern[MAX_PATH * 2];
    snprintf(pattern, sizeof pattern, "%s\\*.dll", dir);
    SetDllDirectoryA(dir);
    WIN32_FIND_DATAA fd;
    HANDLE h = FindFirstFileA(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) { fprintf(stderr, "no dll found\n"); return 1; }
    int ok = 0, fail = 0;
    do {
        char full[MAX_PATH * 2];
        snprintf(full, sizeof full, "%s\\%s", dir, fd.cFileName);
        SetLastError(0);
        HMODULE m = LoadLibraryExA(full, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
        if (m) { ok++; FreeLibrary(m); }
        else { fail++; printf("FAIL %-40s err=%lu\n", fd.cFileName, GetLastError()); }
    } while (FindNextFileA(h, &fd));
    FindClose(h);
    printf("ok=%d fail=%d\n", ok, fail);
    return 0;
}
