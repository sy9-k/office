; SK Hub Office のインストーラー（1 つのインストーラーで、Doc・Sheet・Slide の 3 つのアプリを選んで入れる）
;   3 つは別々のアプリ（別の実行ファイル・別のスタートメニュー・別のファイル関連付け）。インストーラーが 1 つなだけ
;   このファイルは UTF-8（BOM つき）。scripts/build-suite.mjs が makensis に渡す:
;     /DVERSION=0.1.0  /DSTAGE=<Doc.exe・Sheet.exe・Slide.exe があるフォルダ>  /DOUTFILE=<出力>  /DICON=<.ico>
;   ユーザー単位（管理者の権限は要らない）。ファイルや設定（ドキュメント\SK Hub Office・%APPDATA%\SK Hub Office）は、アンインストールしても消さない
Unicode true
!include "MUI2.nsh"
!include "FileFunc.nsh"
!include "LogicLib.nsh"

!define PRODUCT "SK Hub Office"
!define UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\SKHubOffice"

Name "${PRODUCT}"
OutFile "${OUTFILE}"
InstallDir "$LOCALAPPDATA\Programs\${PRODUCT}"
InstallDirRegKey HKCU "Software\SKHubOffice" "InstallDir"
RequestExecutionLevel user
SetCompressor /SOLID lzma
BrandingText "${PRODUCT} ${VERSION}"

VIProductVersion "${VERSION}.0"
VIAddVersionKey /LANG=1041 "ProductName" "${PRODUCT}"
VIAddVersionKey /LANG=1041 "FileDescription" "${PRODUCT} インストーラー"
VIAddVersionKey /LANG=1041 "FileVersion" "${VERSION}"
VIAddVersionKey /LANG=1041 "ProductVersion" "${VERSION}"
VIAddVersionKey /LANG=1041 "LegalCopyright" "SK Hub"

!define MUI_ICON "${ICON}"
!define MUI_UNICON "${ICON}"
!define MUI_ABORTWARNING
!define MUI_COMPONENTSPAGE_SMALLDESC
!define MUI_FINISHPAGE_NOAUTOCLOSE

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Japanese"
!insertmacro MUI_LANGUAGE "English"

; ---- アプリ 1 つぶんの登録（実行ファイル・スタートメニュー・ファイルの種類） ----
;   ファイルの種類は HKCU\Software\Classes に書く（ユーザー単位）。同じ拡張子をほかのアプリが持っていても、ここは上書きしない設計にはしない（SK Hub Office 専用の拡張子）
!macro InstallApp APP EXT PROGID TITLE DESC
  SetOutPath "$INSTDIR"
  File "/oname=SK Hub Office ${APP}.exe" "${STAGE}\${APP}.exe"
  CreateShortCut "$SMPROGRAMS\SK Hub Office ${APP}.lnk" "$INSTDIR\SK Hub Office ${APP}.exe" "" "$INSTDIR\SK Hub Office ${APP}.exe" 0
  WriteRegStr HKCU "Software\Classes\.${EXT}" "" "${PROGID}"
  WriteRegStr HKCU "Software\Classes\.${EXT}" "Content Type" "application/x-sk-hub-office-${EXT}"
  WriteRegStr HKCU "Software\Classes\${PROGID}" "" "${DESC}"
  WriteRegStr HKCU "Software\Classes\${PROGID}\DefaultIcon" "" '"$INSTDIR\SK Hub Office ${APP}.exe",0'
  WriteRegStr HKCU "Software\Classes\${PROGID}\shell\open" "FriendlyAppName" "${TITLE}"
  WriteRegStr HKCU "Software\Classes\${PROGID}\shell\open\command" "" '"$INSTDIR\SK Hub Office ${APP}.exe" "%1"'
  WriteRegStr HKCU "Software\SKHubOffice\Apps" "${APP}" "1"
!macroend

!macro UninstallApp APP EXT PROGID
  Delete "$INSTDIR\SK Hub Office ${APP}.exe"
  Delete "$SMPROGRAMS\SK Hub Office ${APP}.lnk"
  DeleteRegKey HKCU "Software\Classes\${PROGID}"
  ReadRegStr $0 HKCU "Software\Classes\.${EXT}" ""
  ${If} $0 == "${PROGID}"
    DeleteRegKey HKCU "Software\Classes\.${EXT}"
  ${EndIf}
!macroend

; ---- WebView2（画面の部品。Windows 11 には入っている。なければ、入れる） ----
Function CheckWebView2
  ReadRegStr $0 HKLM "SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" "pv"
  ${If} $0 == ""
    ReadRegStr $0 HKCU "Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" "pv"
  ${EndIf}
  ${If} $0 == ""
  ${OrIf} $0 == "0.0.0.0"
    DetailPrint "WebView2 をダウンロードしています…"
    nsExec::ExecToLog 'powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -UseBasicParsing -Uri https://go.microsoft.com/fwlink/p/?LinkId=2124703 -OutFile $$env:TEMP\MicrosoftEdgeWebview2Setup.exe"'
    ${If} ${FileExists} "$TEMP\MicrosoftEdgeWebview2Setup.exe"
      DetailPrint "WebView2 をインストールしています…"
      ExecWait '"$TEMP\MicrosoftEdgeWebview2Setup.exe" /silent /install'
      Delete "$TEMP\MicrosoftEdgeWebview2Setup.exe"
    ${Else}
      MessageBox MB_OK|MB_ICONEXCLAMATION "WebView2 を入れられませんでした。インターネットにつないで、もう一度インストールしてください。"
    ${EndIf}
  ${EndIf}
FunctionEnd

Section "-共通" SecCore
  SectionIn RO
  Call CheckWebView2
  SetOutPath "$INSTDIR"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\SKHubOffice" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayName" "${PRODUCT}"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNINST_KEY}" "Publisher" "SK Hub"
  WriteRegStr HKCU "${UNINST_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayIcon" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "${UNINST_KEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoRepair" 1
SectionEnd

Section "Doc（文書）" SecDoc
  !insertmacro InstallApp "Doc" "skdoc" "SKHubOffice.Doc" "SK Hub Office Doc" "SK Hub Office 文書"
SectionEnd

Section "Sheet（表計算）" SecSheet
  !insertmacro InstallApp "Sheet" "sksheet" "SKHubOffice.Sheet" "SK Hub Office Sheet" "SK Hub Office 表計算"
SectionEnd

Section "Slide（スライド）" SecSlide
  !insertmacro InstallApp "Slide" "skslide" "SKHubOffice.Slide" "SK Hub Office Slide" "SK Hub Office スライド"
SectionEnd

Section "-仕上げ"
  ; ファイルの関連付けを、エクスプローラーにすぐ知らせる
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
SectionEnd

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecDoc} "文書を書くアプリ（.skdoc）"
  !insertmacro MUI_DESCRIPTION_TEXT ${SecSheet} "表計算のアプリ（.sksheet）"
  !insertmacro MUI_DESCRIPTION_TEXT ${SecSlide} "スライドを作って発表するアプリ（.skslide）"
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Section "Uninstall"
  !insertmacro UninstallApp "Doc" "skdoc" "SKHubOffice.Doc"
  !insertmacro UninstallApp "Sheet" "sksheet" "SKHubOffice.Sheet"
  !insertmacro UninstallApp "Slide" "skslide" "SKHubOffice.Slide"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINST_KEY}"
  DeleteRegKey HKCU "Software\SKHubOffice"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
SectionEnd
