"""SK Hub Office が書き出した Office のファイル（.docx・.xlsx・.pptx）を、Python で確かめる

  python _tests/office/check.py [ファイルかフォルダ …]   （何も書かないと _tests/office/out/ の中を全部）

確かめること:
  ・ZIP として正しいか（CRC も）、XML がこわれていないか
  ・要素の順番（ECMA-376 の決まり）。Word・Excel・PowerPoint は順番がちがうと「読めない内容があります」になるので、
    書き出しでよく使う要素だけ、決まった順番の表（ORDER）と比べる。同じ要素が 2 回あるのも見つける
  ・.xlsx: セルの種類（t）とエラーの値、新しい関数に _xlfn. が付いているか
  ・.pptx: python-pptx で開けるか（入っていれば）
ブラウザの中のテスト（run.js）は、ZIP のしくみ・つながり（rels）・読みこみなおしを確かめる。こちらはその続き
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# 子の要素の順番。[名前, …] の中の名前は、この順で出てこないといけない。( ) でくくったものは、同じ順番（どれが先でもよい）
ORDER = {
    # ---- WordprocessingML ----
    (W, "rPr"): ["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
                 "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
                 "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
    (W, "pPr"): ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
                 "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
                 "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    (W, "sectPr"): [("headerReference", "footerReference"), "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
                    "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
    (W, "tblPr"): ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
                   "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"],
    (W, "tblBorders"): ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV"],
    (W, "tcPr"): ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "tcPrChange"],
    (W, "tcBorders"): ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl"],
    (W, "trPr"): [("cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"), "ins", "del", "trPrChange"],
    (W, "tbl"): ["bookmarkStart", "tblPr", "tblGrid", "tr"],
    (W, "tr"): ["tblPrEx", "trPr", "tc"],
    (W, "tc"): ["tcPr", "p"],
    (W, "pBdr"): ["top", "left", "bottom", "right", "between", "bar"],
    (W, "style"): ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
                   "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"],
    (W, "styles"): ["docDefaults", "latentStyles", "style"],
    (W, "lvl"): ["start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText", "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr"],
    (W, "abstractNum"): ["nsid", "multiLevelType", "tmpl", "name", "styleLink", "numStyleLink", "lvl"],
    (W, "numbering"): ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"],
    (W, "num"): ["abstractNumId", "lvlOverride"],
    # sectPr の中と settings.xml の中（こちらは区切り線の footnote を続けて書ける）
    (W, "footnotePr"): ["pos", "numFmt", "numStart", "numRestart", "footnote"],
    (W, "settings"): ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
                      "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts",
                      "saveFormsData", "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop",
                      "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate", "linkStyles",
                      "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge", "revisionView", "trackRevisions", "doNotTrackMoves",
                      "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
                      "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
                      "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
                      "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
                      "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
                      "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
                      "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
                      "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
                      "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                      "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults",
                      "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"],
    # ---- SpreadsheetML ----
    (S, "worksheet"): ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
                       "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
                       "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
                       "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"],
    (S, "styleSheet"): ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"],
    (S, "workbook"): ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr",
                      "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"],
    (S, "border"): ["start", "left", "end", "right", "top", "bottom", "diagonal", "vertical", "horizontal"],
    (S, "xf"): ["alignment", "protection", "extLst"],
    (S, "c"): ["f", "v", "is", "extLst"],
    (S, "dxf"): ["font", "numFmt", "fill", "alignment", "border", "protection", "extLst"],
    (S, "cfRule"): ["formula", "colorScale", "dataBar", "iconSet", "extLst"],
    (S, "patternFill"): ["fgColor", "bgColor"],
    (S, "sheetView"): ["pane", "selection", "pivotSelection", "extLst"],
    (S, "dataBar"): ["cfvo", "color"],
    (S, "colorScale"): [("cfvo",), "color"],
    (S, "dataValidation"): ["formula1", "formula2"],
    # ---- DrawingML・PresentationML ----
    (A, "rPr"): ["ln", ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"), ("effectLst", "effectDag"), "highlight", ("uLnTx", "uLn"), ("uFillTx", "uFill"),
                 "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"],
    (A, "endParaRPr"): ["ln", ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"), ("effectLst", "effectDag"), "highlight", ("uLnTx", "uLn"), ("uFillTx", "uFill"),
                        "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"],
    (A, "pPr"): ["lnSpc", "spcBef", "spcAft", ("buClrTx", "buClr"), ("buSzTx", "buSzPct", "buSzPts"), ("buFontTx", "buFont"), ("buNone", "buAutoNum", "buChar", "buBlip"),
                 "tabLst", "defRPr", "extLst"],
    (A, "p"): ["pPr", ("r", "br", "fld"), "endParaRPr"],
    (A, "ln"): [("noFill", "solidFill", "gradFill", "pattFill"), "prstDash", ("round", "bevel", "miter"), "headEnd", "tailEnd", "extLst"],
    (P, "spPr"): ["xfrm", ("custGeom", "prstGeom"), ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"), "ln", ("effectLst", "effectDag"), "scene3d", "sp3d", "extLst"],
    (P, "sp"): ["nvSpPr", "spPr", "style", "txBody", "extLst"],
    (P, "pic"): ["nvPicPr", "blipFill", "spPr", "style", "extLst"],
    (P, "grpSp"): ["nvGrpSpPr", "grpSpPr", ("sp", "grpSp", "graphicFrame", "cxnSp", "pic", "contentPart"), "extLst"],
    (P, "graphicFrame"): ["nvGraphicFramePr", "xfrm", "graphic", "extLst"],
    (P, "sld"): ["cSld", "clrMapOvr", "transition", "timing", "extLst"],
    (P, "notes"): ["cSld", "clrMapOvr", "extLst"],
    (P, "notesMaster"): ["cSld", "clrMap", "hf", "notesStyle", "extLst"],
    (P, "sldMaster"): ["cSld", "clrMap", "sldLayoutIdLst", "transition", "timing", "hf", "txStyles", "extLst"],
    (P, "cSld"): ["bg", "spTree", "custDataLst", "controls", "extLst"],
    (P, "presentation"): ["sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "sldIdLst", "sldSz", "notesSz", "smartTags", "embeddedFontLst", "custShowLst", "photoAlbum",
                          "custDataLst", "kinsoku", "defaultTextStyle", "modifyVerifier", "extLst"],
    (A, "tbl"): ["tblPr", "tblGrid", "tr"],
    (A, "tcPr"): ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"), "headers", "extLst"],
    (A, "tc"): ["txBody", "tcPr", "extLst"],
    (A, "bodyPr"): ["prstTxWarp", ("noAutofit", "normAutofit", "spAutoFit"), "scene3d", ("sp3d", "flatTx"), "extLst"],
}
# 1 つだけしか置けないのに、2 回あってはいけないもの（ORDER の表のうち、くり返せるもの以外）
REPEAT_OK = {"style", "tc", "tr", "p", "r", "br", "fld", "lvl", "abstractNum", "num", "lvlOverride", "headerReference", "footerReference", "cfvo", "color", "sp", "grpSp",
             "graphicFrame", "cxnSp", "pic", "contentPart", "conditionalFormatting", "numPicBullet", "formula", "ins", "del", "footnote", "activeWritingStyle", "attachedSchema", "smartTagType"}

# Excel 2010 より後の関数は、ファイルの中では _xlfn. を付ける（付けないと Excel で #NAME? になる）
XLFN = {"CONCAT", "TEXTJOIN", "IFS", "MAXIFS", "MINIFS", "SWITCH", "XLOOKUP", "XMATCH", "IFNA", "XOR", "DAYS", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "RANK.EQ", "RANK.AVG",
        "MODE.SNGL", "CEILING.MATH", "FLOOR.MATH", "NETWORKDAYS.INTL", "WORKDAY.INTL", "SEQUENCE", "FILTER", "SORT", "SORTBY", "UNIQUE", "LET", "ISOWEEKNUM", "AGGREGATE"}
XL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}


def rank_map(order):
    out = {}
    for i, x in enumerate(order):
        for n in (x if isinstance(x, tuple) else (x,)):
            out[n] = i
    return out


RANKS = {k: rank_map(v) for k, v in ORDER.items()}


def check_order(part, root, problems):
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        q = etree.QName(el)
        ranks = RANKS.get((q.namespace, q.localname))
        if not ranks:
            continue
        last, seen = -1, set()
        for ch in el:
            if not isinstance(ch.tag, str):
                continue
            cq = etree.QName(ch)
            if cq.namespace != q.namespace and cq.localname not in ("extLst",):
                continue
            r = ranks.get(cq.localname)
            if r is None:
                problems.append(f"{part}: <{q.localname}> の中に知らない要素 <{cq.localname}>（{el.sourceline} 行目）")
                continue
            if r < last:
                problems.append(f"{part}: <{q.localname}> の中の <{cq.localname}> の順番がちがいます（{el.sourceline} 行目）")
            if cq.localname in seen and cq.localname not in REPEAT_OK:
                problems.append(f"{part}: <{q.localname}> の中に <{cq.localname}> が 2 回あります（{el.sourceline} 行目）")
            seen.add(cq.localname)
            last = max(last, r)


def check_xlsx(z, parts, problems):
    for name, root in parts.items():
        if not name.startswith("xl/worksheets/"):
            continue
        for c in root.iter(f"{{{S}}}c"):
            t = c.get("t", "n")
            ref = c.get("r")
            if t not in ("b", "d", "e", "inlineStr", "n", "s", "str"):
                problems.append(f"{name} {ref}: セルの種類 t=\"{t}\" は使えません")
            v = c.find(f"{{{S}}}v")
            if t == "e" and v is not None and v.text not in XL_ERRORS:
                problems.append(f"{name} {ref}: Excel にないエラー {v.text}")
            if t == "n" and v is not None and not re.fullmatch(r"-?\d+(\.\d+)?([eE][-+]?\d+)?", v.text or ""):
                problems.append(f"{name} {ref}: 数でない値 {v.text!r}")
            f = c.find(f"{{{S}}}f")
            if f is not None and f.text:
                body = re.sub(r'"(?:[^"]|"")*"', '""', f.text)
                for m in re.finditer(r"(?<![\w.])([A-Z][A-Z0-9.]*)\s*\(", body):
                    if m.group(1) in XLFN:
                        problems.append(f"{name} {ref}: {m.group(1)} に _xlfn. が付いていません")


def check_pptx(path, problems):
    try:
        from pptx import Presentation
    except ImportError:
        return
    try:
        prs = Presentation(str(path))
        def walk(shapes):
            for shape in shapes:
                _ = shape.shape_type, shape.left, shape.top
                if shape.shape_type == 6:  # GROUP
                    walk(shape.shapes)
                if shape.has_text_frame:
                    _ = shape.text_frame.text
                if getattr(shape, "has_table", False):
                    _ = shape.table.rows
        for slide in prs.slides:
            walk(slide.shapes)
            if slide.has_notes_slide:
                _ = slide.notes_slide.notes_text_frame.text
    except Exception as e:  # noqa: BLE001  python-pptx が開けないこと自体を知らせる
        problems.append(f"python-pptx で開けません: {e!r}")


def check_xlsx_openpyxl(path, problems):
    try:
        import openpyxl
    except ImportError:
        return
    try:
        wb = openpyxl.load_workbook(str(path))
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.comment is not None:
                        _ = c.comment.text, c.comment.author
            _ = list(ws.data_validations.dataValidation)
        # グラフ: 書き出したグラフの数と、openpyxl が読めたグラフの数が同じで、どれも系列を持つ
        wrote = sum(1 for n in zipfile.ZipFile(path).namelist() if re.match(r"xl/charts/chart\d+\.xml$", n))
        charts = [ch for ws in wb.worksheets for ch in ws._charts]
        if wrote != len(charts):
            problems.append(f"グラフ {wrote} 個を書いたのに、読めたのは {len(charts)} 個です")
        for ch in charts:
            if not ch.series:
                problems.append(f"{type(ch).__name__} に系列がありません")
    except Exception as e:  # noqa: BLE001  openpyxl が開けないこと自体を知らせる
        problems.append(f"openpyxl で開けません: {e!r}")


def check_file(path):
    problems = []
    try:
        z = zipfile.ZipFile(path)
        bad = z.testzip()
        if bad:
            problems.append(f"ZIP の {bad} がこわれています（CRC）")
    except zipfile.BadZipFile as e:
        return [f"ZIP として読めません: {e}"]
    parts = {}
    for name in z.namelist():
        if not re.search(r"\.(xml|rels)$", name):
            continue
        try:
            parts[name] = etree.fromstring(z.read(name))
        except etree.XMLSyntaxError as e:
            problems.append(f"{name}: XML がこわれています: {e}")
    for name, root in parts.items():
        check_order(name, root, problems)
    if path.name == "doc-design.docx":
        # デザイン「和風」: 見出しの色・明朝・h2 の左の線・引用の線の色が styles.xml に入っている
        styles = z.read("word/styles.xml").decode("utf-8")
        for want in ['w:val="B3261E"', 'w:ascii="Yu Mincho"', "<w:pBdr><w:left", 'w:color="B3261E"', 'w:val="2A2118"']:
            if want not in styles:
                problems.append(f"styles.xml にデザインの {want} がありません")
    if path.suffix == ".xlsx":
        check_xlsx(z, parts, problems)
        check_xlsx_openpyxl(path, problems)
    if path.suffix == ".pptx":
        check_pptx(path, problems)
    return problems


def main(args):
    here = Path(__file__).resolve().parent
    targets = [Path(a) for a in args] or [here / "out"]
    files = []
    for t in targets:
        files += sorted(p for p in (t.iterdir() if t.is_dir() else [t]) if p.suffix in (".docx", ".xlsx", ".pptx"))
    ng = 0
    for p in files:
        problems = check_file(p)
        if problems:
            ng += 1
            print(f"NG  {p.name}")
            for x in problems[:40]:
                print(f"      {x}")
            if len(problems) > 40:
                print(f"      …ほか {len(problems) - 40} 件")
        else:
            print(f"OK  {p.name}")
    print(f"\n{len(files)} ファイル中 {ng} ファイルが NG" if ng else f"\nすべて OK（{len(files)} ファイル）")
    return 1 if ng or not files else 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv[1:]))
