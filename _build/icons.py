# SK Hub Office のアプリのアイコン（PWA・ホーム画面・ショートカット・製品ページ）を作るスクリプト
#   python _build/icons.py
# 元の絵は icons/src/ の 1024px の PNG（sk-doc・sk-sheet・sk-slide。六角形のアプリアイコン。Office のアイコンは、この 3 つを扇形に合わせて作る。sk-office.png は旧版）。
# 書き出すもの（icons/）:
#   icon-192.png・icon-512.png         … ふつうのアイコン（白い角丸の板）
#   maskable-192.png・maskable-512.png … Android などで形を切り抜かれても大丈夫なアイコン（白い四角いっぱい）
#   apple-touch-icon.png（180）        … iPhone・iPad のホーム画面
#   doc-96.png・sheet-96.png・slide-96.png … アプリのショートカット（manifest の shortcuts）
#   doc-192/512.png・doc-maskable-192/512.png・doc-apple.png（Sheet・Slide も同じ） … アプリごとの PWA のアイコン
#   favicon-32.png                     … ファビコン
#   {office,doc,sheet,slide}-logo.png  … アプリ内のロゴ（六角形だけ・透明。192）
#   office-part-{sheet,slide,doc}.png  … Office のアイコンを 3 つに分けたもの（起動画面。192）
#   {office,doc,sheet,slide}-file.png  … ファイル一覧のアイコン（紙に六角形のバッジ・透明。192。元は sk-○○-file.png）
import pathlib
from PIL import Image, ImageDraw

REPO = pathlib.Path(__file__).resolve().parents[1]
SRC = REPO / "icons" / "src"
OUT = REPO / "icons"
OUT.mkdir(parents=True, exist_ok=True)
WHITE = (255, 255, 255, 255)


CENTER = (512, 512)  # 六角形の真ん中（Doc・Sheet・Slide の絵は同じ形・同じ位置）
HEX_BOX = (197, 169, 827, 855)  # 六角形のほんたい（影はふくまない）


# (アプリ, 回す角度, 入れる扇形)。もとの絵の下（270°）が、上（90°）・左下（210°）・右下（330°）に来る
OFFICE_PARTS = (("sheet", 180, (30, 150)), ("slide", -60, (150, 270)), ("doc", 60, (270, 390)))


def wedge(angles, r=1400):
    """真ん中から見て、angles（度。数学の向き）のあいだだけを白くしたマスク。上・右下・左下の 3 つに分けるときに使う"""
    import math
    mask = Image.new("L", (1024 * 2, 1024 * 2), 0)
    pts = [(CENTER[0] * 2, CENTER[1] * 2)]
    a0, a1 = angles
    n = 24
    for i in range(n + 1):
        a = math.radians(a0 + (a1 - a0) * i / n)
        pts.append((CENTER[0] * 2 + r * 2 * math.cos(a), CENTER[1] * 2 - r * 2 * math.sin(a)))
    ImageDraw.Draw(mask).polygon(pts, fill=255)
    return mask.resize((1024, 1024), Image.LANCZOS)


def office_parts():
    """office_mark を 3 つの扇形に分けたもの（起動画面で、3 つが組み上がるのに使う）。名前 → 絵"""
    parts = {}
    for name, turn, angles in OFFICE_PARTS:
        src = Image.open(SRC / f"sk-{name}.png").convert("RGBA").rotate(turn, resample=Image.BICUBIC, center=CENTER)
        piece = Image.new("RGBA", src.size, (0, 0, 0, 0))
        piece.paste(src, (0, 0), wedge(angles))
        parts[name] = piece
    return parts


def office_mark():
    """Office のアイコン。Doc（青）・Sheet（緑）・Slide（オレンジ）の六角形のうち、記号のない下の 120° を、それぞれ上・左下・右下に回して 1 つにしたもの（六角形は 60° ずつ回しても同じ形）"""
    out = Image.new("RGBA", (1024, 1024), (0, 0, 0, 0))
    for piece in office_parts().values():
        out.alpha_composite(piece)
    return out

def office_file():
    """Office のファイルアイコン。sk-office-file.png の紙に、office_mark の六角形をバッジとして重ねる"""
    paper = Image.open(SRC / "sk-office-file.png").convert("RGBA")
    x0, y0, x1, y1 = HEX_BOX
    m = 70  # 影のぶん
    badge = office_mark().crop((x0 - m, y0 - m, x1 + m, y1 + m))
    k = 350 / (y1 - y0)
    badge = badge.resize((round(badge.width * k), round(badge.height * k)), Image.LANCZOS)
    cx, cy = 758, 735
    paper.alpha_composite(badge, (cx - badge.width // 2, cy - badge.height // 2))
    return paper


def source(name):
    return office_mark() if name == "office" else Image.open(SRC / f"sk-{name}.png").convert("RGBA")


def load(name):
    """元の絵を白い地に敷いて返す（六角形のまわりの影もそのまま残す）"""
    src = source(name)
    base = Image.new("RGBA", src.size, WHITE)
    base.alpha_composite(src)
    return base


def icon(art, size, maskable=False, rounded=True):
    """白い角丸の板（maskable・rounded=False は四角いっぱい）に絵を入れる。六角形はもともと真ん中の 7 割に収まっている"""
    img = art.resize((size, size), Image.LANCZOS)
    if maskable or not rounded:
        return img
    mask = Image.new("L", (size * 4, size * 4), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, size * 4 - 1, size * 4 - 1], radius=size * 4 * 0.24, fill=255)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(img, (0, 0), mask.resize((size, size), Image.LANCZOS))
    return out


def tight(art, size):
    """ファビコン用に、六角形だけを大きく切り出す"""
    w = art.convert("RGB")
    bbox = Image.eval(w, lambda v: 255 - v).convert("L").point(lambda v: 255 if v > 24 else 0).getbbox()
    x0, y0, x1, y1 = bbox
    side = max(x1 - x0, y1 - y0)
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    box = (cx - side // 2, cy - side // 2, cx + side // 2, cy + side // 2)
    return art.crop(box).resize((size, size), Image.LANCZOS)


def trim(im, size, like=None):
    """透明な余白を切り落として、正方形の透明な PNG にする（影はそのまま残す）。like を渡すと、その絵と同じ切り方にする"""
    x0, y0, x1, y1 = (like or im).getchannel("A").getbbox()
    side = max(x1 - x0, y1 - y0)
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    return im.crop((cx - side // 2, cy - side // 2, cx + side // 2, cy + side // 2)).resize((size, size), Image.LANCZOS)


def write_set(art, prefix):
    for n in (192, 512):
        icon(art, n).save(OUT / (f"icon-{n}.png" if prefix == "" else f"{prefix}-{n}.png"))
        icon(art, n, maskable=True).save(OUT / (f"maskable-{n}.png" if prefix == "" else f"{prefix}-maskable-{n}.png"))
    icon(art, 180, rounded=False).convert("RGB").save(OUT / ("apple-touch-icon.png" if prefix == "" else f"{prefix}-apple.png"))


office = load("office")
write_set(office, "")
tight(office, 32).save(OUT / "favicon-32.png")

for kind in ("doc", "sheet", "slide"):
    art = load(kind)
    write_set(art, kind)
    icon(art, 96).save(OUT / f"{kind}-96.png")

for kind in ("office", "doc", "sheet", "slide"):
    trim(source(kind), 192).save(OUT / f"{kind}-logo.png")
    trim(office_file() if kind == "office" else Image.open(SRC / f"sk-{kind}-file.png").convert("RGBA"), 192).save(OUT / f"{kind}-file.png")

mark = office_mark()
for name, piece in office_parts().items():
    trim(piece, 192, like=mark).save(OUT / f"office-part-{name}.png")  # 起動画面用

print("wrote", ", ".join(sorted(p.name for p in OUT.iterdir())))
