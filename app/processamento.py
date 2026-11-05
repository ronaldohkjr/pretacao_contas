# Tratamento das imagens das notas (lógica do CLAUDE.md) + modo "inteira" (foto inteira, sem recorte)
import cv2, numpy as np
from PIL import Image, ImageOps


def load_rgb(p):
    return np.array(ImageOps.exif_transpose(Image.open(p)).convert("RGB"))


def order_pts(pts):
    r = np.zeros((4, 2), "float32"); s = pts.sum(1); d = np.diff(pts, 1)
    r[0] = pts[np.argmin(s)]; r[2] = pts[np.argmax(s)]; r[1] = pts[np.argmin(d)]; r[3] = pts[np.argmax(d)]; return r


def _warp(rgb, box):
    (tl, tr, br, bl) = box
    W = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
    H = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
    dst = np.array([[0, 0], [W - 1, 0], [W - 1, H - 1], [0, H - 1]], "float32")
    M = cv2.getPerspectiveTransform(box.astype("float32"), dst)
    return cv2.warpPerspective(rgb, M, (W, H)), W, H


def scan_effect(rgb, strong=True):
    g = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY); k = max(31, (max(g.shape) // 18) | 1)
    bg = cv2.GaussianBlur(g, (k, k), 0); n = cv2.divide(g, bg, scale=255)
    if strong:
        lo, hi = np.percentile(n, 3), np.percentile(n, 99)
        n = np.clip((n.astype(np.float32) - lo) * 255 / max(hi - lo, 1), 0, 255).astype(np.uint8)
        n = cv2.createCLAHE(1.5, (8, 8)).apply(n)
        n = cv2.addWeighted(n, 1.35, cv2.GaussianBlur(n, (0, 0), 1.1), -0.35, 0)
    else:
        lo, hi = np.percentile(n, 1), np.percentile(n, 99.5)
        n = np.clip((n.astype(np.float32) - lo) * 255 / max(hi - lo, 1), 0, 255).astype(np.uint8)
    return n


def border(p, px=45): return ImageOps.expand(p, border=px, fill=255)


# Foto de cupom: isola o papel pela "brancura" (baixa saturação + alto brilho)
def process_photo(path):
    rgb = load_rgb(path); h, w = rgb.shape[:2]; sc = 1100 / max(h, w)
    sm = cv2.resize(rgb, (int(w * sc), int(h * sc)))
    hsv = cv2.cvtColor(sm, cv2.COLOR_RGB2HSV); _, S, V = cv2.split(hsv)
    m = ((cv2.GaussianBlur(S, (7, 7), 0) < 70) & (cv2.GaussianBlur(V, (7, 7), 0) > 150)).astype(np.uint8) * 255
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, np.ones((21, 21), np.uint8), iterations=3)
    m = cv2.morphologyEx(m, cv2.MORPH_OPEN, np.ones((9, 9), np.uint8), iterations=2)
    c, _ = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if c:
        big = max(c, key=cv2.contourArea); frac = cv2.contourArea(big) / (sm.shape[0] * sm.shape[1])
        hull = cv2.convexHull(big); ap = cv2.approxPolyDP(hull, 0.02 * cv2.arcLength(hull, True), True)
        quad = ap.reshape(4, 2).astype("float32") if len(ap) == 4 else cv2.boxPoints(cv2.minAreaRect(big)).astype("float32")
        fill = cv2.contourArea(big) / max(cv2.contourArea(quad.astype("float32")), 1)
        if 0.10 < frac < 0.97 and fill > 0.60:
            warp, W, H = _warp(rgb, order_pts(quad / sc))
            if W > 200 and H > 200:
                return border(Image.fromarray(scan_effect(warp, True)))
    return Image.fromarray(scan_effect(rgb, True))


# Documento A4 em folha (fundo escuro/mesa) -> Otsu + minAreaRect (pega a pagina inteira)
def process_document(path):
    rgb = load_rgb(path); h, w = rgb.shape[:2]; sc = 1200 / max(h, w)
    sm = cv2.resize(rgb, (int(w * sc), int(h * sc)))
    g = cv2.GaussianBlur(cv2.cvtColor(sm, cv2.COLOR_RGB2GRAY), (5, 5), 0)
    _, th = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    if th.mean() > 127: th = 255 - th
    if (th > 0).mean() > 0.6: th = 255 - th
    th = cv2.morphologyEx(th, cv2.MORPH_CLOSE, np.ones((15, 15), np.uint8), iterations=3)
    c, _ = cv2.findContours(th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE); big = max(c, key=cv2.contourArea)
    warp, _, _ = _warp(rgb, order_pts(cv2.boxPoints(cv2.minAreaRect(big)) / sc))
    return border(Image.fromarray(scan_effect(warp, True)))


# Screenshot de pedagio (ja e digital): so realce leve, sem recorte
def process_screenshot(path):
    return Image.fromarray(scan_effect(load_rgb(path), False))


# Foto inteira com efeito scanner (quando o recorte corta cabeçalho/valor)
def process_full(path):
    return border(Image.fromarray(scan_effect(load_rgb(path), True)))


def render_pdf(path, dpi=150):
    import pymupdf
    pags = []
    with pymupdf.open(path) as d:
        for pg in d:
            pm = pg.get_pixmap(dpi=dpi)
            pags.append(Image.frombytes("RGB", (pm.width, pm.height), pm.samples))
    return pags


MODOS = {"cupom": process_photo, "documento": process_document, "print": process_screenshot, "inteira": process_full}


def paginas(path, modo):
    """Lista de PIL.Image (uma por página) já tratadas para o PDF."""
    if str(path).lower().endswith(".pdf"):
        return render_pdf(path)
    return [MODOS.get(modo, process_photo)(path)]
