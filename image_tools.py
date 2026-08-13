# -*- coding: utf-8 -*-
"""图片下载与图像处理模块（基于 Pillow）。"""
import hashlib
import io
import os
import urllib.request
from urllib.parse import urlparse

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps

FONT_CANDIDATES = [
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/arial.ttf",
]

FORMAT_EXT = {"JPG": "jpg", "PNG": "png", "WEBP": "webp"}
PIL_FORMAT = {"JPG": "JPEG", "PNG": "PNG", "WEBP": "WEBP"}


def _url_ext(url):
    ext = os.path.splitext(urlparse(url).path)[1]
    return ext if ext.lower() in (".jpg", ".jpeg", ".png", ".webp", ".bmp") else ".jpg"


def ensure_downloaded(url, cache_dir, timeout=30, retries=3):
    """下载 url 到 cache_dir（按 url 哈希命名）。已存在且非空则直接复用。返回 (path, ok)。"""
    os.makedirs(cache_dir, exist_ok=True)
    name = hashlib.sha1(url.encode("utf-8")).hexdigest()[:20] + _url_ext(url)
    path = os.path.join(cache_dir, name)
    if os.path.exists(path) and os.path.getsize(path) > 0:
        return path, True
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
            if data:
                with open(path, "wb") as fp:
                    fp.write(data)
                return path, True
        except Exception:
            if attempt == retries - 1:
                break
    return None, False


def load_image(path):
    img = Image.open(path)
    img.load()
    return img.convert("RGB")


def _crop(img, cl, cr, ct, cb):
    w, h = img.size
    left = int(w * cl / 100.0)
    right = int(w * (100 - cr) / 100.0)
    top = int(h * ct / 100.0)
    bottom = int(h * (100 - cb) / 100.0)
    if right - left < 1 or bottom - top < 1:
        return img
    return img.crop((left, top, right, bottom))


def _draw_watermark(img, text, size, opacity, pos):
    if not text or not text.strip():
        return img
    txt = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(txt)
    font = None
    for fp in FONT_CANDIDATES:
        if os.path.exists(fp):
            try:
                font = ImageFont.truetype(fp, max(int(size), 8))
                break
            except Exception:
                font = None
    if font is None:
        font = ImageFont.load_default()
    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    w, h = img.size
    margin = 20
    pos_map = {
        "左上": (margin, margin),
        "中上": ((w - tw) // 2, margin),
        "右上": (w - tw - margin, margin),
        "左中": (margin, (h - th) // 2),
        "居中": ((w - tw) // 2, (h - th) // 2),
        "右中": (w - tw - margin, (h - th) // 2),
        "左下": (margin, h - th - margin),
        "中下": ((w - tw) // 2, h - th - margin),
        "右下": (w - tw - margin, h - th - margin),
    }
    x, y = pos_map.get(pos, pos_map["右下"])
    alpha = max(0, min(255, int(255 * opacity / 100.0)))
    d.text((x - bbox[0], y - bbox[1]), text, font=font, fill=(255, 255, 255, alpha))
    return Image.alpha_composite(img.convert("RGBA"), txt).convert("RGB")


def apply_transforms(img, p):
    """按参数 p 依次处理图片；p 中缺失的键取默认值。"""
    if p.get("cl") or p.get("cr") or p.get("ct") or p.get("cb"):
        img = _crop(img, p.get("cl", 0), p.get("cr", 0), p.get("ct", 0), p.get("cb", 0))
    rot = p.get("rot", 0) % 360
    if rot:
        img = img.rotate(rot, expand=True, resample=Image.BICUBIC)
    if p.get("flip_h"):
        img = ImageOps.mirror(img)
    if p.get("flip_v"):
        img = ImageOps.flip(img)
    scale = p.get("scale", 100)
    if scale != 100 and scale > 0:
        w, h = img.size
        img = img.resize((max(1, int(w * scale / 100.0)), max(1, int(h * scale / 100.0))), Image.LANCZOS)
    img = ImageEnhance.Brightness(img).enhance(p.get("bri", 1.0))
    img = ImageEnhance.Contrast(img).enhance(p.get("con", 1.0))
    img = ImageEnhance.Color(img).enhance(p.get("sat", 1.0))
    shp = p.get("shp", 0)
    if shp:
        img = ImageEnhance.Sharpness(img).enhance(1.0 + shp / 50.0)
    blr = p.get("blr", 0)
    if blr:
        img = img.filter(ImageFilter.GaussianBlur(radius=float(blr)))
    if p.get("gray"):
        img = img.convert("L").convert("RGB")
    if p.get("inv"):
        img = ImageOps.invert(img)
    if p.get("bin"):
        thr = p.get("thr", 128)
        img = img.convert("L").point(lambda x: 255 if x >= thr else 0).convert("RGB")
    img = _draw_watermark(img, p.get("wm", ""), p.get("wmsz", 40), p.get("wmop", 40), p.get("wmpos", "右下"))
    return img


def save_image(img, path, fmt="JPG", quality=90):
    pf = PIL_FORMAT.get(str(fmt).upper(), "JPEG")
    if pf in ("JPEG", "WEBP") and img.mode != "RGB":
        img = img.convert("RGB")
    img.save(path, pf, quality=int(quality))


def image_to_bytes(img, fmt="JPG", quality=90):
    buf = io.BytesIO()
    save_image(img, buf, fmt, quality)
    return buf.getvalue()


if __name__ == "__main__":
    import time

    url = "https://b2c-wms-admin.cnbj0.mi-fds.com/b2c-wms-admin/photo/JDLD13082811738_1780121290_2_1.jpg"
    here = os.path.dirname(os.path.abspath(__file__))
    cache = os.path.join(here, ".photo_cache")
    out = os.path.join(here, "export_images")
    t0 = time.time()
    path, ok = ensure_downloaded(url, cache)
    print(f"下载 {'成功' if ok else '失败'} {path} ({time.time()-t0:.1f}s)")
    if ok:
        img = load_image(path)
        params = {"rot": 15, "flip_h": True, "scale": 80, "bri": 1.2, "con": 1.1,
                  "gray": False, "bin": False, "thr": 128, "wm": "测试水印", "wmsz": 40,
                  "wmop": 50, "wmpos": "右下", "fmt": "JPG", "q": 90}
        p = apply_transforms(img, params)
        os.makedirs(out, exist_ok=True)
        save_image(p, os.path.join(out, "_test_output.jpg"), "JPG", 90)
        print(f"处理保存完成，新尺寸: {p.size}")
