# preprocess.py
from PIL import Image, ImageOps, ImageFilter
import numpy as np

def resize_image(image, scale_factor=3.0):
    w, h = image.size
    new_size = (int(w * scale_factor), int(h * scale_factor))
    return image.resize(new_size, Image.LANCZOS)

def auto_gamma(img, clamp=(0.6, 1.4)):
    arr = np.asarray(img, dtype=np.float32)
    mean = arr.mean() / 255.0 + 1e-6
    target = 0.5
    import math
    gamma = math.log(target) / math.log(mean)
    gamma = max(clamp[0], min(clamp[1], gamma))
    lut = [int(((i/255.0) ** gamma) * 255 + 0.5) for i in range(256)]
    return img.point(lut)

def enhance_image_for_ocr(image):
    # 그레이스케일
    g = image.convert('L')

    # 배경 밝기 추정 후 어두우면 반전
    arr = np.array(g)
    corners = np.hstack([
        arr[:50, :50].ravel(),
        arr[:50, -50:].ravel(),
        arr[-50:, :50].ravel(),
        arr[-50:, -50:].ravel()
    ])
    if np.median(corners) < 110:
        g = ImageOps.invert(g)

    # 동적 감마
    g = auto_gamma(g, clamp=(0.6, 1.4))

    # autocontrast + 샤픈
    g = ImageOps.autocontrast(g, cutoff=2)
    g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=2))
    return g
