import os
import re
import shutil
import sys

import requests
from PIL import Image

from easytrader import exceptions

VERIFY_CODE_CHARS = "0123456789abcdefghijklmnopqrstuvwxyz"


def find_tesseract_cmd() -> str:
    """定位 tesseract 可执行文件

    Windows 下 winget / 安装包默认不写 PATH(或装完终端没重启), 这里主动找常见安装位置,
    否则 pytesseract 会抛 TesseractNotFoundError
    """
    cmd = shutil.which("tesseract")
    if cmd:
        return cmd
    if sys.platform.startswith("win"):
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        candidates = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.join(local_app_data, r"Programs\Tesseract-OCR\tesseract.exe"),
            os.path.join(local_app_data, r"Tesseract-OCR\tesseract.exe"),
        ]
        for path in candidates:
            if os.path.isfile(path):
                return path
    return ""


def _ocr(im: Image.Image) -> str:
    """对二值化后的验证码图片做 OCR, 返回识别出的数字字母串"""
    import pytesseract

    tesseract_cmd = find_tesseract_cmd()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    # 验证码图很小, 放大后识别率明显更高
    im = im.resize((im.width * 3, im.height * 3), Image.LANCZOS)

    best = ""
    for psm in ("7", "8", "13"):
        res = pytesseract.image_to_string(
            im,
            config=f"--psm {psm} --oem 3 -c tessedit_char_whitelist={VERIFY_CODE_CHARS}",
        )
        res = re.sub(r"[^0-9a-zA-Z]", "", res)
        if len(res) == 4:  # 券商验证码一般是 4 位
            return res
        if len(res) > len(best):
            best = res
    return best


def captcha_recognize(img_path):
    im = Image.open(img_path).convert("L")
    # 1. threshold the image
    threshold = 200
    table = []
    for i in range(256):
        if i < threshold:
            table.append(0)
        else:
            table.append(1)

    out = im.point(table, "1")
    # 2. recognize with tesseract
    return _ocr(out)


def recognize_verify_code(image_path, broker="ht"):
    """识别验证码，返回识别后的字符串，使用 tesseract 实现
    :param image_path: 验证码图片路径
    :param broker: 券商 ['ht', 'yjb', 'gf', 'yh']
    :return: 识别出的验证码字符串"""
    if broker == "gf":
        return detect_gf_result(image_path)
    if broker in ["yh_client", "gj_client"]:
        return detect_yh_client_result(image_path)
    # 调用 tesseract 识别
    return default_verify_code_detect(image_path)


def detect_yh_client_result(image_path):
    """封装了tesseract的识别，部署在阿里云上，
    服务端源码地址为： https://github.com/shidenggui/yh_verify_code_docker"""
    api = "http://yh.ez.shidenggui.com:5000/yh_client"
    with open(image_path, "rb") as f:
        rep = requests.post(api, files={"image": f})
    if rep.status_code != 201:
        error = rep.json()["message"]
        raise exceptions.TradeError("request {} error: {}".format(api, error))
    return rep.json()["result"]


def input_verify_code_manual(image_path):
    """没有 OCR 环境时的兜底: 打开图片让人手动输入"""
    from PIL import Image

    image = Image.open(image_path)
    image.show()
    code = input(
        "image path: {}, input verify code answer:".format(image_path)
    )
    return code


def default_verify_code_detect(image_path):
    from PIL import Image

    img = Image.open(image_path)
    return invoke_tesseract_to_recognize(img)


def detect_gf_result(image_path):
    from PIL import ImageFilter, Image

    img = Image.open(image_path)
    if hasattr(img, "width"):
        width, height = img.width, img.height
    else:
        width, height = img.size
    for x in range(width):
        for y in range(height):
            if img.getpixel((x, y)) < (100, 100, 100):
                img.putpixel((x, y), (256, 256, 256))
    gray = img.convert("L")
    two = gray.point(lambda p: 0 if 68 < p < 90 else 256)
    min_res = two.filter(ImageFilter.MinFilter)
    med_res = min_res.filter(ImageFilter.MedianFilter)
    for _ in range(2):
        med_res = med_res.filter(ImageFilter.MedianFilter)
    return invoke_tesseract_to_recognize(med_res)


def invoke_tesseract_to_recognize(img):
    import pytesseract

    tesseract_cmd = find_tesseract_cmd()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    try:
        res = pytesseract.image_to_string(img)
    except FileNotFoundError:
        raise Exception(
            "tesseract 未安装，请至 https://github.com/tesseract-ocr/tesseract/wiki 查看安装教程"
        )
    except pytesseract.TesseractNotFoundError:
        raise Exception(
            "未找到 tesseract, 请安装并把 tesseract.exe 加入 PATH: "
            "winget install -e --id UB-Mannheim.TesseractOCR"
        )
    valid_chars = re.findall("[0-9a-z]", res, re.IGNORECASE)
    return "".join(valid_chars)
