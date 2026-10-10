# -*- coding: utf-8 -*-
"""导出通达信“通讯设置”中的“高级行情产品”数据。

使用前请先启动并登录通达信客户端。不同版本的菜单文字可能不同，
可以先手动打开目标页面，再运行本脚本以获得最稳定的结果。
"""

import argparse
import csv
import json
import re
import subprocess
import sys
import time
from pathlib import Path

from pywinauto import Desktop


def text_of(control):
    """读取控件文本，兼容 UIA 控件的不同实现。"""
    try:
        text = control.window_text()
    except Exception:
        text = ""
    return (text or "").strip()


def find_window(title_pattern):
    pattern = re.compile(title_pattern, re.IGNORECASE)
    windows = Desktop(backend="win32").windows(visible_only=True)
    matches = [window for window in windows if pattern.search(text_of(window))]
    if not matches:
        raise RuntimeError(
            "找不到目标窗口，请确认通讯设置已经打开，或使用 --title-regex 指定标题。"
        )
    return matches[0]


def click_text(window, wanted_text):
    """点击窗口中第一个匹配文字的控件。"""
    controls = [window] + list(window.descendants())
    for control in controls:
        if text_of(control) == wanted_text:
            try:
                control.click_input()
                return True
            except Exception:
                continue
    return False


def collect_controls(window, keyword):
    """收集目标区域中的控件文本和表格行。"""
    controls = [window] + list(window.descendants())
    keyword_controls = [
        control for control in controls if keyword in text_of(control)
    ]
    root = keyword_controls[0] if keyword_controls else window

    records = []
    for control in [root] + list(root.descendants()):
        control_text = text_of(control)
        if not control_text:
            continue
        try:
            control_type = control.element_info.control_type or ""
            class_name = control.element_info.class_name or ""
        except Exception:
            control_type = ""
            class_name = ""
        records.append(
            {
                "text": control_text,
                "control_type": control_type,
                "class_name": class_name,
            }
        )

    rows = []
    row_types = {"ListItem", "DataItem", "TreeItem", "Row"}
    for control in [root] + list(root.descendants()):
        try:
            control_type = control.element_info.control_type
        except Exception:
            continue
        if control_type not in row_types:
            continue
        cells = [text_of(child) for child in control.descendants()]
        cells = [cell for cell in cells if cell]
        if not cells and text_of(control):
            cells = [text_of(control)]
        if cells and cells not in rows:
            rows.append(cells)

    return {
        "keyword": keyword,
        "window_title": text_of(window),
        "rows": rows,
        "controls": records,
    }


def extract_advanced_server(window):
    """打开主站设置窗口并提取高级行情服务器地址。"""
    combos = window.descendants(class_name="ComboBox")
    combo = next(
        (item for item in combos if "高级行情" in text_of(item)), None
    )
    if combo is None:
        raise RuntimeError("找不到高级行情主站下拉框。")

    settings_button = next(
        (
            item
            for item in window.descendants(class_name="Button")
            if text_of(item) == "主站设置"
        ),
        None,
    )
    if settings_button is None:
        raise RuntimeError("找不到“主站设置”按钮。")
    before_handles = {
        item.handle
        for item in Desktop(backend="win32").windows(visible_only=True)
    }
    settings_button.click_input()
    time.sleep(0.3)

    dialogs = [
        item
        for item in Desktop(backend="win32").windows(visible_only=True)
        if item.handle not in before_handles
        and item.handle != window.handle
        and item.window_text()
    ]
    dialog = (
        next((item for item in dialogs if "主站" in text_of(item)), dialogs[0])
        if dialogs
        else window
    )
    controls = [dialog] + list(dialog.descendants())
    values = []
    for control in controls:
        value = text_of(control)
        if value:
            values.append(value)

    addresses = []
    address_pattern = re.compile(
        r"(?:(?:\d{1,3}\.){3}\d{1,3}|[a-zA-Z][a-zA-Z0-9.-]+)"
        r"(?::\d{1,6})?"
    )
    for value in values:
        if "." in value or ":" in value:
            addresses.extend(address_pattern.findall(value))

    if not addresses:
        process_id = window.process_id()
        result = subprocess.run(
            ["netstat", "-ano", "-p", "tcp"],
            capture_output=True,
            text=True,
            encoding="ascii",
            errors="ignore",
            check=False,
        )
        pid_text = str(process_id)
        for line in result.stdout.splitlines():
            fields = line.split()
            if len(fields) >= 5 and fields[0].upper() == "TCP":
                if fields[3].upper() == "ESTABLISHED" and fields[4] == pid_text:
                    addresses.append(fields[2])

    return {
        "server_name": text_of(combo),
        "dialog_title": text_of(dialog),
        "values": values,
        "addresses": sorted(set(addresses)),
    }


def write_output(data, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.with_suffix(".json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    rows = data["rows"]
    if not rows:
        rows = [
            [item["text"], item["control_type"], item["class_name"]]
            for item in data["controls"]
        ]
    if rows:
        width = max(len(row) for row in rows)
        with output_path.with_suffix(".csv").open(
            "w", newline="", encoding="utf-8-sig"
        ) as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(row + [""] * (width - len(row)) for row in rows)


def main():
    parser = argparse.ArgumentParser(
        description="导出通达信通讯设置中的高级行情产品数据"
    )
    parser.add_argument("--title-regex", default=r"^通讯设置$", help="目标窗口标题正则表达式")
    parser.add_argument("--keyword", default="高级行情产品")
    parser.add_argument(
        "--get-ip",
        action="store_true",
        help="打开高级行情主站设置并提取 IP/端口",
    )
    parser.add_argument(
        "--open-dialog",
        action="store_true",
        help="尝试点击当前窗口中的“通讯设置”",
    )
    parser.add_argument("--output", default="tdx_advanced_quotes")
    args = parser.parse_args()

    window = find_window(args.title_regex)
    window.set_focus()
    if args.open_dialog and "通讯设置" not in text_of(window):
        if not click_text(window, "通讯设置"):
            raise RuntimeError(
                "找不到“通讯设置”菜单项，请手动打开通讯设置后重试。"
            )
        time.sleep(1)
        window = find_window(r"^通讯设置$")

    data = collect_controls(window, args.keyword)
    if args.get_ip:
        data["advanced_server"] = extract_advanced_server(window)
    write_output(data, args.output)
    print(
        "已导出 {} 条表格行，控件明细 {} 条。文件：{}.json / {}.csv".format(
            len(data["rows"]),
            len(data["controls"]),
            args.output,
            args.output,
        )
    )
    if args.get_ip:
        print("高级行情 IP/地址：{}".format(data["advanced_server"]["addresses"]))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("已取消。")
    except Exception as error:
        sys.exit("错误：{}".format(error))