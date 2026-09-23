# -*- coding: utf-8 -*-
"""
同花顺信用账户(银河证券)客户端查询脚本

前提:
1. 手动打开 D:\\同花顺软件\\同花顺\\xiadan.exe 并登录交易账户
   (登录窗口有安全控件/验证码, 无法自动化, 只能用 connect() 连接已登录的客户端)
2. 运行: python -X utf8 query_ths.py

适配要点(本机信用账户客户端):
- 左侧菜单树实测为: 查询[F4] -> 资金股票 / 当日成交 / 当日委托 / 历史成交 / 历史委托 /
  资金明细 / 对帐单, 与 easytrader 默认一致, 无需覆盖 *_MENU_PATH
- 资金控件 ID 也和默认一致(可用金额=1016, 股票市值=1014, 总资产=1015, 可取金额=1017),
  无需覆盖 BALANCE_CONTROL_ID_GROUP
- 表格读取必须用 Xls 策略(把 grid 另存为 xls 再读)。客户端有"拷贝保护": 只要走剪贴板
  (Copy / WMCopy) 连续复制就会弹"检测到您正在拷贝数据, 请先输入验证码"弹窗并卡死后续所有
  查询; Xls 走的是另存文件, 不触发该保护
- 交易表单(买入/卖出)的自绘输入框 set_edit_text 写不进去, 必须先调用
  enable_type_keys_for_editor() 用真实键盘输入 (window_text() 读不回内容属正常现象,
  写入是否成功需截图确认)
- connect 前会等待并还原主窗口: xiadan.exe 处于登录框阶段/最小化到托盘/主窗口未就绪时,
  pywinauto 的 connect(path=) 会报 "No windows for that process could be found"
"""
import time

import win32con
import win32gui
from pywinauto import Desktop, findwindows

import easytrader
from easytrader.grid_strategies import WMCopy
from easytrader.utils.captcha import find_tesseract_cmd

ths_exe_path = r"D:\同花顺软件\同花顺\xiadan.exe"
MAIN_WINDOW_TITLE = "网上股票交易系统5.0"
CONNECT_TIMEOUT = 60


def wait_client_window(title=MAIN_WINDOW_TITLE, timeout=CONNECT_TIMEOUT):
    """等待客户端主窗口出现并返回其 PID; 若窗口被隐藏/最小化到托盘则先还原.

    pywinauto 默认只枚举可见窗口(visible_only=True), 窗口隐藏时 connect(path=)
    能找到进程但拿不到窗口, 会抛 RuntimeError: No windows for that process could be found
    """
    deadline = time.time() + timeout
    while True:
        for elem in findwindows.find_elements(top_level_only=True, visible_only=False):
            if elem.name == title:
                hwnd = elem.handle
                if not win32gui.IsWindowVisible(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                return elem.process_id
        if time.time() > deadline:
            raise RuntimeError(
                f"找不到标题为 {title!r} 的客户端窗口, 请手动打开 {ths_exe_path} 并完成登录"
            )
        time.sleep(1)


pid = wait_client_window()

user = easytrader.use('ths')
user.connect(exe_path=ths_exe_path)
print(f"已连接客户端 PID={pid}")

# --- 信用账户客户端适配 ---
# 该客户端同一时刻存在 5 个 control_id=1047 的 grid(每个页面各一个, 只有当前页面那个可见),
# easytrader 默认取第一个(隐藏的小 grid) => 读出来是空表/报错 ElementAmbiguousError。
# 这里改为"取可见且面积最大的那个"
class VisibleGridWMCopy(WMCopy):
    def _get_grid(self, control_id: int):
        try:
            candidates = findwindows.find_elements(
                class_name="CVirtualGridCtrl",
                control_id=control_id,
                top_level_only=False,
                visible_only=True,
                parent=self._trader.main.element_info.handle,
            )
            if candidates:
                biggest = max(
                    candidates,
                    key=lambda e: e.rectangle.width() * e.rectangle.height(),
                )
                return Desktop(backend="win32").window(handle=biggest.handle)
        except Exception:  # noqa: BLE001
            pass
        return self._trader.main.child_window(
            control_id=control_id, class_name="CVirtualGridCtrl"
        )


user.grid_strategy = VisibleGridWMCopy()
# 如需下单(buy/sell), 先开启键盘输入模式:
# user.enable_type_keys_for_editor()

# WMCopy/Copy 策略自带验证码 OCR(easytrader.grid_strategies 里识别 0x965 图片 + 0x964 输入框),
# 依赖 tesseract; 没装时退化为"弹图 + 手动输入", 否则会抛 TesseractNotFoundError
if not find_tesseract_cmd():
    from easytrader import grid_strategies
    from easytrader.utils.captcha import input_verify_code_manual

    grid_strategies.captcha_recognize = input_verify_code_manual
    print("提示: 未检测到 tesseract, 出现验证码弹窗时会弹图要求手动输入;")
    print("      想全自动请安装后重开终端: winget install -e --id UB-Mannheim.TesseractOCR")


def close_copy_captcha():
    """兜底关掉拷贝/保存保护验证码弹窗(残留时会让后续表格读取全部失败)"""
    for win in user.app.windows(class_name="#32770", visible_only=True):
        try:
            texts = [c.window_text() for c in win.descendants(class_name="Static")]
        except Exception:  # noqa: BLE001
            continue
        if any("正在拷贝数据" in t or "正在保存数据" in t for t in texts if t):
            win.child_window(control_id=2, class_name="Button").click()
            print("已关闭保护验证码弹窗")
            user.wait(0.5)


def drop_empty_rows(rows):
    """客户端空表格会带若干空白行, 过滤掉"""
    if not isinstance(rows, list):
        return rows
    return [
        row
        for row in rows
        if any(str(v).strip() not in ("", "0", "0.0") for v in row.values())
    ]


def safe(getter):
    """单步失败不影响后续查询"""
    try:
        return drop_empty_rows(getter())
    except Exception as ex:  # noqa: BLE001
        return f"<读取失败: {type(ex).__name__}: {ex}>"


if __name__ == "__main__":
    print("连接成功！")
    close_copy_captcha()
    print("资金信息:", safe(lambda: user.balance))
    print("持仓信息:", safe(lambda: user.position))
    print("当日委托:", safe(lambda: user.today_entrusts))
    print("当日成交:", safe(lambda: user.today_trades))
    print("可撤委托:", safe(lambda: user.cancel_entrusts))
