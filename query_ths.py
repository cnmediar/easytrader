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

表格读取策略(--strategy, 默认 wmcopy):
- wmcopy: 走剪贴板(WM_COMMAND 0xE122 复制)。本客户端有"拷贝保护", 每次读表格都会弹
  「检测到您正在拷贝数据」验证码弹窗, 由 easytrader 的 OCR 自动填写并重试
- xls: 走 Ctrl+S 另存为文件, 不碰剪贴板。但本客户端"保存"同样受保护——Ctrl+S 弹出的也是
  「检测到您正在保存数据」验证码弹窗(控件与拷贝保护弹窗相同: 图片 0x965 / 输入框 0x964),
  而不是文件另存为对话框, 所以这里并没有更快, 仅作备用/它用
- auto: 先试 xls, 失败自动回退 wmcopy
"""
import argparse
import time

import win32con
import win32gui
from pywinauto import Desktop, findwindows

import easytrader
from easytrader.grid_strategies import WMCopy, Xls
from easytrader.utils.captcha import find_tesseract_cmd

ths_exe_path = r"D:\同花顺软件\同花顺\xiadan.exe"
MAIN_WINDOW_TITLE = "网上股票交易系统5.0"
CONNECT_TIMEOUT = 60

# 表格读取策略: xls(另存为文件, 不走剪贴板) / wmcopy(走剪贴板, 会触发拷贝保护验证码) /
# auto(先试 xls, 失败自动回退 wmcopy). 用 --strategy=xls 覆盖
parser = argparse.ArgumentParser(description="同花顺信用账户查询")
parser.add_argument(
    "--strategy",
    choices=["auto", "xls", "wmcopy"],
    default="wmcopy",
    help="表格读取策略(默认 wmcopy: 本客户端保存也弹验证码, xls 无速度优势)",
)
args = parser.parse_args()


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
class VisibleGridMixin:
    """选可见且面积最大的 grid(覆盖默认"取第一个"的行为)"""

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


class VisibleGridWMCopy(VisibleGridMixin, WMCopy):
    """走剪贴板: 会触发客户端"拷贝保护"验证码弹窗, 需要 OCR"""


class VisibleGridXls(VisibleGridMixin, Xls):
    """走另存为文件(Ctrl+S): 不碰剪贴板, 理论上不弹验证码"""


class AutoGridStrategy:
    """先试 Xls; 若该客户端不支持 Ctrl+S 导出则自动回退 WMCopy"""

    def __init__(self):
        self._trader = None
        self._strategy = None

    def set_trader(self, trader):
        self._trader = trader

    def get(self, control_id: int):
        if self._strategy is None:
            self._use(VisibleGridXls())
        try:
            return self._strategy.get(control_id)
        except Exception as ex:  # noqa: BLE001
            if isinstance(self._strategy, VisibleGridXls):
                print(
                    f"Xls 策略读取失败({type(ex).__name__}: {ex}), 自动回退剪贴板策略"
                )
                close_all_popups()
                self._use(VisibleGridWMCopy())
                return self._strategy.get(control_id)
            raise

    def _use(self, strategy):
        self._strategy = strategy
        self._strategy.set_trader(self._trader)


def make_grid_strategy(name: str):
    if name == "xls":
        return VisibleGridXls()
    if name == "wmcopy":
        return VisibleGridWMCopy()
    return AutoGridStrategy()


user.grid_strategy = make_grid_strategy(args.strategy)
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


def close_all_popups():
    """关掉残留的另存为/报错弹窗, 避免挡住后续操作(仅用于策略回退前的清理)"""
    for win in user.app.windows(visible_only=True):
        try:
            title = win.window_text()
            if win.element_info.class_name != "#32770" and "另存为" not in title:
                continue
        except Exception:  # noqa: BLE001
            continue
        try:
            win.close()
            user.wait(0.2)
        except Exception:  # noqa: BLE001
            pass


def drop_empty_rows(rows):
    """客户端空表格会带若干空白行, 过滤掉"""
    if not isinstance(rows, list):
        return rows
    return [
        row
        for row in rows
        if any(str(v).strip() not in ("", "0", "0.0") for v in row.values())
    ]


def current_account():
    """从客户端状态栏读当前登录的资金账号

    easytrader 的 ths 没有 login/prepare, 挂到哪个窗口就是哪个账号, 所以切换后用它确认。
    客户端是自绘界面(GetMenu 返回 0, 无标准菜单), 只能按文本特征从 Static 里筛:
    账号形如 mncg116(纯 ASCII 字母数字, 长度 5~16, 不是纯数字)
    """
    for c in user.main.descendants(class_name="Static"):
        try:
            text = (c.window_text() or "").strip()
        except Exception:  # noqa: BLE001
            continue
        if (
            5 <= len(text) <= 16
            and text.isascii()
            and text.isalnum()
            and not text.isdigit()
        ):
            return text
    return None


def safe(getter):
    """单步失败不影响后续查询"""
    try:
        return drop_empty_rows(getter())
    except Exception as ex:  # noqa: BLE001
        return f"<读取失败: {type(ex).__name__}: {ex}>"


if __name__ == "__main__":
    print(f"连接成功！表格读取策略: {args.strategy}")
    print("当前登录账号:", current_account())
    close_copy_captcha()
    print("资金信息:", safe(lambda: user.balance))
    print("持仓信息:", safe(lambda: user.position))
    print("当日委托:", safe(lambda: user.today_entrusts))
    print("当日成交:", safe(lambda: user.today_trades))
    print("可撤委托:", safe(lambda: user.cancel_entrusts))
