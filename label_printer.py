"""
label_printer.py — 华耕嘉成 发货贴纸打印工具
独立 exe，不依赖主程序数据库
所有参数从同目录 config.json 读取，首次启动自动创建
"""

import json
import math
import os
import sys
import tempfile
import threading
from pathlib import Path

import wx

import region_data

APP_TITLE = "华耕嘉成 · 发货贴纸打印"

# ── config.json 默认值及说明 ──────────────────────────────
DEFAULT_CONFIG = {
    "_说明": {
        "paper_width_mm":    "贴纸宽度（毫米），默认150",
        "paper_height_mm":   "贴纸高度（毫米），默认100",
        "location_x_mm":     "收货地址文字的左边距（毫米），从贴纸左边缘算起",
        "location_y_mm":     "收货地址文字的上边距（毫米），从贴纸上边缘算起",
        "box_info_x_mm":     "箱号文字的左边距（毫米），从贴纸左边缘算起",
        "box_info_y_mm":     "箱号文字的上边距（毫米），从贴纸上边缘算起",
        "font_size_location": "收货地址文字大小（磅/pt）",
        "font_size_box":      "箱号文字大小（磅/pt）",
        "window_width":       "主窗口宽度（像素）",
        "window_height":      "主窗口高度（像素）"
    },
    "paper_width_mm":     150,
    "paper_height_mm":    100,
    "location_x_mm":       20,
    "location_y_mm":       70,
    "box_info_x_mm":       20,
    "box_info_y_mm":       80,
    "font_size_location":  14,
    "font_size_box":       14,
    "window_width":       540,
    "window_height":      500
}


def _config_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "config.json"
    return Path(__file__).parent / "config.json"


def load_config() -> dict:
    """读取 config.json，不存在则创建默认值"""
    path = _config_path()
    if not path.exists():
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        return dict(DEFAULT_CONFIG)
    try:
        with open(path, "r", encoding="utf-8") as f:
            saved = json.load(f)
        # 用默认值补全缺失键
        result = dict(DEFAULT_CONFIG)
        for k, v in saved.items():
            result[k] = v
        return result
    except Exception:
        return dict(DEFAULT_CONFIG)


# 全局配置，启动时加载一次
CFG: dict = {}


# ── 打印核心 ──────────────────────────────────────────────

def _do_print_label(location: str, box_str: str) -> tuple[bool, str]:
    """生成 PDF 并打印，返回 (success, errmsg)"""
    try:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as rl_canvas

        pw = CFG["paper_width_mm"]  * mm
        ph = CFG["paper_height_mm"] * mm
        lx = CFG["location_x_mm"]  * mm
        ly = CFG["location_y_mm"]  * mm
        bx = CFG["box_info_x_mm"]  * mm
        by = CFG["box_info_y_mm"]  * mm
        fs_loc = CFG["font_size_location"]
        fs_box = CFG["font_size_box"]

        # PDF 坐标原点在左下角，y 需翻转
        def flip(y_from_top): return ph - y_from_top

        _register_font()

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_path = tmp.name
        tmp.close()

        c = rl_canvas.Canvas(tmp_path, pagesize=(pw, ph))
        c.setFont(_FONT_NAME, fs_loc)
        c.drawString(lx, flip(ly), location)
        c.setFont(_FONT_NAME, fs_box)
        c.drawString(bx, flip(by), box_str)
        c.save()

        ok, msg = _send_to_printer(tmp_path)
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return ok, msg

    except ImportError:
        return _fallback_print(location, box_str)
    except Exception as e:
        return False, str(e)


_FONT_NAME = "HWMSFont"
_font_registered = False


def _register_font():
    global _font_registered
    if _font_registered:
        return
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    candidates = [
        Path(getattr(sys, "_MEIPASS", "")) / "NotoSansSC-Regular.ttf",
        Path(__file__).parent / "NotoSansSC-Regular.ttf",
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc"),
        Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
    ]
    for p in candidates:
        if p.exists():
            try:
                pdfmetrics.registerFont(TTFont(_FONT_NAME, str(p)))
                _font_registered = True
                return
            except Exception:
                continue
    _font_registered = True


def _send_to_printer(pdf_path: str) -> tuple[bool, str]:
    import platform
    try:
        if platform.system() == "Windows":
            import win32api, win32print
            printer = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", pdf_path,
                                  f'/d:"{printer}"', ".", 0)
            return True, ""
        else:
            ret = os.system(f'lp "{pdf_path}"')
            return ret == 0, "" if ret == 0 else "lp 命令失败"
    except Exception as e:
        return False, str(e)


def _fallback_print(location: str, box_str: str) -> tuple[bool, str]:
    try:
        import platform
        tmp = tempfile.NamedTemporaryFile(
            suffix=".txt", delete=False, mode="w", encoding="utf-8"
        )
        tmp.write(f"{location}\n{box_str}\n")
        tmp_path = tmp.name
        tmp.close()
        if platform.system() == "Windows":
            import win32api, win32print
            printer = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", tmp_path,
                                  f'/d:"{printer}"', ".", 0)
        else:
            os.system(f'lp "{tmp_path}"')
        return True, ""
    except Exception as e:
        return False, str(e)


# ── 打印确认窗口 ──────────────────────────────────────────

class PrintDialog(wx.Dialog):
    """每点一次打印一张，打完最后一张自动关闭"""

    def __init__(self, parent, location: str, total_boxes: int):
        super().__init__(parent, title="打印发货贴纸",
                         style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP)
        self.location    = location
        self.total_boxes = total_boxes
        self.current     = 0

        panel = wx.Panel(self)
        vbox  = wx.BoxSizer(wx.VERTICAL)

        self.lbl = wx.StaticText(panel, label=self._label_text(),
                                 style=wx.ALIGN_CENTER)
        self.lbl.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT,
                                 wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))

        self.btn = wx.Button(panel, label="打  印", size=(160, 55))
        self.btn.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT,
                                 wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        self.btn.SetBackgroundColour(wx.Colour(34, 139, 230))
        self.btn.SetForegroundColour(wx.WHITE)

        vbox.Add(self.lbl, 0, wx.ALL | wx.EXPAND, 20)
        vbox.Add(self.btn, 0, wx.ALL | wx.ALIGN_CENTER, 16)
        panel.SetSizer(vbox)
        vbox.Fit(panel)
        self.SetClientSize(panel.GetBestSize())
        self.Centre()

        self.btn.Bind(wx.EVT_BUTTON, self._on_print)

    def _label_text(self) -> str:
        done = self.current
        left = self.total_boxes - done
        return (
            f"发货地：{self.location}\n"
            f"总箱数：{self.total_boxes}    已打：{done}    待打：{left}\n\n"
            f"下一张：{self.location}  {self.total_boxes}-{done + 1}"
        )

    def _on_print(self, _):
        self.current += 1
        box_str = f"{self.total_boxes}-{self.current}"
        self.btn.Disable()
        self.btn.SetLabel("打印中...")

        def do():
            ok, err = _do_print_label(self.location, box_str)
            wx.CallAfter(self._after_print, ok, err)

        threading.Thread(target=do, daemon=True).start()

    def _after_print(self, ok: bool, err: str):
        if not ok:
            wx.MessageBox(f"打印失败：{err}\n请检查打印机。",
                          "错误", wx.OK | wx.ICON_ERROR, self)
            self.current -= 1
            self.btn.Enable()
            self.btn.SetLabel("打  印")
            return
        if self.current >= self.total_boxes:
            self.EndModal(wx.ID_OK)
        else:
            self.lbl.SetLabel(self._label_text())
            self.btn.Enable()
            self.btn.SetLabel("打  印")
            self.Layout()


# ── 主窗口 ────────────────────────────────────────────────

class MainFrame(wx.Frame):
    def __init__(self):
        w = CFG.get("window_width",  540)
        h = CFG.get("window_height", 500)
        super().__init__(None, title=APP_TITLE, size=(w, h))
        self.SetMinSize((480, 460))
        self._refreshing = False

        panel = wx.Panel(self)
        self._build_ui(panel)
        self._fill_provinces()
        self.Centre()
        self.Show()

    def _build_ui(self, panel):
        root = wx.BoxSizer(wx.VERTICAL)

        # ── 标题 ──────────────────────────────────────────
        title = wx.StaticText(panel, label="华耕嘉成  发货贴纸打印")
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT,
                              wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        root.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 14)

        # ── 地址区 ────────────────────────────────────────
        addr_box = wx.StaticBoxSizer(
            wx.StaticBox(panel, label="收货地址"), wx.VERTICAL)

        # 国家
        r1 = wx.BoxSizer(wx.HORIZONTAL)
        r1.Add(wx.StaticText(panel, label="国家"), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self.cb_country = wx.ComboBox(panel, value="中国",
                                      choices=["中国", "其他"],
                                      style=wx.CB_DROPDOWN, size=(80, -1))
        r1.Add(self.cb_country, 0, wx.ALIGN_CENTER_VERTICAL)
        addr_box.Add(r1, 0, wx.ALL, 6)

        # 省 + 市
        r2 = wx.BoxSizer(wx.HORIZONTAL)
        r2.Add(wx.StaticText(panel, label="省"), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self.cb_prov = wx.ComboBox(panel, choices=[],
                                   style=wx.CB_DROPDOWN, size=(130, -1))
        r2.Add(self.cb_prov, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 16)
        r2.Add(wx.StaticText(panel, label="市"), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self.cb_city = wx.ComboBox(panel, choices=[],
                                   style=wx.CB_DROPDOWN, size=(130, -1))
        r2.Add(self.cb_city, 0, wx.ALIGN_CENTER_VERTICAL)
        addr_box.Add(r2, 0, wx.ALL, 6)

        # 区 + 详细地址
        r3 = wx.BoxSizer(wx.HORIZONTAL)
        r3.Add(wx.StaticText(panel, label="区"), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self.cb_dist = wx.ComboBox(panel, choices=[],
                                   style=wx.CB_DROPDOWN, size=(130, -1))
        r3.Add(self.cb_dist, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 16)
        r3.Add(wx.StaticText(panel, label="详细地址"), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self.tc_detail = wx.TextCtrl(panel, size=(160, -1))
        self.tc_detail.SetHint("街道/门牌（可空）")
        r3.Add(self.tc_detail, 1, wx.ALIGN_CENTER_VERTICAL)
        addr_box.Add(r3, 0, wx.ALL, 6)

        root.Add(addr_box, 0, wx.ALL | wx.EXPAND, 10)

        # ── 发货箱数 ──────────────────────────────────────
        qty_row = wx.BoxSizer(wx.HORIZONTAL)
        qty_row.Add(wx.StaticText(panel, label="发货箱数"), 0,
                    wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)
        self.tc_qty = wx.TextCtrl(panel, size=(100, -1))
        self.tc_qty.SetFont(wx.Font(13, wx.FONTFAMILY_DEFAULT,
                                    wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        qty_row.Add(self.tc_qty, 0, wx.ALIGN_CENTER_VERTICAL)
        root.Add(qty_row, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

        # ── 预览 ──────────────────────────────────────────
        preview_box = wx.StaticBoxSizer(
            wx.StaticBox(panel, label="打印预览"), wx.VERTICAL)
        self.lbl_preview = wx.StaticText(
            panel, label="收货地址：—\n箱号格式：—",
            style=wx.ST_NO_AUTORESIZE)
        self.lbl_preview.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT,
                                          wx.FONTSTYLE_NORMAL,
                                          wx.FONTWEIGHT_NORMAL))
        preview_box.Add(self.lbl_preview, 0, wx.ALL | wx.EXPAND, 8)
        root.Add(preview_box, 0, wx.ALL | wx.EXPAND, 10)

        # ── 开始打印按钮 ──────────────────────────────────
        self.btn_start = wx.Button(panel, label="开始打印", size=(-1, 44))
        self.btn_start.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT,
                                        wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        self.btn_start.SetBackgroundColour(wx.Colour(34, 139, 230))
        self.btn_start.SetForegroundColour(wx.WHITE)
        root.Add(self.btn_start, 0, wx.ALL | wx.EXPAND, 10)

        panel.SetSizer(root)

        # ── 事件绑定 ──────────────────────────────────────
        self.cb_prov.Bind(wx.EVT_COMBOBOX, self._on_prov_change)
        self.cb_prov.Bind(wx.EVT_TEXT,     self._on_prov_change)
        self.cb_city.Bind(wx.EVT_COMBOBOX, self._on_city_change)
        self.cb_city.Bind(wx.EVT_TEXT,     self._on_city_change)
        self.cb_dist.Bind(wx.EVT_COMBOBOX, self._update_preview)
        self.cb_dist.Bind(wx.EVT_TEXT,     self._update_preview)
        self.tc_qty .Bind(wx.EVT_TEXT,     self._update_preview)
        self.btn_start.Bind(wx.EVT_BUTTON, self._on_start)

    # ── 地区联动 ──────────────────────────────────────────

    def _fill_provinces(self):
        provs = region_data.get_provinces()
        self.cb_prov.Clear()
        for p in provs:
            self.cb_prov.Append(p)
        self.cb_prov.AutoComplete(provs)

    def _on_prov_change(self, _evt=None):
        prov = self.cb_prov.GetValue().strip()
        cities = region_data.get_cities(prov)
        self.cb_city.Clear()
        for c in cities:
            self.cb_city.Append(c)
        self.cb_city.AutoComplete(cities)
        self.cb_city.SetValue("")
        self.cb_dist.Clear()
        self.cb_dist.SetValue("")
        self._update_preview()

    def _on_city_change(self, _evt=None):
        city = self.cb_city.GetValue().strip()
        dists = region_data.get_districts(city)
        self.cb_dist.Clear()
        for d in dists:
            self.cb_dist.Append(d)
        self.cb_dist.AutoComplete(dists)
        self.cb_dist.SetValue("")
        self._update_preview()

    # ── 预览更新 ──────────────────────────────────────────

    def _get_location(self) -> str:
        prov = self.cb_prov.GetValue().strip()
        dist = self.cb_dist.GetValue().strip()
        if not prov:
            return ""
        return f"{prov}{dist}" if dist else prov

    def _get_boxes(self) -> int:
        try:
            n = int(self.tc_qty.GetValue().strip())
            return n if n > 0 else 0
        except ValueError:
            return 0

    def _update_preview(self, _evt=None):
        loc   = self._get_location()
        boxes = self._get_boxes()
        if loc and boxes > 0:
            self.lbl_preview.SetLabel(
                f"收货地址：{loc}\n"
                f"箱号格式：{loc}  {boxes}-1  /  {loc}  {boxes}-2  /  ..."
            )
        else:
            self.lbl_preview.SetLabel("收货地址：—\n箱号格式：—")

    # ── 开始打印 ──────────────────────────────────────────

    def _on_start(self, _evt):
        loc = self._get_location()
        if not loc:
            wx.MessageBox("请选择省份和区", "提示", wx.OK | wx.ICON_INFORMATION)
            return
        boxes = self._get_boxes()
        if boxes <= 0:
            wx.MessageBox("请输入有效的发货箱数", "提示", wx.OK | wx.ICON_INFORMATION)
            return

        dlg = PrintDialog(self, loc, boxes)
        dlg.ShowModal()
        dlg.Destroy()


# ── 入口 ──────────────────────────────────────────────────

def main():
    global CFG
    CFG = load_config()
    region_data.init_builtin()
    app = wx.App(False)
    MainFrame()
    app.MainLoop()


if __name__ == "__main__":
    main()
