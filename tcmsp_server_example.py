#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TCMSP 剪贴板管线 Agent — 伺服模式（Daemon）

核心流程（零打字，纯热键）：
    1. 你从R脚本复制药材名（如：当归）
    2. 在TCMSP确认条目存在 → 按 [记录热键]，Agent记录药材名并弹出通知
    3. 你复制网页中的JSON数组 → 按 [处理热键]，Agent自动：提取→转Excel→拼音命名→归档→记日志→弹出通知
    4. 自动回到步骤1，处理下一个

两种运行模式:
    普通模式（默认）: 控制台显示状态看板，事件驱动刷新（平时不占用CPU）
    静默模式        : 纯后台运行，控制台零刷新，全部交互通过系统通知完成

用法:
    python tcmsp_server.py           # 普通模式
    python tcmsp_server.py --silent  # 静默模式（推荐长期使用）

安装依赖:
    pip install pandas pyperclip pypinyin openpyxl keyboard plyer
"""

import argparse
import json
import os
import re
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

# ===================== 依赖检查 =====================
try:
    import pandas as pd
except ImportError:
    print("[错误] 缺少 pandas: pip install pandas")
    sys.exit(1)

try:
    import pyperclip
except ImportError:
    print("[错误] 缺少 pyperclip: pip install pyperclip")
    sys.exit(1)

try:
    from pypinyin import lazy_pinyin
except ImportError:
    print("[错误] 缺少 pypinyin: pip install pypinyin")
    sys.exit(1)

try:
    import keyboard
except ImportError:
    print("[错误] 缺少 keyboard: pip install keyboard")
    print("提示: Windows下若全局热键无效，请尝试以管理员身份运行终端")
    sys.exit(1)

# 通知模块（静默模式必需）
try:
    from plyer import notification
    NOTIFY_AVAILABLE = True
except ImportError:
    NOTIFY_AVAILABLE = False

# ===================== 用户配置 =====================
TARGET_DIR = Path(r"#")
LOG_FILE = TARGET_DIR / "processing_log.tsv"

# 别名映射: {"常用别名/异名": "标准名"}
HERB_ALIASES = {
    # "黄耆": "黄芪",
    # "广木香": "木香",
    # "川穹": "川芎",
}

# 热键配置 — 使用组合键避免与浏览器/系统冲突
# 支持的写法: 'ctrl+shift+1', 'alt+1', 'shift+f2', 'ctrl+alt+q' 等
HOTKEY_RECORD = 'ctrl+shift+1'     # 记录药材名
HOTKEY_PROCESS = 'ctrl+shift+2'    # 处理JSON并归档
HOTKEY_RESET = 'ctrl+shift+4'      # 重置状态
HOTKEY_QUIT = 'ctrl+shift+q'       # 退出程序
# ===================================================

# 音效反馈（Windows）
try:
    import winsound
    def beep_ok():
        winsound.MessageBeep(winsound.MB_OK)
    def beep_error():
        winsound.MessageBeep(winsound.MB_ICONHAND)
except ImportError:
    def beep_ok():
        print('\a', end='', flush=True)
    def beep_error():
        print('\a', end='', flush=True)


class TCMSPServer:
    STATE_IDLE = "idle"
    STATE_WAITING_JSON = "waiting_json"

    def __init__(self, silent=False):
        self.silent = silent
        self.state = self.STATE_IDLE
        self.lock = threading.Lock()
        self.running = True

        # 当前药材信息
        self.original_name = None
        self.standard_name = None
        self.pinyin_name = None

        # 日志历史（普通模式用）
        self.history = []

        # 统计
        self.success_count = 0
        self.error_count = 0

    # ---------- 工具方法 ----------

    def _extract_json_array(self, text: str) -> str:
        """从文本中提取JSON数组，使用栈匹配算法"""
        text = text.strip()
        if not text:
            return None

        start_idx = text.find('[')
        if start_idx == -1:
            return None

        depth = 0
        in_string = False
        escape_next = False
        end_idx = -1

        for i, ch in enumerate(text[start_idx:], start_idx):
            if escape_next:
                escape_next = False
                continue
            if ch == '\\':
                escape_next = True
                continue
            if ch == '"':
                in_string = not in_string
                continue
            if not in_string:
                if ch == '[':
                    depth += 1
                elif ch == ']':
                    depth -= 1
                    if depth == 0:
                        end_idx = i
                        break

        if end_idx != -1:
            candidate = text[start_idx:end_idx + 1]
            try:
                json.loads(candidate)
                return candidate
            except json.JSONDecodeError:
                pass

        # 回退：匹配 data: [...] 模式
        pattern = re.compile(r'data:\s*(\[.*?\])\s*(?:pageSize|pageNum|total)', re.DOTALL | re.IGNORECASE)
        match = pattern.search(text)
        if match:
            candidate = match.group(1)
            try:
                json.loads(candidate)
                return candidate
            except json.JSONDecodeError:
                pass

        return None

    def _clean_herb_name(self, text: str) -> str:
        """清洗从R控制台等来源复制的药材名"""
        text = text.strip()
        # 去除R控制台前缀: [1] "当归" -> 当归
        text = re.sub(r'^\s*\[\d+\]\s*["\']?', '', text)
        text = re.sub(r'["\']?\s*$', '', text)
        # 取第一行第一个空白分隔的词
        text = text.split('\n')[0]
        text = text.split('\t')[0]
        text = text.split()[0] if text.split() else text
        return text.strip()

    def _is_chinese_name(self, text: str) -> bool:
        """检查文本是否像药材名（含中文，不含JSON特征）"""
        if len(text) > 50:
            return False
        if '[' in text or '{' in text:
            return False
        return any('\u4e00' <= ch <= '\u9fff' for ch in text)

    def _to_pinyin(self, name: str) -> str:
        """中文转拼音文件名: 当归 -> Dang Gui"""
        pinyin_list = lazy_pinyin(name.strip())
        return " ".join(word.capitalize() for word in pinyin_list)

    def _parse_dataframe(self, json_str: str) -> pd.DataFrame:
        data = json.loads(json_str)
        if isinstance(data, list):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            return pd.DataFrame([data])
        raise ValueError(f"未预期的数据结构: {type(data)}")

    def _save_excel(self, df: pd.DataFrame) -> Path:
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        filename = f"{self.pinyin_name}.xlsx"
        output_path = TARGET_DIR / filename

        counter = 1
        base = self.pinyin_name
        while output_path.exists():
            filename = f"{base}_{counter}.xlsx"
            output_path = TARGET_DIR / filename
            counter += 1

        df.to_excel(output_path, index=False, engine='openpyxl')
        return output_path

    def _log(self, output_path: Path, row_count: int):
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        if not LOG_FILE.exists():
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                f.write("timestamp\toriginal_name\tstandard_name\tpinyin_name\tfilename\trow_count\n")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()}\t"
                    f"{self.original_name}\t{self.standard_name}\t"
                    f"{self.pinyin_name}\t{output_path.name}\t{row_count}\n")

    def _notify(self, title: str, message: str):
        """系统通知（不阻塞）+ 追加历史日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_line = f"[{timestamp}] {title}: {message}"
        self.history.append(log_line)
        if len(self.history) > 20:
            self.history = self.history[-20:]

        if NOTIFY_AVAILABLE:
            try:
                notification.notify(
                    title=title,
                    message=message,
                    timeout=3,
                    app_name="TCMSP Agent"
                )
            except Exception:
                pass

    def _reset_state(self):
        with self.lock:
            self.state = self.STATE_IDLE
            self.original_name = None
            self.standard_name = None
            self.pinyin_name = None

    # ---------- 热键回调 ----------

    def on_record(self):
        """热键: 读取剪贴板为药材名"""
        with self.lock:
            try:
                text = pyperclip.paste()
            except Exception as e:
                beep_error()
                self._notify("错误", f"读取剪贴板失败: {e}")
                if not self.silent:
                    self.render()
                return

            name = self._clean_herb_name(text)

            if not name:
                beep_error()
                self._notify("错误", "剪贴板为空，请先复制药材名")
                if not self.silent:
                    self.render()
                return

            if not self._is_chinese_name(name):
                beep_error()
                preview = name[:30].replace('\n', ' ')
                self._notify("错误", f"剪贴板不像药材名: {preview}")
                if not self.silent:
                    self.render()
                return

            self.original_name = name
            self.standard_name = HERB_ALIASES.get(name, name)
            self.pinyin_name = self._to_pinyin(self.standard_name)
            self.state = self.STATE_WAITING_JSON

            beep_ok()
            if self.original_name != self.standard_name:
                self._notify("药材已记录", f"{self.original_name}→{self.standard_name} ({self.pinyin_name})")
            else:
                self._notify("药材已记录", f"{self.standard_name} -> {self.pinyin_name}.xlsx")
            if not self.silent:
                self.render()

    def on_process(self):
        """热键: 读取剪贴板为JSON并处理归档"""
        with self.lock:
            if self.state != self.STATE_WAITING_JSON:
                beep_error()
                self._notify("提示", f"请先按 {HOTKEY_RECORD.upper()} 记录药材名")
                if not self.silent:
                    self.render()
                return

            try:
                text = pyperclip.paste()
            except Exception as e:
                beep_error()
                self._notify("错误", f"读取剪贴板失败: {e}")
                if not self.silent:
                    self.render()
                return

            json_str = self._extract_json_array(text)
            if json_str is None:
                beep_error()
                preview = text[:80].replace('\n', ' ')
                self._notify("错误", f"剪贴板中未找到有效JSON。前80字: {preview}")
                self.error_count += 1
                if not self.silent:
                    self.render()
                return

            try:
                df = self._parse_dataframe(json_str)
            except Exception as e:
                beep_error()
                self._notify("错误", f"JSON解析失败: {e}")
                self.error_count += 1
                if not self.silent:
                    self.render()
                return

            try:
                output_path = self._save_excel(df)
            except Exception as e:
                beep_error()
                self._notify("错误", f"保存Excel失败: {e}")
                self.error_count += 1
                if not self.silent:
                    self.render()
                return

            self._log(output_path, len(df))
            self.success_count += 1

            beep_ok()
            self._notify("归档完成", f"{self.standard_name} | {len(df)}条 | {output_path.name}")

            # 自动重置为空闲
            self.state = self.STATE_IDLE
            self.original_name = None
            self.standard_name = None
            self.pinyin_name = None
            if not self.silent:
                self.render()

    def on_reset(self):
        """热键: 重置状态"""
        with self.lock:
            if self.state == self.STATE_WAITING_JSON:
                self._notify("重置", f"已放弃: {self.standard_name}")
            self._reset_state()
            beep_ok()
            if not self.silent:
                self.render()

    def on_exit(self):
        """热键: 退出程序"""
        self.running = False
        self._notify("退出", "TCMSP Agent 已停止")

    # ---------- UI 渲染（事件驱动） ----------

    def render(self):
        """清屏并重绘当前状态看板。只在热键事件后调用，平时不消耗资源。"""
        if self.silent:
            return
        os.system('cls')
        print("=" * 64)
        print("  TCMSP 剪贴板管线 Agent [伺服模式]")
        print("=" * 64)

        if self.state == self.STATE_IDLE:
            print("  状态: 空闲  —  等待药材名")
            print(f"  操作: 复制药材名 → 按 [{HOTKEY_RECORD.upper()}] 记录")
            print("-" * 64)
        else:
            print("  状态: 等待 JSON  —  已锁定药材")
            line = f"  当前: {self.original_name}"
            if self.original_name != self.standard_name:
                line += f" → {self.standard_name}"
            line += f"  ({self.pinyin_name})"
            print(line)
            print(f"  操作: 复制JSON → 按 [{HOTKEY_PROCESS.upper()}] 自动处理归档")
            print("-" * 64)

        print(f"  统计: 成功 {self.success_count}  |  错误 {self.error_count}")
        print("-" * 64)

        if self.history:
            print("  最近日志:")
            for line in self.history[-8:]:
                print(f"    {line}")
        else:
            print("  等待操作...")

        print("-" * 64)
        print(f"  热键: [{HOTKEY_RECORD.upper()}]记录药材  [{HOTKEY_PROCESS.upper()}]处理JSON  [{HOTKEY_RESET.upper()}]重置  [{HOTKEY_QUIT.upper()}]退出")
        print("=" * 64)
        print()

    # ---------- 主入口 ----------

    def run(self):
        # 注册全局热键
        keyboard.add_hotkey(HOTKEY_RECORD, self.on_record)
        keyboard.add_hotkey(HOTKEY_PROCESS, self.on_process)
        keyboard.add_hotkey(HOTKEY_RESET, self.on_reset)
        keyboard.add_hotkey(HOTKEY_QUIT, self.on_exit)

        if self.silent:
            # 静默模式: 只做一次性输出，然后纯后台
            if not NOTIFY_AVAILABLE:
                print("[警告] 静默模式推荐安装 plyer 以获得系统通知")
                print("       pip install plyer")
            print(f"[Silent Mode] Agent 已后台启动")
            print(f"  热键: {HOTKEY_RECORD}=记录  {HOTKEY_PROCESS}=处理  {HOTKEY_RESET}=重置  {HOTKEY_QUIT}=退出")
            print(f"  所有反馈将通过系统通知/提示音呈现。\n")
        else:
            # 普通模式: 启动时渲染一次，之后事件驱动
            self.render()
            print("Agent 已启动，全局热键注册完成。此窗口可最小化到后台。\n")

        # 阻塞主线程等待退出（keyboard.wait 本身不消耗CPU）
        keyboard.wait(HOTKEY_QUIT)
        self.running = False
        time.sleep(0.3)
        print("\n已安全退出。")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="TCMSP 剪贴板管线 Agent")
    parser.add_argument(
        '-s', '--silent',
        action='store_true',
        help='静默模式: 控制台零刷新，纯后台运行，交互通过系统通知完成'
    )
    args = parser.parse_args()

    server = TCMSPServer(silent=args.silent)
    server.run()
