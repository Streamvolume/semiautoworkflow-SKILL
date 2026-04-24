# SemiAutoWorkflow Skill

## 目标

引导用户将其**重复性手动操作流程**（尤其是跨浏览器/跨应用的"数据搬运"类任务）转化为**人机协作的半自动化剪贴板管线**。核心哲学是：

> **人只承担不可替代的认知判断，Agent接管所有可确定的机械劳动。交互界面极简化为"复制 + 按键"。**

## 适用场景

当用户描述的任务具有以下特征时，启用本 Skill：

1. **跨应用孤岛**：数据在浏览器（网页数据库）、本地文件、R/Python 输出之间流转，无开放 API
2. **高重复、低批量**：任务需重复 N 次（N=10~1000），但每次的具体参数（搜索词、文件名）不同
3. **存在人工判断点**：如"搜索结果是否存在"、"选择哪个条目"、"命名是否符合规范"等无法全自动化的决策
4. **操作流程结构化**：虽然每次参数不同，但操作步骤序列相对固定

**典型场景**：
- 从网页数据库检索 → 下载/复制数据 → 格式转换 → 重命名 → 归档
- 从 R/Python 输出复制关键词 → 到网页查询 → 复制结果 → 回填到表格
- 任何"人在浏览器里点来点去，但步骤固定"的数据搬运任务

## 核心方法论：渐进式自动化（Progressive Automation）

不要追求一次性 100% 全自动。按以下四阶段推进：

| 阶段 | 人做什么 | Agent 做什么 | 投入产出比 |
|:---|:---|:---|:---:|
| **P0 手工** | 全部 | 旁观记录 | - |
| **P1 剪贴板** | 搜索、判断、复制 | 读取剪贴板 → 解析 → 转换 → 命名 → 归档 | **最佳起点** |
| **P2 热键** | 搜索、判断、复制 | P1 + 全局热键触发 + 系统通知反馈 | **推荐长期态** |
| **P3 半自动** | 仅做判断确认 | Agent 控制浏览器填充搜索、提取数据、人只确认 | 投入高 |

**本 Skill 聚焦 P1→P2 的快速落地。**

## Agent 引导流程（标准交互步骤）

当用户说"帮我搭建一个半自动工作流"或描述类似上述场景时，Agent 按以下步骤执行：

### Step 1: 问题诊断与操作元拆解

要求用户描述完整操作流程，然后将其拆解为**操作元（Atomic Operations）**。使用以下模板：

```
阶段 | 操作元 | 执行者 | 自动化难度 | 收益评级
```

**关键问题**：
1. 你的起点数据是什么？（R 输出？Excel 列表？手动输入？）
2. 数据在哪个网页/应用中处理？
3. 处理后的数据格式是什么？（JSON？表格？纯文本？）
4. 中间有没有格式转换步骤？（如 JSON 转 Excel、PDF 转 CSV）
5. 最终文件名有规律吗？（如"中文拼音分字大写"）
6. 最终归档到哪个目录？
7. 每一步中，哪些步骤**必须**由人判断？（如"搜索结果是否存在"）

### Step 2: ROI 分析与方案选择

根据拆解结果，将操作元分类：

- **红色高成本区**：立即自动化（通常是格式转换、重命名、文件移动、规则明确的计算）
- **黄色保留区**：暂时保留人工，但可辅助（如搜索结果判断、歧义选择）
- **绿色已自动**：无需处理

**向用户明确说明**：
- 哪些步骤完全不需要网页/第三方工具（如"JSON 转 Excel 网页"完全可用本地 pandas/jsonlite 替代，这是最大收益点之一）
- 哪些步骤需要人参与，以及如何最小化人的操作次数

**推荐方案架构**：
- **输入**：用户复制数据到剪贴板（或从固定文件读取）
- **触发**：热键或脚本调用
- **处理**：本地解析 → 转换 → 生成规范文件名
- **输出**：保存到目标目录 + 系统通知 + 日志记录

### Step 3: 最小可行产品（MVP）构建

与用户确认目标目录、命名规则、数据格式后，立即编写**可运行的 Python 脚本**。MVP 必须包含：

1. **剪贴板读取**（`pyperclip`）
2. **数据解析**（根据实际格式：JSON 栈匹配、正则提取、HTML 解析等）
3. **本地转换**（pandas 处理表格，不要用网页工具）
4. **规范命名**（根据用户规则实现函数，如拼音、日期、编号等）
5. **归档保存**（`pathlib` 处理路径，自动防覆盖）
6. **基础日志**（TSV 格式，记录时间、原始名、标准名、文件名、行数）

**MVP 可以是单次运行脚本**（用户复制 → 运行脚本 → 控制台输入确认），先验证流程正确。

### Step 4: 迭代优化为伺服模式

MVP 验证通过后，升级为**常驻后台的伺服程序**：

1. **全局热键**：用 `keyboard` 库注册组合键（推荐 `ctrl+shift+数字/字母`，避免与浏览器冲突）
   - 记录热键：读取剪贴板为"输入标识"（如药材名、关键词）
   - 处理热键：读取剪贴板为"数据载荷"（如 JSON、表格文本）
   - 重置热键：放弃当前状态
   - 退出热键：安全退出

2. **状态机**：
   - `IDLE` → 按记录热键 → `WAITING_DATA` → 按处理热键 → 处理完成 → `IDLE`
   - 状态转换伴随系统通知反馈

3. **系统通知**：用 `plyer` 弹出 Windows/macOS/Linux 原生通知，实现"零看控制台"交互

4. **音效反馈**：用 `winsound` / `print('\a')` 给成功/错误提供听觉反馈

5. **事件驱动 UI**：
   - **严禁**使用 `while True` + `time.sleep()` 循环刷新控制台
   - 只在热键触发时调用 `render()` 清屏重绘一次
   - 提供 `--silent` 参数：完全无控制台刷新，纯后台+通知

6. **别名映射**：提供字典配置解决异名同义、错别字问题

### Step 5: 固化与文档化

脚本稳定后：
1. 在脚本头部写好**依赖安装命令**和**热键说明注释**
2. 提供 `.bat` 或快捷方式方便一键启动
3. 如果需要长期后台运行，指导用户用 `pythonw`（Windows 无窗口）或系统服务方式启动
4. 将脚本和本 Skill 一起归档到项目 `scripts/` 或 `.kimi/` 目录

## 技术实现模板（可直接复用）

以下是一个通用剪贴板管线的骨架，Agent 应根据用户具体需求填充解析逻辑和命名规则。

```python
#!/usr/bin/env python3
import json, os, re, sys, threading, time, argparse
from datetime import datetime
from pathlib import Path
import pandas as pd
import pyperclip
from pypinyin import lazy_pinyin

try:
    import keyboard
except ImportError:
    sys.exit("pip install keyboard")

try:
    from plyer import notification
    NOTIFY = True
except ImportError:
    NOTIFY = False

# ============ 用户配置区 ============
TARGET_DIR = Path(r"D:\your_project\data")
LOG_FILE = TARGET_DIR / "log.tsv"
ALIASES = {}  # {"别名": "标准名"}

# 热键（避免浏览器冲突，用组合键）
HK_RECORD = 'ctrl+shift+1'   # 记录输入标识
HK_PROCESS = 'ctrl+shift+2'  # 处理数据载荷
HK_RESET = 'ctrl+shift+4'    # 重置
HK_QUIT = 'ctrl+shift+q'     # 退出
# ===================================

class ClipboardPipeline:
    STATE_IDLE = "idle"
    STATE_WAITING = "waiting"

    def __init__(self, silent=False):
        self.silent = silent
        self.state = self.STATE_IDLE
        self.lock = threading.Lock()
        self.running = True
        self.label = None       # 当前输入标识（如药材名）
        self.history = []
        self.ok = self.err = 0

    # --- 用户需自定义的核心逻辑 ---
    def extract_payload(self, text: str):
        """从剪贴板文本中提取结构化数据。返回解析后的对象或 None"""
        # 示例: 提取 JSON 数组
        start = text.find('[')
        end = text.rfind(']')
        if start != -1 and end > start:
            try:
                return json.loads(text[start:end+1])
            except:
                pass
        return None

    def to_dataframe(self, payload):
        """将提取的数据转为 DataFrame"""
        if isinstance(payload, list):
            return pd.DataFrame(payload)
        raise ValueError("Unsupported payload type")

    def generate_filename(self, label: str) -> str:
        """根据标识生成规范文件名（不含扩展名）"""
        # 示例: 中文转拼音
        name = ALIASES.get(label, label)
        pinyin = " ".join(w.capitalize() for w in lazy_pinyin(name))
        return pinyin

    # --- 通用基础设施（通常无需修改）---
    def save(self, df: pd.DataFrame, filename_base: str) -> Path:
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        path = TARGET_DIR / f"{filename_base}.xlsx"
        c = 1
        while path.exists():
            path = TARGET_DIR / f"{filename_base}_{c}.xlsx"
            c += 1
        df.to_excel(path, index=False, engine='openpyxl')
        return path

    def log(self, path: Path, rows: int):
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        if not LOG_FILE.exists():
            open(LOG_FILE, 'w', encoding='utf-8').write("time\tlabel\tfilename\trows\n")
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now().isoformat()}\t{self.label}\t{path.name}\t{rows}\n")

    def notify(self, title: str, msg: str):
        t = datetime.now().strftime("%H:%M:%S")
        self.history.append(f"[{t}] {title}: {msg}")
        self.history = self.history[-20:]
        if NOTIFY:
            try:
                notification.notify(title=title, message=msg, timeout=3, app_name="Pipeline")
            except:
                pass

    def render(self):
        if self.silent:
            return
        os.system('cls')
        print("=" * 60)
        print("  Clipboard Pipeline Agent")
        print("=" * 60)
        st = "空闲-等待标识" if self.state == self.STATE_IDLE else f"等待数据 [{self.label}]"
        print(f"  状态: {st}")
        print(f"  统计: 成功 {self.ok} | 错误 {self.err}")
        print("-" * 60)
        for line in self.history[-6:]:
            print(f"    {line}")
        print("-" * 60)
        print(f"  {HK_RECORD}=记录 {HK_PROCESS}=处理 {HK_RESET}=重置 {HK_QUIT}=退出")
        print("=" * 60)

    # --- 热键回调 ---
    def on_record(self):
        with self.lock:
            text = pyperclip.paste().strip()
            if not text:
                self.notify("错误", "剪贴板为空"); self.render(); return
            self.label = text
            self.state = self.STATE_WAITING
            self.notify("已记录", f"{self.label} -> {self.generate_filename(self.label)}")
            self.render()

    def on_process(self):
        with self.lock:
            if self.state != self.STATE_WAITING:
                self.notify("提示", f"先按 {HK_RECORD.upper()}"); self.render(); return
            text = pyperclip.paste()
            payload = self.extract_payload(text)
            if payload is None:
                self.notify("错误", "剪贴板中未识别到有效数据"); self.err += 1; self.render(); return
            try:
                df = self.to_dataframe(payload)
                fname = self.generate_filename(self.label)
                path = self.save(df, fname)
                self.log(path, len(df))
                self.ok += 1
                self.notify("完成", f"{self.label} | {len(df)}条 | {path.name}")
            except Exception as e:
                self.notify("错误", str(e)); self.err += 1
            self.label = None
            self.state = self.STATE_IDLE
            self.render()

    def on_reset(self):
        with self.lock:
            self.label = None
            self.state = self.STATE_IDLE
            self.notify("重置", "已回到空闲")
            self.render()

    def on_exit(self):
        self.running = False
        self.notify("退出", "Pipeline 已停止")

    def run(self):
        keyboard.add_hotkey(HK_RECORD, self.on_record)
        keyboard.add_hotkey(HK_PROCESS, self.on_process)
        keyboard.add_hotkey(HK_RESET, self.on_reset)
        keyboard.add_hotkey(HK_QUIT, self.on_exit)
        if self.silent:
            print(f"[Silent] {HK_RECORD}=记录 {HK_PROCESS}=处理 {HK_RESET}=重置 {HK_QUIT}=退出")
        else:
            self.render()
            print("Agent 已启动，可最小化到后台。\n")
        keyboard.wait(HK_QUIT)
        print("\n已退出。")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', '--silent', action='store_true', help='静默模式')
    args = parser.parse_args()
    ClipboardPipeline(silent=args.silent).run()
```

## 常见陷阱与解法

| 陷阱 | 现象 | 解法 |
|:---|:---|:---|
| **热键冲突** | 按热键触发浏览器/系统功能（如 F1 打开帮助） | 使用组合键：`ctrl+shift+数字/字母` |
| **循环刷屏** | 控制台不断闪烁刷新，CPU 占用高 | 删除 `while True` 循环，改为**事件驱动**（只在热键回调里调用 `render()`） |
| **伪 JSON 解析失败** | 网页复制的内容带 `data:`、`pageSize:` 等 JS 前缀 | 使用**栈匹配括号算法**找最外层 `[...]`，而非直接 `json.loads` |
| **剪贴板误判** | 把 JSON 误当药材名、或把药材名误当 JSON | 用启发式规则区分：长度>50/含`[`→数据；含中文字符→标识 |
| **静默模式无反馈** | 后台运行时不知道成功还是失败 | 必须安装 `plyer`，用系统通知 + 提示音做反馈 |
| **权限问题** | 全局热键无效 | Windows 下以**管理员身份**运行终端 |
| **R 控制台复制带前缀** | 复制"当归"实际得到 `[1] "当归"` | 用正则清洗：`re.sub(r'^\s*\[\d+\]\s*["\']?', '', text)` |

## 参考案例：TCMSP 药材成分提取管线

### 用户原始痛点
- R 脚本输出 100+ 味药材清单
- 每味药需手动：复制名称 → TCMSP 搜索 → 确认存在 → F12 找 JSON → 复制 → 打开 JSON 转 Excel 网页 → 下载 → 重命名为"拼音分字大写" → 剪切到项目文件夹
- 单人单线程，频繁跨应用切换

### 操作元拆解（23 个原子操作）
按"阶段-操作元-执行者-难度-收益"表格拆解，识别出：
- **最大收益点**：JSON 转 Excel 完全不需要网页（pandas 本地 3 行代码替代 10 次点击）
- **必须保留人工**：搜索结果是否存在、选择哪个条目
- **高价值自动化**：拼音命名、文件归档、日志记录

### 迭代过程
1. **MVP**：单次运行脚本，控制台输入药材名，读取剪贴板 JSON，自动后续流程
2. **伺服化**：改为常驻后台，F1/F2 热键触发（后发现 F1 浏览器冲突）
3. **事件驱动**：去掉 `while True` 循环刷新，改为热键事件后单次渲染
4. **组合键 + 静默模式**：热键改为 `ctrl+shift+1/2/4/q`，增加 `--silent` 纯后台模式

### 最终成品特性
- 状态机：`IDLE` → `WAITING_JSON` → `IDLE`
- 全局热键：`Ctrl+Shift+1` 记录药材，`Ctrl+Shift+2` 处理归档
- JSON 提取：栈匹配算法处理 `data: [...] pageSize: 15` 非标准格式
- 命名：`pypinyin` 实现"当归"→"Dang Gui"
- 反馈：`plyer` 系统通知 + `winsound` 提示音
- 双模式：普通模式（事件刷新看板）/ 静默模式（纯后台零刷新）
- 零阻塞开销：`keyboard.wait()` 休眠等待，无按键时 CPU 占用为 0

### 用户最终评价
> "已经达到实际应用要求"
