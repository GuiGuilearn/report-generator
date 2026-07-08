"""代码沙箱模块：AST 安全扫描 + 子进程隔离执行 + 自动修复调度。

职责：
- 第一层：AST 静态扫描，禁止危险模块导入
- 第二层：子进程隔离执行（subprocess）
- 第三层：超时控制（默认 30s）
- 自动修复循环：执行失败后调用 LLM 修复，最多 3 次
- input() 替换器：每个 Cell 使用独立变量名
- 图形输出捕获：matplotlib preamble 拦截 plt.show()
- 外部依赖自动安装
"""

import ast
import logging
import os
import subprocess
import sys
import time
from typing import Optional

logger = logging.getLogger(__name__)

# ─── 禁止导入的危险模块 ───
FORBIDDEN_MODULES = {
    "os", "subprocess", "shutil", "socket", "pickle",
    "ctypes", "importlib", "multiprocessing", "signal", "pty",
    "pdb", "code", "codeop", "compileall",
}

# ─── Python 标准库模块集合（用于检测第三方依赖） ───
STD_LIB_MODULES = {
    "os", "sys", "math", "random", "json", "re", "datetime", "collections",
    "itertools", "functools", "typing", "io", "pathlib", "csv", "sqlite3",
    "xml", "html", "http", "urllib", "socket", "hashlib", "base64", "struct",
    "array", "copy", "enum", "logging", "unittest", "argparse", "time",
    "threading", "multiprocessing", "subprocess", "pickle", "shelve", "dbm",
    "zlib", "gzip", "bz2", "lzma", "tarfile", "zipfile", "fnmatch", "glob",
    "string", "textwrap", "difflib", "pprint", "statistics", "decimal",
    "fractions", "numbers", "operator", "inspect", "ast", "dis", "traceback",
    "warnings", "weakref", "gc", "atexit", "signal", "mmap", "errno", "ctypes",
    "platform", "sysconfig", "locale", "gettext", "abc", "shutil", "tempfile",
    "tkinter", "turtle", "webbrowser", "uuid", "secrets", "dataclasses",
    "contextlib", "asyncio", "concurrent", "venv", "ensurepip", "pip",
    "pkgutil", "importlib", "runpy", "zipapp", "email", "smtplib", "imaplib",
    "poplib", "ssl", "crypt", "hashlib", "hmac", "typing",
}

# ─── 图形输出捕获 preamble ───
GRAPHICS_PREAMBLE = """
# ====== 图形输出捕获 preamble ======
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

_original_show = plt.show
def _capture_show(*args, **kwargs):
    plt.savefig('_output_figure.png', dpi=150, bbox_inches='tight')
    plt.close()
    print("[图表已生成]")
plt.show = _capture_show
# ====== preamble 结束 ======
"""


def scan_imports(code: str) -> list[str]:
    """返回代码中导入的危险模块列表。

    解析代码中的 import 和 import from 语句，检测是否导入了
    FORBIDDEN_MODULES 中的危险模块。

    Args:
        code: Python 源代码字符串

    Returns:
        检测到的危险模块名称列表
    """
    try:
        tree = ast.parse(code)
        imports = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    imports.add(alias.name.split('.')[0])
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    imports.add(node.module.split('.')[0])
        return [m for m in imports if m in FORBIDDEN_MODULES]
    except SyntaxError:
        return []  # 语法错误由执行阶段处理


def extract_third_party_imports(code: str) -> list[str]:
    """扫描代码中的 import 语句，返回标准库之外的第三方包名。

    Args:
        code: Python 源代码字符串

    Returns:
        需要安装的第三方包名列表（已排序）
    """
    try:
        tree = ast.parse(code)
        imports = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    imports.add(alias.name.split('.')[0])
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    imports.add(node.module.split('.')[0])
        return sorted(imports - STD_LIB_MODULES)
    except SyntaxError:
        return []


def build_input_replacer(test_inputs: list[str], task_index: int) -> str:
    """为指定任务构建 input() 替换器代码。

    每个 Cell 使用独立变量名（_test_inputs_{idx}、_idx_{idx}），
    避免跨 Cell 串扰。同时替换 sys.stdin.readline 和 sys.stdin.read。

    Args:
        test_inputs: 模拟输入值列表
        task_index: 任务索引（用于生成独立变量名）

    Returns:
        input 替换器代码字符串
    """
    if not test_inputs:
        return ""

    inputs_repr = repr(test_inputs)

    replacer = f"""
# ====== 自动生成的 input 替换器（Cell {task_index} 专用）======
_test_inputs_{task_index} = {inputs_repr}
_idx_{task_index} = 0

def _mock_input_{task_index}(prompt=""):
    global _idx_{task_index}
    print(prompt, end="")
    if _idx_{task_index} < len(_test_inputs_{task_index}):
        val = _test_inputs_{task_index}[_idx_{task_index}]
        _idx_{task_index} += 1
        print(val)
        return val
    return ""

import sys
__builtins__.input = _mock_input_{task_index}
sys.stdin.readline = lambda: _mock_input_{task_index}("")
sys.stdin.read = lambda: "\\n".join(_test_inputs_{task_index})
# ====== 替换器结束，以下是原始代码 ======
"""
    return replacer


def build_execution_code(code: str, test_inputs: list[str], task_index: int) -> str:
    """构建完整的执行代码（含 preamble 和替换器）。

    组装顺序：
    1. 图形输出捕获 preamble（如果代码中使用了 matplotlib）
    2. input 替换器（如果有 test_inputs）
    3. 原始代码

    Args:
        code: 原始 Python 代码
        test_inputs: 模拟输入值列表
        task_index: 任务索引

    Returns:
        完整的可执行代码字符串
    """
    parts = []

    # 检测是否需要图形输出捕获
    if "matplotlib" in code or "plt." in code:
        parts.append(GRAPHICS_PREAMBLE)

    # 添加 input 替换器
    if test_inputs:
        parts.append(build_input_replacer(test_inputs, task_index))

    parts.append(code)
    return "\n".join(parts)


def install_third_party_packages(code: str, task_index: int) -> None:
    """自动安装代码中依赖的第三方包。

    Args:
        code: Python 源代码字符串
        task_index: 任务索引（用于日志）
    """
    packages = extract_third_party_imports(code)
    if not packages:
        return

    for pkg in packages:
        logger.info(f"[Task {task_index}] 安装第三方依赖: {pkg}")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", pkg, "--break-system-packages", "-q"],
                capture_output=True, text=True,
                timeout=120,
            )
        except Exception as e:
            logger.warning(f"[Task {task_index}] 安装 {pkg} 失败: {e}")


def execute_in_sandbox(
    code: str,
    timeout: int = 30,
    cwd: str = "temp/",
) -> tuple[bool, str, str]:
    """在子进程中执行代码，返回 (success, stdout, stderr)。

    使用 subprocess 启动独立 Python 进程执行代码。
    子进程崩溃不影响主流程。

    Args:
        code: 要执行的 Python 代码
        timeout: 执行超时时间（秒）
        cwd: 子进程工作目录

    Returns:
        (success, stdout, stderr) 三元组
    """
    try:
        os.makedirs(cwd, exist_ok=True)
        result = subprocess.run(
            [sys.executable, "-c", code],
            capture_output=True, text=True,
            timeout=timeout,
            cwd=cwd,
        )
        return (
            result.returncode == 0,
            result.stdout,
            result.stderr,
        )
    except subprocess.TimeoutExpired:
        return False, "", f"执行超时（>{timeout}s）"
    except Exception as e:
        return False, "", str(e)


def execute_code_task(
    code: str,
    test_inputs: list[str],
    task_index: int,
    timeout: int = 30,
    auto_fix_fn: Optional[callable] = None,
    auto_fix_max_attempts: int = 3,
) -> tuple[bool, str, str | None, str]:
    """执行单个代码任务，含完整的安全检查 + 自动修复流程。

    流程：
    1. AST 扫描，检测危险模块导入
    2. 安装第三方依赖
    3. 子进程隔离执行
    4. 若失败，自动修复循环（最多 auto_fix_max_attempts 次）

    Args:
        code: 原始 Python 代码
        test_inputs: 模拟输入值列表
        task_index: 任务索引
        timeout: 执行超时时间（秒）
        auto_fix_fn: 自动修复回调函数，签名为 fn(code, error_message) -> str | None
        auto_fix_max_attempts: 自动修复最大次数

    Returns:
        (success, stdout, stderr_or_error, final_code) 四元组
        - success: 是否执行成功
        - stdout: 标准输出文本
        - stderr_or_error: 失败时的错误信息（成功时为 None）
        - final_code: 最终执行的代码（可能经过修复）
    """
    # 第一层：AST 扫描
    dangerous = scan_imports(code)
    if dangerous:
        msg = f"检测到危险模块导入：{', '.join(dangerous)}"
        logger.warning(f"[Task {task_index}] {msg}")
        return False, "", msg, code

    # 安装第三方依赖
    install_third_party_packages(code, task_index)

    # 构建完整执行代码
    exec_code = build_execution_code(code, test_inputs, task_index)

    # 第二层 + 第三层：子进程隔离执行 + 超时控制
    success, stdout, stderr = execute_in_sandbox(exec_code, timeout=timeout)

    if success:
        return True, stdout, None, code

    # 自动修复循环
    current_code = code
    for attempt in range(auto_fix_max_attempts):
        logger.info(
            f"[Task {task_index}] 执行失败，尝试自动修复 (第 {attempt + 1}/{auto_fix_max_attempts} 次)"
        )

        if auto_fix_fn is None:
            break

        error_msg = stderr or stdout or "未知错误"
        try:
            fixed_code = auto_fix_fn(current_code, error_msg)
        except Exception as e:
            logger.warning(f"[Task {task_index}] 自动修复调用失败: {e}")
            break

        if fixed_code is None or fixed_code == current_code:
            logger.warning(f"[Task {task_index}] 修复未产生变化，停止尝试")
            break

        current_code = fixed_code
        fixed_exec_code = build_execution_code(current_code, test_inputs, task_index)
        success, stdout, stderr = execute_in_sandbox(fixed_exec_code, timeout=timeout)

        if success:
            logger.info(f"[Task {task_index}] 自动修复成功")
            return True, stdout, None, current_code

    return False, stdout, stderr, current_code
