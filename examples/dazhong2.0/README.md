# 环境安装说明

## 重要：python-can 必须使用本项目自带的定制版本

本项目对 `python-can` 进行了定制，新增了 `bmcan` 接口（位于 `can/interfaces/bmcan/`）。
**不能使用 PyPI 上的官方 python-can**，否则会报 `Unknown interface type "bmcan"`。

## 安装步骤

### 1. 进入项目根目录

```bash
cd python-can-4.0.0
```

### 2. 以开发模式安装本项目（包含 bmcan 接口）

```bash
pip install -e .
```

这一步会自动安装 python-can 的依赖（wrapt、typing_extensions、windows-curses、pywin32 等），
并注册本项目目录中的 `can` 包，使 `bmcan` 接口可用。

### 3. 安装其余依赖

```bash
cd examples
pip install -r requirements.txt
```

requirements.txt 包含的依赖（不包含 python-can，因为它必须通过步骤2安装）：

| 模块 | 用途 |
|------|------|
| `pandas` | Excel 配置文件读写 |
| `openpyxl` | Excel Workbook 操作 |
| `pyserial` | 串口通信（继电器控制） |
| `typing_extensions` | 类型支持 |

### 4. 验证安装

```bash
python -c "import can; print('bmcan' in can.VALID_INTERFACES)"
```

输出 `True` 即表示 `bmcan` 接口可用。

## 常见问题

### Q: 提示 `Unknown interface type "bmcan"`？
A: 说明装成了 PyPI 上的官方版本。请先 `pip uninstall python-can`，再回到步骤2安装本项目。

### Q: 提示 `ModuleNotFoundError: No module named 'pkg_resources'`？
A: Python 3.12+ 已移除 `pkg_resources`。本项目源码已改为使用 `importlib.metadata`，无需额外处理。

### Q: 提示 `No module named 'pandas'`？
A: 执行 `pip install pandas openpyxl`。
