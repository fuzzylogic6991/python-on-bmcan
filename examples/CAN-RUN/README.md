# CAN-RUN 使用说明

## 文件夹结构

```
CAN-RUN/
├── dazhong2.0.py      # 主脚本（CAN 报文发送/接收/测试）
├── run.bat            # 一键运行（自动查找当前目录 .py 文件）
├── install_env.bat    # 一键安装环境
├── requirements.txt   # Python 依赖列表
├── BMAPI64.dll        # bmcan 接口驱动（运行时需要）
├── input_模板.xlsx    # 输入 Excel 模板
└── README.md          # 本文件
```

## 快速开始

### 1. 安装环境（仅首次）

双击 `install_env.bat`，或手动执行：

```bash
cd python-can-4.0.0
pip install -e .                        # 安装定制版 python-can（含 bmcan 接口）
pip install -r examples\CAN-RUN\requirements.txt
```

验证：`python -c "import can; print('bmcan' in can.VALID_INTERFACES)"` 输出 `True` 即成功。

### 2. 运行测试

双击 `run.bat` 即可。脚本会自动查找当前目录下的 `.py` 文件并运行。

即使主脚本改名为 `dazhong3.0.py` 等，`run.bat` 无需修改。

### 3. 配置输入文件

编辑 `dazhong2.0.py` 顶部的路径配置：

```python
EXCEL_PLAN_PATH = r"E:\...\input\大众22验证.xlsx"   # 测试用例 Excel
DLL_PATH = r"E:\...\VW_seed_to_key.dll"             # 安全算法 DLL
```

或参考 `input_模板.xlsx` 创建自己的测试用例。

## 文件说明

### run.bat
自动设置 `PYTHONPATH` 和工作目录，查找并运行当前目录下的 `.py` 文件。
- 文件名变化时无需修改
- 确保 `import can` 和 `BMAPI64.dll` 加载都能正常工作

### install_env.bat
一键完成：检查 Python → 安装定制版 python-can → 安装其余依赖 → 验证 bmcan 接口。

### requirements.txt

| 模块 | 用途 |
|------|------|
| `pandas` | Excel 配置文件读写 |
| `openpyxl` | Excel Workbook 操作 |
| `pyserial` | 串口通信（继电器控制） |
| `typing_extensions` | 类型支持 |

> **注意：** `python-can` 不在此列表中，必须通过 `pip install -e .` 从项目根目录安装定制版。

### BMAPI64.dll
`bmcan` 接口运行时依赖的驱动 DLL，需与本脚本放在同一目录。

## 常见问题

### `Unknown interface type "bmcan"`
装成了 PyPI 官方版。执行 `pip uninstall python-can`，然后重新 `pip install -e .`。

### `No module named 'pandas'`
执行 `pip install pandas openpyxl`。

### `Channel 0 is not connected or is in use`
CAN 硬件未连接或被其他程序占用。检查设备连接，关闭占用 CAN 通道的其他工具。

### `ModuleNotFoundError: No module named 'pkg_resources'`
Python 3.12+ 已移除 `pkg_resources`，本项目源码已改为 `importlib.metadata`，无需额外处理。
