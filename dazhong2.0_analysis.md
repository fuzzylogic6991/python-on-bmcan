# dazhong2.0.py 功能分析文档

## 概述

本项目是一个基于 `python-can` 库的 CAN/UDS 诊断自动化测试工具，主要面向大众车系 ECU 诊断测试场景。项目包含两个版本：
- **dazhong2.0.py**（V7，约561行）：轻量版，聚焦 CAN 报文发送/接收与数据库查询
- **can_send_receive_to_excelV6.1.py**（V9.4，约1891行）：完整版，增加串口继电器控制、安全访问、多服务展开等高级功能

---

## 一、CAN 报文发送与接收（python-can）

### 1.1 CAN 总线初始化

使用 `python-can` 库的 BMAPI 后端（BUSMASTER API），支持 Classic CAN 和 CAN FD 两种模式。

```python
bus = can.interface.Bus(
    bustype='bmcan',
    channel=0,
    bitrate=500000,
    is_fd=True,
    data_bitrate=2000000,
    tres=True
)
```

关键配置项：
- **bustype**: `bmcan` — BUSMASTER CAN 卡接口
- **channel**: 通道号（0/1）
- **bitrate**: 仲裁域波特率（500kbps）
- **data_bitrate**: CAN FD 数据域波特率（2Mbps）
- **can_mode**: 0=经典CAN, 1=CAN FD(无BRS), 2=CAN FD(有BRS), 3=CAN FD(BRS,短帧不填充)
- **is_extended_id**: 是否扩展帧
- **tres**: 自动发送流控帧（V9.4 新增）

### 1.2 报文发送

单帧发送：
```python
msg = can.Message(
    arbitration_id=arb_id,
    data=msg_data,
    is_extended_id=False,
    is_fd=True
)
bus.send(msg)
```

多帧发送（ISO-TP 流控）：
1. 发送首帧（FF，PCI=0x1X）
2. 等待接收方发送流控帧（FC，PCI=0x30）
3. 收到 FC 后，依次发送连续帧（CF，PCI=0x2X）

### 1.3 报文接收

```python
msg = bus.recv(timeout=0.1)
if msg and msg.arbitration_id in allowed_ids:
    logger.log_received_message(msg, time.time())
```

接收过程支持：
- 多帧接收：检测首帧后自动发送流控帧
- 超时重刷新：收到响应帧后重置计时器
- TesterPresent 过滤（V9.4）：自动跳过 `02 7E 00` 响应帧

### 1.4 ISO-TP 协议处理

**数据编码（parse_hex_input）**：
- 单帧（≤7字节）：`[length, data...]` → 填充
- 扩展单帧（8-62字节，CAN FD）：`[0x00, length, data...]` → 填充
- 多帧（>最大单帧）：首帧 `[0x1X, len_low, data...]` + 连续帧 `[0x2X, data...]`

**数据解码（extract_isotp_payload）**：
- 单帧：根据 PCI 低4位提取长度
- 扩展单帧：PCI=0x00 时从 data[1] 读取长度
- 多帧：根据首帧的总长度字段拼接连续帧数据

**CAN FD DLC 映射表**：
```
DLC 0→0, 1→1, 2→2, 3→3, 4→4, 5→5, 6→6, 7→7, 8→8,
9→12, 10→16, 11→20, 12→24, 13→32, 14→48, 15→64
```

### 1.5 UDS 诊断服务（V9.4）

| 服务  | Handler              | 说明                         |
|-------|----------------------|------------------------------|
| 10    | Service10Handler     | 会话控制（不展开）           |
| 22    | Service22Handler     | 读取DID，支持展开步骤配置    |
| 27    | Service27Handler     | VW安全访问（seed→key）       |
| 27    | ServiceBA27Handler   | BA安全访问（XOR算法）        |
| 28    | Service28Handler     | 通信控制，支持监控目标ID     |
| 31    | Service31Handler     | 例程控制（不展开）           |
| 3E    | TesterPresentHandler | 保持会话活跃（TP/STP）       |

---

## 二、串口操作（V9.4）

### 2.1 串口初始化

```python
ser = serial.Serial(
    port="COM99", baudrate=9600,
    bytesize=serial.EIGHTBITS,
    parity=serial.PARITY_NONE,
    stopbits=serial.STOPBITS_ONE,
    timeout=0.5
)
ser.flushInput()
ser.flushOutput()
```

### 2.2 继电器控制（程控电源）

通过串口发送 HEX 命令控制 KL15/KL30 上电下电：

```python
RELAY_COMMANDS = {
    'KL15 on':  'A0 02 00 A2',
    'KL15 off': 'A0 02 01 A3',
    'KL30 on':  'A0 01 00 A1',
    'KL30 off': 'A0 01 01 A2',
}
```

命令编码后通过 `ser.write()` 发送，使用 `delay=0.5s` 等待继电器响应。

### 2.3 串口关闭

```python
def close_serial(ser):
    if ser and ser.is_open:
        ser.close()
```

在 `finally` 块中确保关闭，防止资源泄漏。

---

## 三、日志记录

### 3.1 CanMessageLogger 类

核心日志类，功能包括：

| 方法                            | 功能                                   |
|---------------------------------|----------------------------------------|
| `log_sent_message()`            | 记录发送的 CAN 报文                     |
| `log_received_message()`        | 记录接收的 CAN 报文                     |
| `finalize_and_analyze_response()` | 分析响应（正响应/负响应/22服务数据）     |
| `apply_comparison()`            | 结果比对（DB比对/HEX精确匹配/前缀匹配）  |
| `save_to_excel()`               | 保存到 Excel（不分段）                  |
| `save_to_excel_segmented()`     | 分段保存到 Excel（V9.4）               |
| `log_service28_monitor()`       | 28服务监控日志（V9.4）                  |

### 3.2 日志字段

每一条日志记录包含以下字段：

| 字段                 | 说明                           |
|----------------------|--------------------------------|
| 时间                 | 精确到毫秒的时间戳             |
| 类型                 | 发送 / 接收                    |
| CAN ID (十六进制)     | 如 `0x711`                     |
| 扩展帧               | 是 / 否                        |
| 数据长度             | DLC                            |
| 数据 (十六进制)       | 如 `02 10 03 CC CC CC CC CC`   |
| 数据 (ASCII)         | 可打印字符显示，否则 `.`       |
| 测试用例ID           | 从 Excel 读取的用例ID          |
| 测试项               | 测试名称                       |
| 结果                 | 通过 / 失败 / 否定响应 / 超时  |
| 肯定响应值           | 如 `62 F1 90`                  |
| 否定响应值           | 如 `7F 22 31`                  |
| 22服务内容(hex)      | 22服务返回的数据（HEX格式）    |
| 22服务内容(ascii)    | 22服务返回的数据（ASCII格式）  |
| 数据库获取值         | adb 查询到的数据库值           |
| 期望来源             | 比对依据（DB id=X / 期望HEX）  |
| 轮次                 | 当前循环轮次（V9.4）           |

### 3.3 分段写入（V9.4）

- Excel 单段最大行数：4000 行（可配置）
- Log 单文件最大大小：10MB（可配置）
- 自动保存间隔：每 500 条测试用例
- 最终合并所有分段，删除临时文件

### 3.4 通道1监听日志（V9.4）

独立的监听线程，记录 channel 1 上的全部 CAN 报文到分段日志文件：
- 格式：`时间  RX  CAN_ID  [FD/CL]  DLC=XX  HEX_DATA`
- 达到大小上限自动切新文件
- `flush()` 立即刷新，确保数据不丢失

---

## 四、ADB Shell 操作

### 4.1 基本 ADB 命令执行

```python
def execute_adb_command(command: str) -> Optional[str]:
    subprocess.run(['adb', 'shell', 'sync'], capture_output=True, text=True, check=True)
    result = subprocess.run(['adb', 'shell', command], capture_output=True, text=True, check=True)
    return result.stdout.strip()
```

先执行 `adb shell sync` 确保文件系统同步，再执行目标命令。

### 4.2 查询终端数据库（SQLite）

```python
def query_tbox_data_by_id(param_id: int) -> Optional[str]:
    cmd = f"cd /oemdata/parameters/ && sqlite3 paramsDb 'select * from paramstbl where id={param_id};'"
    result = execute_adb_command(cmd)
    ...
```

**查询流程**：
1. 通过 `adb shell` 进入终端
2. 切换到参数数据库目录 `/oemdata/parameters/`
3. 使用 `sqlite3` 执行 SQL 查询 `paramstbl` 表
4. 按 `param_id` 检索参数值
5. 解析 `|` 分隔的返回结果

**用于**：将 CAN 诊断读取的数据与终端数据库中的实际值进行比对验证。

### 4.3 数据库比对逻辑

```python
ascii_val = result.get('22服务内容(ascii)', '')
db_val = query_tbox_data_by_id(cfg['expected_db_id']) or ''
passed = (ascii_val == db_val)
```

---

## 五、配置管理

### 5.1 配置源

| 配置类型         | 来源                       | 说明                         |
|------------------|----------------------------|------------------------------|
| CAN 通信配置     | DEFAULT_CAN_CONFIG 字典    | 波特率/CAN模式/ID等          |
| 执行控制配置     | LOOP_COUNT / LOOP_GAP      | 循环次数与间隔               |
| 文件路径配置     | DLL/Excel/输出/监听目录      | 各资源路径                   |
| 继电器配置       | SERIAL_PORT / RELAY_COMMANDS | 串口号与命令                |
| 22服务展开规则   | SERVICE22_EXPANSION_STEPS   | 步骤列表（可注释/取消注释）  |
| 27服务安全访问   | SERVICE27_CONFIG           | VW seed→key 流程参数         |
| BA安全访问       | SERVICE_BA27_CONFIG        | BA seed→key 流程参数         |
| 分段写入配置     | FILE_SEGMENT_CONFIG        | 文件大小与行数限制           |
| 测试用例配置     | Excel 文件                  | 每行一个测试用例             |

### 5.2 Excel 输入格式

Excel 列字段：
- 是否启用 / 测试用例ID / 测试项 / CANID / 请求数据 / 期望DBID / 期望HEX / 响应超时时间 / 等待间隔时间 / 是否周期发送 / 周期时间(秒) / 周期CAN模式 / 监控CANID / 监控时长(秒) / 期望报文数

### 5.3 特殊触发值（V9.4）

| 请求数据值 | 含义                       |
|-----------|----------------------------|
| `KEY`     | VW安全访问（2701/2702）    |
| `KEY1`    | VW安全访问（2705/2706）    |
| `KEYBB`   | BA安全访问                 |
| `TP3E`    | 启动 Tester Present        |
| `STP3E`   | 停止 Tester Present        |
| `KL30ON/OFF` | 继电器控制               |
| `KL15ON/OFF` | 继电器控制               |

---

## 六、安全访问算法（V9.4）

### 6.1 VW Seed to Key（DLL 调用）

```python
dll = ctypes.CDLL("VW_seed_to_key.dll")
dll.VW_Seed2Key.argtypes = [ctypes.POINTER(ctypes.c_uint8), ctypes.POINTER(ctypes.c_uint8)]
dll.VW_Seed2Key.restype = ctypes.c_int16
```

通过 `ctypes` 调用外部 DLL，传入4字节种子，获取4字节密钥。

### 6.2 BA Seed to Key（纯 Python）

```python
xor_bytes = [0x26, 0xBF, 0x6D, 0x96]
data = [(seed_bytes[i] ^ xor_bytes[i]) & 0xFF for i in range(4)]
# ... 位操作重组得到4字节密钥
```

基于异或和位操作的自定义算法。

---

## 七、执行流程

### 7.1 主流程（V7 简化版）

```
main()
  ├── load_send_configs_from_excel()  → 加载 Excel 测试用例
  ├── 循环 N 次（命令行参数控制）
  │     └── send_and_receive_can_messages()
  │           ├── can.interface.Bus() → 打开 CAN 总线
  │           ├── 遍历每条测试用例：
  │           │     ├── 发送 Keep-Alive（如需要）
  │           │     ├── parse_hex_input() → 解析请求数据（ISO-TP）
  │           │     ├── bus.send() → 发送单帧或多帧
  │           │     ├── bus.recv() → 接收响应
  │           │     ├── finalize_and_analyze_response() → 分析响应
  │           │     └── apply_comparison() → 比对结果
  │           └── bus.shutdown() → 关闭 CAN 总线
  └── logger.save_to_excel() → 保存结果到 Excel
```

### 7.2 主流程（V9.4 完整版）

```
main()
  ├── register_signal_handlers() → 注册紧急保存
  ├── load_and_expand_configs()  → 加载+展开配置
  ├── CanMessageLogger() → 创建日志器
  ├── 循环 N 轮：
  │     └── send_and_receive_can_messages()
  │           ├── init_serial() → 打开串口
  │           ├── can.interface.Bus() → 打开 CAN
  │           ├── 启动 channel1_listener 线程
  │           ├── 创建 TesterPresentHandler 单例
  │           ├── 遍历每条测试用例：
  │           │     ├── get_handler() → 选择 Handler
  │           │     ├── handler.execute_test() → 执行
  │           │     ├── tp3e_handler.check_and_send() → 检查3E
  │           │     └── 分片等待 + 检查3E
  │           ├── 停止周期任务
  │           ├── 停止监听线程
  │           ├── bus.shutdown()
  │           └── close_serial()
  └── logger.save_to_excel_segmented() → 分段保存
```

### 7.3 异常保护（V9.4）

```python
signal.signal(signal.SIGINT, emergency_save)   # Ctrl+C
signal.signal(signal.SIGTERM, emergency_save)   # kill 信号
atexit.register(emergency_save)                 # 正常退出
```

异常退出前自动保存已收集的日志数据。

---

## 八、依赖项

| 库          | 用途                          |
|-------------|-------------------------------|
| `python-can` | CAN 总线通信（BMAPI 后端）    |
| `pandas`    | Excel 读写、数据处理           |
| `openpyxl`  | Excel 引擎                     |
| `pyserial`  | 串口通信（继电器控制）         |
| `ctypes`    | 调用外部 DLL（VW 安全算法）   |
| `threading` | 多线程（周期发送、监听、3E）  |
| `signal`    | 信号处理（紧急保存）           |
| `subprocess`| 执行 ADB 命令                  |

---

## 九、线程安全（V9.4）

由于 BMAPI 后端非线程安全，使用全局锁保护 `bus.send()` 调用：

```python
BUS_SEND_LOCK = threading.Lock()

with BUS_SEND_LOCK:
    bus.send(msg)
```

---

## 十、扩展性设计（V9.4）

所有 UDS 服务处理器继承自 `UDSServiceHandler` 基类：

```python
class UDSServiceHandler:
    def get_expanded_configs(cfg) → 展开配置
    def send(bus, arb_id, cfg, logger) → 发送
    def receive(bus, cfg, logger) → 接收
    def analyze_and_judge(cfg, logger) → 分析判定
    def execute_test(bus, arb_id, cfg, logger) → 完整流程
```

新增服务只需：
1. 继承 `UDSServiceHandler`
2. 重写需自定义的方法
3. 在 `get_handler()` 中添加路由规则
