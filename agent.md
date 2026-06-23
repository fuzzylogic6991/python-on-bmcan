# agent.md — AI Agent 工作指南

## 项目身份

这是一个基于 `python-can` 库的 CAN/UDS 诊断自动化测试工具，用于大众车系 TBOX/ECU 的诊断测试验证。

**主工作文件**：
- `examples/dazhong2.0.py`（V7，561行）— 轻量版
- `examples/can_send_receive_to_excelV6.1.py`（V9.4，1891行）— 完整版，**当前主力文件**

**硬件环境**：
- CAN 卡：BUSMASTER（BMAPI 后端）
- 操作系统：Windows
- 被测终端：Android TBOX（通过 ADB 连接）
- 程控电源：通过串口继电器控制（KL15/KL30）

---

## 代码架构

项目采用分层设计（V9.4）：

```
第1层：基础层（工具函数）
├── FD_DLC_MAP / get_fd_dlc()        — CAN FD DLC 映射
├── pad_data()                        — 数据填充
├── parse_hex_input()                 — ISO-TP 编码
├── extract_isotp_payload()           — ISO-TP 解码
├── VW_Seed2Key()                     — VW DLL 安全算法
├── calculate_key_level1()            — BA XOR 安全算法
├── execute_adb_command()             — ADB 命令执行
├── query_tbox_data_by_id()           — 数据库查询
├── init_serial() / execute_relay_command() / close_serial() — 串口继电器
├── periodic_send()                   — 周期发送线程
└── channel1_listener_segmented()     — 通道1监听线程

第2层：类层（服务处理器）
├── UDSServiceHandler                 — 基类（send/receive/analyze/execute_test）
├── Service10Handler                  — 会话控制
├── Service22Handler                  — 22服务（可展开多步）
├── Service27Handler                  — VW 安全访问
├── ServiceBA27Handler                — BA 安全访问
├── Service28Handler                  — 通信控制+监控
├── Service31Handler                  — 例程控制
├── TesterPresentHandler              — 3E00 维持会话
├── RelayHandler                      — 继电器
└── PeriodicHandler                   — 周期发送

第3层：调度层
├── get_handler()                     — 根据配置路由到对应 Handler
├── load_send_configs_from_excel()    — 加载 Excel 配置
└── load_and_expand_configs()         — 加载+展开（22服务）

第4层：异常保护
├── emergency_save()                  — 信号触发的紧急保存
└── register_signal_handlers()        — 注册 SIGINT/SIGTERM/atexit

第5层：执行层
├── send_and_receive_can_messages()   — 核心执行循环
└── main()                            — 入口（多轮循环）
```

---

## 配置系统

### 修改规则

**★★★ 新项目只修改文件顶部【项目配置区】★★★**

`DEFAULT_CAN_CONFIG` 字典约第46-68行，包含所有 CAN 通信参数。

### 关键配置项

| 配置项              | 位置      | 说明                                       |
|---------------------|-----------|--------------------------------------------|
| `bustype`           | can_config | CAN 卡类型，当前固定 `bmcan`                |
| `channel`           | can_config | 通道号（0/1）                               |
| `bitrate`           | can_config | 仲裁域波特率                                |
| `data_bitrate`      | can_config | CAN FD 数据域波特率                         |
| `can_mode`          | can_config | 0=经典CAN, 1=FD无BRS, 2=FD有BRS, 3=FD短帧不填充 |
| `security_request_id` | can_config | 诊断请求 CAN ID（发送）                     |
| `security_response_id` | can_config | 诊断响应 CAN ID（接收）                     |
| `allowed_ids`       | can_config | 日志记录的 CAN ID 白名单                    |
| `LOOP_COUNT`        | 执行控制   | 循环执行次数                                |
| `EXCEL_PLAN_PATH`   | 文件路径   | 输入 Excel 路径                             |
| `OUTPUT_DIR`        | 文件路径   | 结果输出目录                                |
| `DLL_PATH`          | 文件路径   | VW 安全算法 DLL 路径                        |
| `SERIAL_PORT`       | 继电器     | 串口号（如 COM99）                          |
| `RELAY_COMMANDS`    | 继电器     | 继电器 HEX 命令字典                         |
| `SERVICE22_EXPANSION_STEPS` | 22服务 | 展开步骤列表（可注释/取消注释）             |
| `SERVICE27_CONFIG`  | 27服务     | KEY 安全访问参数                            |
| `SERVICE27_CONFIG_1` | 27服务    | KEY1 安全访问参数                           |
| `SERVICE_BA27_CONFIG` | BA安全访问 | KEYBB 安全访问参数                        |
| `FILE_SEGMENT_CONFIG` | 分段写入  | Excel/log 分段参数                          |

### Excel 输入约定

| 请求数据值 | 含义                       |
|-----------|----------------------------|
| `KEY`     | VW安全访问（2701/2702）    |
| `KEY1`    | VW安全访问（2705/2706）    |
| `KEYBB`   | BA安全访问（BB01/BB02）    |
| `TP3E`    | 启动 Tester Present 3E00   |
| `STP3E`   | 停止 Tester Present        |
| `KL30ON`/`KL30OFF` | 继电器 KL30 控制 |
| `KL15ON`/`KL15OFF` | 继电器 KL15 控制 |
| `22XXXX`  | 22服务读取 DID             |
| `28XXXX`  | 28服务通信控制             |
| `10XX`    | 10服务会话控制             |
| `31XXXX`  | 31服务例程控制             |
| `NA`/`N/A`/空 | 期望HEX为空=只要求肯定响应 |

---

## 常见开发工作流

### 工作流1：新增 UDS 服务

1. 创建新 Handler 类，继承 `UDSServiceHandler`
2. 重写需要自定义的方法（`get_expanded_configs` / `execute_test` 等）
3. 在 `get_handler()` 中添加路由规则
4. 如需特殊处理，在 `main()` 的 `send_and_receive_can_messages()` 中添加分支

```python
class ServiceXXHandler(UDSServiceHandler):
    def execute_test(self, bus, arb_id, cfg, logger):
        # 自定义执行逻辑
        self.send(bus, arb_id, cfg, logger)
        self.receive(bus, cfg, logger)
        self.analyze_and_judge(cfg, logger)
```

### 工作流2：修改22服务展开步骤

直接编辑 `SERVICE22_EXPANSION_STEPS` 列表的注释即可：

```python
SERVICE22_EXPANSION_STEPS = [
    {'request': '1001', 'expected': '50 01', 'name': '切默认会话', 'wait': 0.2},
    {'request': '22{DID}', 'expected': '62 {DID}', 'name': '22读取', 'wait': 0.2},
    # 注释掉不需要的步骤
]
```

### 工作流3：修改安全访问算法

- **VW 算法**：替换 `DLL_PATH` 指向的 DLL 文件
- **BA 算法**：修改 `calculate_key_level1()` 函数中的 `xor_bytes` 或位操作逻辑

### 工作流4：添加新的比对模式

在 `CanMessageLogger.apply_comparison()` 中添加新分支：
- 当前支持：数据库比对、HEX 精确匹配、HEX 前缀匹配、7F 否定响应期望
- 新增模式：按需扩展 `comparison_mode` 分支

### 工作流5：修改串口继电器命令

编辑 `RELAY_COMMANDS` 字典：
```python
RELAY_COMMANDS = {
    'KL15 on': 'A0 02 00 A2',
    # 按实际硬件协议修改 HEX 命令
}
```

---

## 线程模型

| 线程               | 说明                           | 生命周期               |
|--------------------|--------------------------------|------------------------|
| 主线程             | 执行主流程，遍历测试用例       | 整个程序运行期间       |
| 周期发送线程       | 每 N 秒发送一条 CAN 报文       | 用例开始→stop_event.set() |
| 通道1监听线程      | 监听 channel 1 全部报文         | 主流程期间             |
| TesterPresent      | 无独立线程，主循环分片检查     | TP3E→STP3E             |

**重要约束**：`bus.send()` 调用必须通过 `BUS_SEND_LOCK` 保护（BMAPI 非线程安全）。

---

## 代码约定

1. **配置文件头部集中**：所有可配置项必须在【项目配置区】
2. **Handler 模式**：每种 UDS 服务一个 Handler 类
3. **命名约定**：
   - 服务处理器：`Service{XX}Handler`
   - 配置字典：`UPPER_SNAKE_CASE`
   - 工具函数：`snake_case`
4. **错误处理**：外层 try/except + traceback，不中断整体流程
5. **资源管理**：CAN bus、串口、线程均在 finally 中清理
6. **填充字节**：V9.4 使用 `pad_byte=0xCC`，V7 使用 `0x00`
7. **ISO-TP 编码**：parse_hex_input 自动判断单帧/扩展单帧/多帧
8. **日志字段**：统一使用中文列名（面向国内团队）

---

## 关键约束与陷阱

### 必须注意

1. **BMAPI DLL 路径**：`BMAPI64.dll` 必须在工作目录或 PATH 中
2. **ADB 可用性**：终端必须通过 ADB 连接，否则数据库查询跳过
3. **串口可用性**：如串口不存在，`init_serial()` 返回 None，不影响 CAN 测试
4. **DLL 可用性**：如 DLL 不存在，安全访问测试将失败，不影响其他测试
5. **CAN FD 模式**：`can_mode=2`时需要硬件支持 BRS
6. **异常不中断**：单条测试失败不影响后续执行

### 常见问题

| 问题             | 原因                     | 解决                       |
|------------------|--------------------------|----------------------------|
| CAN 总线打开失败 | BUSMASTER 未启动         | 检查 BUSMASTER 状态        |
| 数据库查询失败   | ADB 未连接或权限不足     | `adb devices` 检查         |
| 串口打开失败     | COM 口不存在             | 检查 `SERIAL_PORT` 配置    |
| 安全访问失败     | DLL 不存在或算法不匹配   | 检查 `DLL_PATH` 和种子格式 |
| 数据丢失         | 文件未刷新               | V9.4 已使用 `flush()` 确保 |
| 线程冲突         | 多线程调用 bus.send()    | 必须使用 `BUS_SEND_LOCK`   |

---

## 测试与验证

### 运行方式

```bash
# 单次执行
python examples/can_send_receive_to_excelV6.1.py

# 指定循环次数（V7 版本支持）
python examples/dazhong2.0.py 5
```

### 输出文件

| 文件                                    | 说明                     |
|-----------------------------------------|--------------------------|
| `OUTPUT_DIR/CAN测试结果_{timestamp}.xlsx` | Excel 测试结果（主输出） |
| `LISTENER_LOG_DIR/can_channel1_*.log`    | 通道1监听日志            |

### 结果字段

结果列的可能值：
- `通过` — 与期望值匹配
- `失败` — 与期望值不匹配
- `否定响应` — 收到 7F XX XX
- `失败（期望肯定响应）` — 期望肯定但收到否定或无响应
- `超时无响应` — 超时未收到任何响应

---

## 新增需求时的检查清单

- [ ] 是否需要修改【项目配置区】？
- [ ] 是否需要新增 Excel 输入列？检查 `load_send_configs_from_excel()`
- [ ] 是否需要新增 UDS 服务？检查 `get_handler()` 路由
- [ ] 是否需要新增比对模式？检查 `apply_comparison()`
- [ ] 是否需要新增日志字段？检查 `_create_log_entry()` 和 Excel 列定义
- [ ] 是否需要新增特殊触发值？
- [ ] 是否影响线程安全？检查 `BUS_SEND_LOCK` 使用
- [ ] 是否需要紧急保存支持？
- [ ] 是否需要在 `finally` 中添加清理逻辑？
