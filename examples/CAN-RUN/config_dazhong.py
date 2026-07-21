# -*- coding: utf-8 -*-
"""
大众项目 CAN 测试配置文件
复制此文件，按需修改，即可适配不同项目
"""

import threading
import serial

# ────────────────────────────────────────────────
# 一、CAN通信配置
# ────────────────────────────────────────────────
DEFAULT_CAN_CONFIG = {
    'interface': 'bmcan',            # ★ CAN卡类型（bmcan/vector等）
    'channel': 0,                    # ★ 通道号
    'bitrate': 500000,               # ★ 波特率
    'data_bitrate': 2000000,         # ★ FD数据波特率
    'can_mode': 2,                   # ★ CAN模式：0=经典CAN, 1=FD(无BRS), 2=FD(有BRS), 3=FD(BRS,短帧不填充)
    'is_extended_id': False,         # ★ 是否扩展帧
    'response_timeout': 2.0,         # 响应超时时间
    'wait_after_request': 0.2,       # 请求后等待时间
    'allowed_ids': {0x711, 0x719},   # ★ 日志记录的CAN ID（tx/rx）
    'flow_control_data': [0x30, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],  # 流控帧数据
    'multi_frame_gap': 0.02,         # 多帧间隔
    'fc_timeout': 1.0,               # 流控超时
    'security_request_id': 0x711,    # ★ 安全访问请求ID（tx）
    'security_response_id': 0x719,   # ★ 安全访问响应ID（rx）
    'security_timeout': 2.0,         # 安全访问超时
    'bus_send_timeout': 1.0,         # ★ bus.send超时（秒），防止ECU断联后阻塞
    'tp3e_interval': 3.0,             # ★ 3E00发送周期（秒）
    'tp3e_wait_after': 0.1,           # ★ 3E00发送后等待（秒）
    'tp3e_arb_id': 0x711,             # ★ 3E00发送ID
    'tres': True,                     # 自动发送流控
    'fd_pad_to_8': True,             # FD帧填充到8字节
    'pad_byte': 0xCC,                # ★★★ 填充字节（保留1.2的0xCC）
}

# ────────────────────────────────────────────────
# 二、执行控制配置
# ────────────────────────────────────────────────
LOOP_COUNT = 1                       # ★ 循环执行次数，1 = 只执行一次
LOOP_GAP = 5.0                       # 循环间隔时间（秒）

# ────────────────────────────────────────────────
# 三、文件路径配置
# ────────────────────────────────────────────────
DLL_PATH = r"E:\Edownload\input\VW_seed_to_key.dll"        # ★ 安全算法DLL路径
EXCEL_PLAN_PATH = r"E:\Edownload\input\CAN测试用例_0x711_19条.xlsx"  # ★ Excel配置文件路径
OUTPUT_DIR = r"E:\Edownload\input\output"                          # ★ 结果输出目录
LISTENER_LOG_DIR = r"E:\Edownload\ouput"               # ★ 监听日志目录
ENABLE_LISTENER_LOG = False                       # ★ 是否启用独立监听日志（大多数时候不需要）

# ────────────────────────────────────────────────
# 四、继电器配置（程控电源控制）
# ────────────────────────────────────────────────
SERIAL_PORT = "COM99"                # ★ 继电器串口号
SERIAL_BAUDRATE = 9600               # 波特率
SERIAL_BYTESIZE = serial.EIGHTBITS   # 数据位
SERIAL_PARITY = serial.PARITY_NONE   # 校验位
SERIAL_STOPBITS = serial.STOPBITS_ONE  # 停止位
SERIAL_TIMEOUT = 0.5                 # 超时时间

RELAY_COMMANDS = {
    # ★ 继电器命令（根据实际硬件协议修改）
    'KL15 on': 'A0 02 00 A2',        # KL15上电
    'KL15 off': 'A0 02 01 A3',       # KL15下电
    'KL30 on': 'A0 01 00 A1',        # KL30上电
    'KL30 off': 'A0 01 01 A2',       # KL30下电
}

# ────────────────────────────────────────────────
# 五、22服务展开规则配置 ★★★
# ────────────────────────────────────────────────
# ★ 修改此列表可自定义22服务的展开步骤
# ★ 每个元素代表一个步骤，格式：{请求, 期望响应前缀, 步骤名称, 等待时间}
# ★ {DID} 会自动替换为实际的DID值
SERVICE22_EXPANSION_STEPS = [
    {'request': '22{DID}', 'expected': '62 {DID}', 'name': '22读取', 'wait': 0.2},
]

# ────────────────────────────────────────────────
# 六、27服务安全访问配置 ★★★
# ────────────────────────────────────────────────
# ★ VW安全访问流程参数（Excel触发值：KEY）
SERVICE27_CONFIG = {
    'session_request': '1003',        # ★ 切换会话请求（默认扩展会话）
    'session_expected': '50 03',      # ★ 会话期望响应
    'seed_request': '2701',           # ★ 请求种子
    'seed_expected_sid': 0x67,        # ★ 种子响应SID
    'seed_expected_sub': 0x01,        # ★ 种子响应子功能
    'key_request': '2702',            # ★ 发送密钥请求
    'key_expected_sid': 0x67,         # ★ 密钥验证成功SID
    'key_expected_sub': 0x02,         # ★ 密钥验证成功子功能
}

# ★ VW安全访问流程参数（Excel触发值：KEY1）
SERVICE27_CONFIG_1 = {
    'session_request': '1002',        # ★ 切换会话请求（刷新会话）
    'session_expected': '50 02',      # ★ 会话期望响应
    'seed_request': '2705',           # ★ 请求种子
    'seed_expected_sid': 0x67,        # ★ 种子响应SID
    'seed_expected_sub': 0x05,        # ★ 种子响应子功能
    'key_request': '2706',            # ★ 发送密钥请求
    'key_expected_sid': 0x67,         # ★ 密钥验证成功SID
    'key_expected_sub': 0x06,         # ★ 密钥验证成功子功能
}

# ────────────────────────────────────────────────
# 七、BA安全访问配置 ★★★
# ────────────────────────────────────────────────
# ★ BA安全访问流程参数（Excel触发值：KEYBB）
SERVICE_BA27_CONFIG = {
    'seed_request': 'BA01',           # ★ 请求种子
    'seed_expected_sid': 0xFA,        # ★ 种子响应SID
    'seed_expected_sub': 0x01,        # ★ 种子响应子功能
    'key_request': 'BA02',            # ★ 发送密钥请求
}

# ────────────────────────────────────────────────
# 八、文件分段写入配置
# ────────────────────────────────────────────────
FILE_SEGMENT_CONFIG = {
    'excel_max_rows': 4000,          # Excel单段最大行数
    'log_max_size_mb': 10,           # Log单文件最大大小（MB）
    'auto_save_interval': 500,        # 自动保存间隔（测试用例数）
    'enable_segment_write': True,    # 是否启用分段写入
}

# ★★★ 全局 bus.send 线程锁：bmcan 后端非线程安全，强制串行发送 ★★★
BUS_SEND_LOCK = threading.Lock()
