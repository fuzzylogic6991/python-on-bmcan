# -*- coding: utf-8 -*-
"""
TBOX 项目 CAN 测试配置文件
从 config_dazhong.py 复制，按需修改
"""

import threading
import serial
import ctypes

# ────────────────────────────────────────────────
# 一、CAN通信配置
# ────────────────────────────────────────────────
DEFAULT_CAN_CONFIG = {
    'interface': 'bmcan',            # ★ CAN卡类型（bmcan/vector等）
    'channel': 0,                    # ★ 通道号
    'bitrate': 500000,               # ★ 波特率
    'data_bitrate': 2000000,         # ★ FD数据波特率
    'can_mode': 0,                   # ★ CAN模式：0=经典CAN, 1=FD(最大8字节，无BRS), 2=FD(最小八字节，有BRS), 3=FD(可变字节，BRS,短帧不填充)
    'is_extended_id': False,         # ★ 是否扩展帧
    'response_timeout': 2.0,         # 响应超时时间
    'wait_after_request': 0.2,       # 请求后等待时间
    'allowed_ids': {0x703, 0x783},   # ★ TODO: 改为TBOX的CAN ID
    'flow_control_data': [0x30, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],
    'multi_frame_gap': 0.02,         # 多帧间隔
    'fc_timeout': 1.0,               # 流控超时
    'security_request_id': 0x703,    # ★ TODO: 改为TBOX的请求ID
    'security_response_id': 0x783,   # ★ TODO: 改为TBOX的响应ID
    'security_timeout': 2.0,         # 安全访问超时
    'bus_send_timeout': 1.0,         # ★ bus.send超时（秒），防止ECU断联后阻塞
    'tp3e_interval': 3.0,             # ★ 3E00发送周期（秒）
    'tp3e_wait_after': 0.1,           # ★ 3E00发送后等待（秒）
    'tp3e_arb_id': 0x703,             # ★ TODO: 改为TBOX的3E00发送ID
    'tres': True,                     # 自动发送流控
    'fd_pad_to_8': True,             # FD帧填充到8字节
    'pad_byte': 0xCC,                # ★★★ 填充字节
}

# ────────────────────────────────────────────────
# 二、执行控制配置
# ────────────────────────────────────────────────
LOOP_COUNT = 1                       # ★ 循环执行次数
LOOP_GAP = 1.0                       # 循环间隔时间（秒）

# ────────────────────────────────────────────────
# 三、文件路径配置
# ────────────────────────────────────────────────
DLL_PATH = r"E:\SRwork_DQJBF\400.project\5.duola\TBOX_GenerateKeyExImpl.dll"                       # ★ TODO: 安全算法DLL路径（如不需要留空）

# ────────────────────────────────────────────────
# 三.1、安全算法DLL函数配置 ★★★
# 更换DLL时只需修改此配置，无需改动 uds.py
# ────────────────────────────────────────────────
SECURITY_DLL_CONFIG = {
    # ★ DLL导出函数名
    'function_name': 'GenerateKeyEx',

    # ★ 函数参数类型列表 (对应C函数签名的每个参数)
    'argtypes': [
        ctypes.POINTER(ctypes.c_uint8),   # 0: seed 数组指针
        ctypes.c_uint32,                   # 1: seed 长度
        ctypes.c_uint32,                   # 2: 安全等级 (1/3/5)
        ctypes.POINTER(ctypes.c_uint8),   # 3: variant 指针 (传空)
        ctypes.POINTER(ctypes.c_uint8),   # 4: key 输出缓冲区指针
        ctypes.c_uint32,                   # 5: key 缓冲区最大长度
        ctypes.POINTER(ctypes.c_uint32),  # 6: 实际 key 长度 (输出)
    ],

    # ★ 函数返回值类型
    'restype': ctypes.c_uint32,

    # ★ 调用参数映射 (按 argtypes 顺序，决定 VW_Seed2Key 如何传参)
    #   可用占位符:
    #     seed        → seed 字节数组
    #     seed_len    → len(seed)
    #     level       → security_level (27服务的子功能码)
    #     variant     → variant 空字节数组
    #     key         → key 输出缓冲区
    #     key_max     → KEY_BUFFER_SIZE
    #     key_len_out → byref(key_len)，返回实际密钥长度
    'arg_map': ['seed', 'seed_len', 'level', 'variant', 'key', 'key_max', 'key_len_out'],

    # ★ 密钥输出缓冲区大小 (字节)
    'key_buffer_size': 16,
}
#
# 示例：如果换成一个简单的 2 参数 DLL (seed*, key*)，配置如下：
# SECURITY_DLL_CONFIG = {
#     'function_name': 'VW_Seed2Key',
#     'argtypes': [
#         ctypes.POINTER(ctypes.c_uint8),   # seed
#         ctypes.POINTER(ctypes.c_uint8),   # key
#     ],
#     'restype': ctypes.c_int16,
#     'arg_map': ['seed', 'key'],
#     'key_buffer_size': 4,
# }

EXCEL_PLAN_PATH = r"E:\Edownload\bigtmp\input_duola_Diag.xlsx"  # ★ TODO: 测试用例Excel路径
OUTPUT_DIR = r"E:\Edownload\bigtmp"                  # ★ 结果输出目录
LISTENER_LOG_DIR = r"d:\Users\15959\Desktop\duola\DIAG"                   # ★ 监听日志目录
ENABLE_LISTENER_LOG = False                               # ★ 是否启用独立监听日志

# ────────────────────────────────────────────────
# 四、继电器配置（程控电源控制）
# ────────────────────────────────────────────────
SERIAL_PORT = "COM99"                # ★ TODO: 继电器串口号
SERIAL_BAUDRATE = 9600
SERIAL_BYTESIZE = serial.EIGHTBITS
SERIAL_PARITY = serial.PARITY_NONE
SERIAL_STOPBITS = serial.STOPBITS_ONE
SERIAL_TIMEOUT = 0.5

RELAY_COMMANDS = {
    # ★ TODO: 根据实际硬件协议修改
    'KL15 on': 'A0 02 00 A2',
    'KL15 off': 'A0 02 01 A3',
    'KL30 on': 'A0 01 00 A1',
    'KL30 off': 'A0 01 01 A2',
}

# ────────────────────────────────────────────────
# 五、22服务展开规则配置
# ────────────────────────────────────────────────
SERVICE22_EXPANSION_STEPS = [
    {'request': '22{DID}', 'expected': '62 {DID}', 'name': '22读取', 'wait': 0.2},
]

# ────────────────────────────────────────────────
# 六、27服务安全访问配置
# ────────────────────────────────────────────────
SERVICE27_CONFIG = {
    'session_request': '1003',
    'session_expected': '50 03',
    'seed_request': '2701',
    'seed_expected_sid': 0x67,
    'seed_expected_sub': 0x01,
    'key_request': '2702',
    'key_expected_sid': 0x67,
    'key_expected_sub': 0x02,
}

SERVICE27_CONFIG_1 = {
    'session_request': '1002',
    'session_expected': '50 02',
    'seed_request': '2705',
    'seed_expected_sid': 0x67,
    'seed_expected_sub': 0x05,
    'key_request': '2706',
    'key_expected_sid': 0x67,
    'key_expected_sub': 0x06,
}

# ────────────────────────────────────────────────
# 六、27服务安全访问配置 (等级3: 2703/2704)
# ────────────────────────────────────────────────
SERVICE27_CONFIG_3 = {
    'session_request': '1002',
    'session_expected': '50 02',
    'seed_request': '2703',
    'seed_expected_sid': 0x67,
    'seed_expected_sub': 0x03,
    'key_request': '2704',
    'key_expected_sid': 0x67,
    'key_expected_sub': 0x04,
    'skip_session': True,             # ★ 跳过会话切换，直接请求种子
}

# ────────────────────────────────────────────────
# 七、BA安全访问配置
# ────────────────────────────────────────────────
SERVICE_BA27_CONFIG = {
    'seed_request': 'BA01',
    'seed_expected_sid': 0xFA,
    'seed_expected_sub': 0x01,
    'key_request': 'BA02',
}

# ────────────────────────────────────────────────
# 八、文件分段写入配置
# ────────────────────────────────────────────────
FILE_SEGMENT_CONFIG = {
    'excel_max_rows': 4000,
    'log_max_size_mb': 10,
    'auto_save_interval': 500,
    'enable_segment_write': True,
}

# ★★★ 全局 bus.send 线程锁 ★★★
BUS_SEND_LOCK = threading.Lock()
