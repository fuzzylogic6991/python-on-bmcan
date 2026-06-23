#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CAN诊断测试工具 v9.4 - 配置集中版 + 分段写入保护

★★★ 新项目只需修改文件顶部【项目配置区】（约第30-145行）★★★

配置区包含：
  一、CAN通信配置（第36行）      - CAN ID、波特率、模式等
  二、执行控制配置（第59行）      - 循环次数、间隔时间
  三、文件路径配置（第65行）      - DLL、Excel、输出目录
  四、继电器配置（第73行）        - 串口、命令
  五、22服务展开规则（第94行）    - 展开步骤列表
  六、27服务安全访问（第115行）   - VW安全访问参数
  七、BA安全访问（第130行）       - BA安全访问参数
  八、文件分段写入配置（第145行）  - Excel和Log分段设置
    Excel触发值: TP3E(启动) / STP3E(停止)
以下代码为通用逻辑，一般无需修改。
"""

import os
import time
import ctypes
import threading
import signal
import atexit
import sys
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Union


import pandas as pd
import can
import serial
import binascii
import subprocess


# ══════════════════════════════════════════════════════════════
# 【项目配置区】新项目只需修改此区域 ★★★
# ══════════════════════════════════════════════════════════════

# ────────────────────────────────────────────────
# 一、CAN通信配置
# ────────────────────────────────────────────────
DEFAULT_CAN_CONFIG = {
    'bustype': 'bmcan',              # ★ CAN卡类型（bmcan/vector等）
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
DLL_PATH = r"D:\python\student\can_send\main\VW_seed_to_key.dll"        # ★ 安全算法DLL路径
EXCEL_PLAN_PATH = r"D:\python\student\can_send\run_excel\大众\临时bug验证.xlsx"  # ★ Excel配置文件路径
OUTPUT_DIR = r"D:\python\student\can_send\log"                          # ★ 结果输出目录
LISTENER_LOG_DIR = r"D:\python\student\can_send\listener"               # ★ 监听日志目录

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
    # # 步骤1：切默认会话
    # {'request': '1001', 'expected': '50 01', 'name': '切默认会话', 'wait': 0.2},
    # # 步骤2：22读取DID
    {'request': '22{DID}', 'expected': '62 {DID}', 'name': '22读取', 'wait': 0.2},
    # # 步骤3：切扩展会话
    # {'request': '1003', 'expected': '50 03', 'name': '切扩展会话', 'wait': 0.2},
    # # 步骤4：22再读取
    # {'request': '22{DID}', 'expected': '62 {DID}', 'name': '22再读取', 'wait': 0.2},
    # # 步骤5：切刷新会话
    # {'request': '1002', 'expected': '50 02', 'name': '切刷新会话', 'wait': 6.0},    # ★ 6秒
    # # 步骤6：刷新会话下22再读取
    # {'request': '22{DID}', 'expected': '62 {DID}', 'name': '刷新会话下22再读取', 'wait': 0.2},
    # # 步骤7：切回默认会话
    # {'request': '1001', 'expected': '50 01', 'name': '切回默认会话', 'wait': 60.0}, # ★ 60秒
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
# 八、文件分段写入配置 ★★★ 新增
# ────────────────────────────────────────────────
FILE_SEGMENT_CONFIG = {
    'excel_max_rows': 4000,          # Excel单段最大行数
    'log_max_size_mb': 10,           # Log单文件最大大小（MB）
    'auto_save_interval': 500,        # 自动保存间隔（测试用例数）
    'enable_segment_write': True,    # 是否启用分段写入
}

# ★★★ 全局 bus.send 线程锁：bmcan 后端非线程安全，强制串行发送 ★★★
BUS_SEND_LOCK = threading.Lock()

# ══════════════════════════════════════════════════════════════
# 【配置区结束】以上为项目配置，以下为通用代码 ★★★
# ══════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════
# 第1层：基础层（以下为通用代码，一般无需修改）
# ══════════════════════════════════════════════════════════════

FD_DLC_MAP = {
    0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8,
    9: 12, 10: 16, 11: 20, 12: 24, 13: 32, 14: 48, 15: 64
}


# ────────────────────────────────────────────────
# 工具函数
# ────────────────────────────────────────────────
def load_vw_dll():
    """加载VW安全算法DLL"""
    try:
        dll = ctypes.CDLL(DLL_PATH)
        dll.VW_Seed2Key.argtypes = [
            ctypes.POINTER(ctypes.c_uint8),  # seed
            ctypes.POINTER(ctypes.c_uint8)  # key
        ]
        dll.VW_Seed2Key.restype = ctypes.c_int16
        return dll
    except Exception as e:
        print(f"⚠ 警告：加载DLL失败: {e}")
        return None


VW_DLL = load_vw_dll()


def calculate_key_level1(seed_bytes: List[int]) -> Optional[List[int]]:
    """BA安全访问算法：4字节种子 → 4字节密钥"""
    if len(seed_bytes) != 4:
        return None
    xor_bytes = [0x26, 0xBF, 0x6D, 0x96]
    data = [(seed_bytes[i] ^ xor_bytes[i]) & 0xFF for i in range(4)]
    key = [0] * 4
    key[0] = ((data[0] & 0x0F) << 4) | (data[3] & 0x0F)
    key[1] = ((data[1] & 0x0F) << 4) | ((data[2] & 0xF0) >> 4)
    key[2] = ((data[3] & 0x1C) << 3) | ((data[0] & 0x3E) >> 1)
    key[3] = (data[2] & 0xF0) | ((data[1] & 0x0F) >> 4)
    return [k & 0xFF for k in key]

def VW_Seed2Key(seed: bytes) -> bytes:
    """
    调用DLL计算安全密钥
    :param seed: 4字节种子
    :return: 4字节密钥
    """
    if VW_DLL is None:
        raise RuntimeError("VW DLL 未正确加载")

    if len(seed) != 4:
        raise ValueError(f"种子长度必须为4字节，当前: {len(seed)}")

    seed_array = (ctypes.c_uint8 * 4)(*seed)
    key_array = (ctypes.c_uint8 * 4)()

    result = VW_DLL.VW_Seed2Key(seed_array, key_array)

    if result != 0:
        raise RuntimeError(f"VW_Seed2Key 返回错误: {result}")

    return bytes(key_array)


def get_fd_dlc(length: int) -> int:
    for dlc, size in sorted(FD_DLC_MAP.items(), key=lambda x: x[1]):
        if length <= size:
            return dlc
    raise ValueError(f"数据长度 {length} 超过 CAN FD 最大 64 字节")


def get_fd_padded_length(dlc: int) -> int:
    return FD_DLC_MAP.get(dlc, 0)


def pad_data(data: List[int], config: Dict) -> List[int]:
    """填充数据，使用配置的填充字节"""
    can_mode = config['can_mode']
    pad_byte = config.get('pad_byte', 0xCC)  # ★★★ 使用配置的填充字节
    # ★★★ 修复：can_mode=1（FD无BRS）也支持扩展单帧，最大64字节 ★★★
    dlc_max = 64 if can_mode in (1, 2, 3) else 8
    padded = list(data)
    current_len = len(padded)

    if current_len > dlc_max:
        raise ValueError(f"数据长度 {current_len} 超过模式 {can_mode} 的最大值 {dlc_max}")

    if can_mode == 0:
        # 经典CAN：固定填充到8字节
        if current_len < 8:
            padded += [pad_byte] * (8 - current_len)
        return padded[:8]
    elif can_mode == 1:
        # ★★★ CAN FD无BRS：支持扩展单帧，按FD DLC填充 ★★★
        if current_len <= 8:
            # 8字节及以下：可选择填充到8字节
            if config.get('fd_pad_to_8', False) and current_len < 8:
                padded += [pad_byte] * (8 - current_len)
            return padded[:8] if config.get('fd_pad_to_8', False) else padded
        else:
            # 超过8字节：按FD DLC填充
            dlc = get_fd_dlc(current_len)
            padded_len = get_fd_padded_length(dlc)
            if current_len < padded_len:
                padded += [pad_byte] * (padded_len - current_len)
            return padded
    elif can_mode == 3:
        # ★★★ CAN FD BRS 不填充短帧：3-8字节直接返回不填充，9-64同mode=2 ★★★
        if current_len <= 8:
            return padded
        dlc = get_fd_dlc(current_len)
        padded_len = get_fd_padded_length(dlc)
        if current_len < padded_len:
            padded += [pad_byte] * (padded_len - current_len)
        return padded
    else:
        # CAN FD有BRS：按FD DLC填充
        if config.get('fd_pad_to_8', False) and current_len < 8:
            padded += [pad_byte] * (8 - current_len)
            return padded
        dlc = get_fd_dlc(current_len)
        padded_len = get_fd_padded_length(dlc)
        if current_len < padded_len:
            padded += [pad_byte] * (padded_len - current_len)
        return padded


def parse_hex_input(hex_str: str, is_request: bool = True,
                    config: Dict = DEFAULT_CAN_CONFIG) -> Union[List[int], List[List[int]]]:
    hex_str = hex_str.strip().upper().replace(' ', '')
    if not hex_str or hex_str in ('NA', 'N/A', '无', 'NONE', 'NULL', '-'):
        return [] if is_request else []
    try:
        bytes_list = [int(hex_str[i:i + 2], 16) for i in range(0, len(hex_str), 2)]
    except ValueError:
        return [] if is_request else []
    if not is_request:
        return bytes_list

    can_mode = config['can_mode']
    dlc_max = 8 if can_mode in (0, 1) else 64
    total_len = len(bytes_list)

    # ★★★ 修复：经典CAN模式（can_mode=0）最大单帧7字节 ★★★
    if can_mode == 0:
        if total_len <= 7:
            return pad_data([total_len] + bytes_list, config)
    # ★★★ 修复：CAN FD模式（can_mode=1或2）支持扩展单帧格式 ★★★
    # 当数据长度在8-62字节之间时，使用扩展单帧格式：[0x00, length, data...]
    elif can_mode in (1, 2, 3):
        if total_len <= 7:
            # 普通单帧：[length, data...]
            return pad_data([total_len] + bytes_list, config)
        elif 8 <= total_len <= 62:
            # ★★★ 扩展单帧格式：[0x00, length, data...] ★★★
            return pad_data([0x00, total_len] + bytes_list, config)

    # 多帧处理（数据长度超过单帧最大容量）
    frames: List[List[int]] = []
    ff_prefix_len = 2
    cf_prefix_len = 1
    ff_data_len = dlc_max - ff_prefix_len
    cf_data_len = dlc_max - cf_prefix_len

    len_high = (total_len >> 8) & 0x0F
    len_low = total_len & 0xFF
    ff = [0x10 | len_high, len_low] + bytes_list[:ff_data_len]
    frames.append(pad_data(ff, config))

    remaining = bytes_list[ff_data_len:]
    seq = 1
    while remaining:
        chunk = remaining[:cf_data_len]
        remaining = remaining[cf_data_len:]
        cf = [0x20 | (seq & 0x0F)] + chunk
        frames.append(pad_data(cf, config))
        seq = (seq + 1) % 16
    return frames


def bytes_to_ascii(data: List[int]) -> str:
    return ''.join(chr(b) if 32 <= b <= 126 else '.' for b in data)


# ★★★ 修改：extract_isotp_payload 增加过滤 TesterPresent 的逻辑 ★★★
def extract_isotp_payload(frames: List[List[int]], can_mode: int = 1,
                          response_id: int = 0x7BE) -> List[int]:
    """
    从响应帧中提取 ISO-TP 载荷
    ★ 新增：过滤掉 TesterPresent 响应（02 7E 00）
    """
    if not frames:
        return []

    # ★★★ 关键修复：过滤掉 TesterPresent 响应 ★★★
    # TesterPresent 响应格式：02 7E 00 CC CC CC CC CC
    filtered_frames = []
    for frame in frames:
        if len(frame) >= 2:
            # 跳过 TesterPresent 响应（PCI=0x02, SID=0x7E）
            if frame[0] == 0x02 and frame[1] == 0x7E:
                continue
        filtered_frames.append(frame)

    if not filtered_frames:
        return []

    first_frame = filtered_frames[0]
    pci = first_frame[0]

    # 单帧处理
    if pci & 0xF0 == 0x00:
        if pci == 0x00 and len(first_frame) >= 2:
            dl = first_frame[1]
            if 8 <= dl <= 62:
                return list(first_frame[2:2 + dl])
            else:
                return []
        else:
            dl = pci & 0x0F
            max_dl = 8 if can_mode in (0, 1) else 15
            if dl > max_dl:
                return []
            return list(first_frame[1:1 + dl])

    # 多帧处理
    if pci & 0xF0 == 0x10:
        if can_mode in (2, 3):
            total_len = first_frame[1]
            payload_start = 2
        else:
            total_len = ((pci & 0x0F) << 8) | first_frame[1]
            payload_start = 2

        payload = first_frame[payload_start:]
        seq = 1

        for cf in filtered_frames[1:]:
            cf_pci = cf[0]
            # 跳过非连续帧（如额外的 TesterPresent）
            if cf_pci & 0xF0 != 0x20:
                continue
            if (cf_pci & 0x0F) != seq:
                continue
            payload.extend(cf[1:])
            seq = (seq + 1) % 16

        return payload[:total_len] if len(payload) > total_len else payload

    return []


# ────────────────────────────────────────────────
# ADB 查询
# ────────────────────────────────────────────────
def execute_adb_command(command: str) -> Optional[str]:
    try:
        subprocess.run(['adb', 'shell', 'sync'],
                       capture_output=True, text=True, check=True)
        result = subprocess.run(['adb', 'shell', command], capture_output=True, text=True, check=True)
        return result.stdout.strip()
    except Exception:
        return None


def query_tbox_data_by_id(param_id: int) -> Optional[str]:
    adb_command = (
        f"cd /oemdata/parameters/ && "
        f"sqlite3 paramsDb 'select * from paramstbl where id={param_id};'"
    )
    result = execute_adb_command(adb_command)
    if result:
        print(f"  数据库查询 (id={param_id}): {result}")
    else:
        print(f"  ⚠ 数据库查询失败 (id={param_id})")
    return result


# ────────────────────────────────────────────────
# 继电器控制
# ────────────────────────────────────────────────
def init_serial() -> Optional[serial.Serial]:
    try:
        ser = serial.Serial(
            port=SERIAL_PORT, baudrate=SERIAL_BAUDRATE,
            bytesize=SERIAL_BYTESIZE, parity=SERIAL_PARITY,
            stopbits=SERIAL_STOPBITS, timeout=SERIAL_TIMEOUT
        )
        ser.flushInput()
        ser.flushOutput()
        return ser
    except Exception:
        return None


def execute_relay_command(ser: serial.Serial, command_key: str, delay: float = 0.5) -> bool:
    if not ser or not ser.is_open:
        return False
    if command_key not in RELAY_COMMANDS:
        return False
    try:
        ser.flushInput()
        cmd_hex = RELAY_COMMANDS[command_key].replace(" ", "")
        send_data = binascii.unhexlify(cmd_hex)
        ser.write(send_data)
        time.sleep(delay)
        return True
    except Exception:
        return False


def close_serial(ser: Optional[serial.Serial]):
    if ser and ser.is_open:
        ser.close()


# ────────────────────────────────────────────────
# 周期发送线程
# ────────────────────────────────────────────────
def periodic_send(bus, arbitration_id: int, data: List[int], period: float,
                  stop_event: threading.Event, config: Dict, logger_obj: 'CanMessageLogger'):
    is_fd = config['can_mode'] > 0
    while not stop_event.is_set():
        msg = can.Message(
            arbitration_id=arbitration_id, data=data,
            is_extended_id=config['is_extended_id'], is_fd=is_fd
        )
        with BUS_SEND_LOCK:
            bus.send(msg)
        logger_obj.log_sent_message(msg, time.time())
        time.sleep(period)


# ────────────────────────────────────────────────
# 通道1 监听线程（分段写入版本）★★★ 新增
# ────────────────────────────────────────────────
def channel1_listener_segmented(stop_event: threading.Event, max_size_mb: int = 10):
    """
    分段写入的通道1监听线程
    - 每达到 max_size_mb 大小，关闭当前文件，创建新文件
    - 确保文件定期释放，避免长时间占用
    """
    os.makedirs(LISTENER_LOG_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    file_index = 0
    logfile = os.path.join(LISTENER_LOG_DIR, f"can_channel1_{timestamp}_part{file_index}.log")
    
    bus_kwargs = {
        'bustype': 'bmcan', 'channel': 1, 'bitrate': 500000,
        'is_fd': True, 'data_bitrate': 2000000, 'tres': True,
    }

    bus = None
    f = None
    try:
        bus = can.interface.Bus(**bus_kwargs)
        f = open(logfile, 'a', encoding='utf-8')
        f.write(f"=== 通道1 监听开始 {datetime.now()} ===\n\n")

        while not stop_event.is_set():
            try:
                msg = bus.recv(timeout=0.3)
                if msg:
                    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
                    id_str = f"0x{msg.arbitration_id:X}"
                    data_hex = ' '.join(f'{b:02X}' for b in msg.data)
                    line = f"{ts}  RX  {id_str:>10}  {'FD' if msg.is_fd else 'CL'}  DLC={msg.dlc:2}  {data_hex}"
                    f.write(line + '\n')
                    f.flush()  # ★★★ 立即刷新到磁盘

                    # ★★★ 检查文件大小，达到限制则切换新文件 ★★★
                    if os.path.getsize(logfile) > max_size_mb * 1024 * 1024:
                        f.write(f"=== 文件大小达到 {max_size_mb}MB，切换到新文件 ===\n")
                        f.flush()
                        f.close()  # ★★★ 关闭当前文件，释放句柄
                        
                        file_index += 1
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        logfile = os.path.join(LISTENER_LOG_DIR, 
                                               f"can_channel1_{timestamp}_part{file_index}.log")
                        f = open(logfile, 'a', encoding='utf-8')
                        f.write(f"=== 通道1 监听继续 {datetime.now()} ===\n\n")
                        print(f"  [通道1] 切换到新文件: {logfile}")

            except Exception:
                time.sleep(1)
    except Exception:
        pass
    finally:
        if bus:
            try:
                bus.shutdown()
            except:
                pass
        if f and not f.closed:
            f.write(f"=== 通道1 监听结束 {datetime.now()} ===\n")
            f.flush()
            f.close()  # ★★★ 确保文件关闭


# ────────────────────────────────────────────────
# CanMessageLogger（★ 修改 apply_comparison 支持前缀匹配 ★）
# ────────────────────────────────────────────────
class CanMessageLogger:
    def __init__(self, allowed_ids: Optional[set] = None):
        self.messages: List[Dict] = []
        self.test_results: Dict[tuple[str, str], Dict] = {}
        self.allowed_ids = allowed_ids or set()
        self.current_test_case_id: Optional[str] = None
        self.current_test_name: Optional[str] = None
        self.current_response_frames: List[List[int]] = []
        self.execution_order: List[tuple[str, str]] = []
        self.service28_monitor_logs: List[Dict] = []
        
        # ★★★ 分段写入相关 ★★★
        self.segment_files: List[str] = []  # 已保存的分段文件列表
        self.current_segment_index: int = 0
        self.messages_in_current_segment: int = 0
        self._auto_save_file: Optional[str] = None

    def _should_log(self, arbitration_id: int) -> bool:
        return not self.allowed_ids or arbitration_id in self.allowed_ids

    def _create_log_entry(self, msg_type: str, msg: can.Message, timestamp: float,
                          test_id: str = '', test_name: str = '') -> Dict:
        data_list = list(msg.data)
        return {
            '时间': datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
            '类型': msg_type,
            'CAN ID (十六进制)': f"0x{msg.arbitration_id:X}",
            '扩展帧': '是' if msg.is_extended_id else '否',
            '数据长度': msg.dlc,
            '数据 (十六进制)': ' '.join(f"{b:02X}" for b in data_list),
            '数据 (ASCII)': bytes_to_ascii(data_list),
            '测试用例ID': test_id or self.current_test_case_id or '',
            '测试项': test_name or self.current_test_name or '',
            '结果': '',
            '肯定响应值': '',
            '否定响应值': '',
            '22服务内容(hex)': '',
            '22服务内容(ascii)': '',
            '数据库获取值': '',
            '期望来源': '',
            '轮次': getattr(self, '_current_round', 1),  # ★ 新增
        }

    def log_service28_monitor(self, test_id: str, test_name: str,
                              monitor_id: int, msg: can.Message, timestamp: float):
        """记录28服务监控到的报文到独立列表"""
        data_list = list(msg.data)
        self.service28_monitor_logs.append({
            '时间': datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
            '测试用例ID': test_id,
            '测试项': test_name,
            '监控目标ID': f'0x{monitor_id:X}',
            '实际CAN ID': f'0x{msg.arbitration_id:X}',
            '数据长度': msg.dlc,
            '数据 (十六进制)': ' '.join(f'{b:02X}' for b in data_list),
            '数据 (ASCII)': bytes_to_ascii(data_list),
        })

    def log_sent_message(self, msg: can.Message, timestamp: float):
        if not self._should_log(msg.arbitration_id):
            return
        entry = self._create_log_entry('发送', msg, timestamp)
        self.messages.append(entry)

    def log_received_message(self, msg: can.Message, timestamp: float):
        if not self._should_log(msg.arbitration_id):
            return
        entry = self._create_log_entry('接收', msg, timestamp)
        self.messages.append(entry)
        self.current_response_frames.append(list(msg.data))

    def finalize_and_analyze_response(self, response_id: int = 0x7BE):
        """分析响应，增加对28服务响应（68 xx）的处理"""
        if not self.current_response_frames:
            return

        payload = extract_isotp_payload(self.current_response_frames, response_id=response_id)
        if not payload:
            return

        key = (self.current_test_case_id, self.current_test_name)
        result_dict = {
            '肯定响应值': '',
            '否定响应值': '',
            '22服务内容(hex)': '',
            '22服务内容(ascii)': '',
            '结果': '未知',
            '数据库获取值': '',
            '期望来源': '',
        }

        sid = payload[0] if payload else None
        FILL_BYTE = 0x00

        if sid == 0x7F and len(payload) >= 3:
            result_dict['否定响应值'] = f"7F {payload[1]:02X} {payload[2]:02X}"
            result_dict['结果'] = '否定响应'
        else:
            if sid == 0x62 and len(payload) >= 3:
                # 22服务响应
                data_part = payload[3:]
                while data_part and data_part[-1] == FILL_BYTE:
                    data_part.pop()
                result_dict['肯定响应值'] = f"62 {payload[1]:02X} {payload[2]:02X}"
                result_dict['22服务内容(hex)'] = ' '.join(f"{b:02X}" for b in data_part)
                result_dict['22服务内容(ascii)'] = bytes_to_ascii(data_part)
            elif sid == 0x68:
                # ★ 28服务响应（新增）
                result_dict['肯定响应值'] = f"68 {payload[1]:02X}" if len(payload) >= 2 else "68"
            else:
                # 其他服务响应
                while payload and payload[-1] == FILL_BYTE:
                    payload.pop()
                result_dict['肯定响应值'] = ' '.join(f"{b:02X}" for b in payload)

            result_dict['结果'] = '通过'

        self.test_results[key] = result_dict
        self.execution_order.append(key)
        self.current_response_frames.clear()

    # ★★★ 修改：apply_comparison 支持前缀匹配 ★★★
    def apply_comparison(self, cfg: Dict):
        tid = cfg['test_case_id']
        tname = cfg['test_name']
        result = self.test_results.get((tid, tname), {})

        # 如果没有结果，尝试不带后缀的匹配（兼容原有逻辑）
        if not result:
            # 尝试模糊匹配
            for (k_tid, k_tname), res in self.test_results.items():
                if k_tid == tid and tname.startswith(k_tname.split(' - ')[0]):
                    if tname == k_tname:
                        result = res
                        break

        if not result:
            return

        # 获取比对模式（默认精确匹配）
        comparison_mode = cfg.get('_comparison_mode', 'exact')

        if cfg.get('expected_db_id') is not None:
            db_val = query_tbox_data_by_id(cfg['expected_db_id']) or ''
            ascii_val = result.get('22服务内容(ascii)', '')
            passed = (ascii_val == db_val)
            result['结果'] = '通过' if passed else '失败'
            result['数据库获取值'] = db_val
            result['期望来源'] = f"DB id={cfg['expected_db_id']}"

        elif cfg.get('expected_hex_str'):
            exp_str_raw = str(cfg['expected_hex_str']).strip().upper()

            if not exp_str_raw or exp_str_raw in ('NA', 'N/A', '无', 'NONE', 'NULL', '-'):
                # 空值 = 只要求肯定响应
                actual = result.get('肯定响应值', '')
                is_positive = bool(actual) and not actual.strip().startswith('7F')
                result['结果'] = '通过' if is_positive else '失败（期望肯定响应）'
                result['期望来源'] = '期望HEX（空=要求肯定响应）'

            elif exp_str_raw.startswith('7F'):
                # 7F开头 = 要求否定响应
                actual_neg = result.get('否定响应值', '')
                is_negative = bool(actual_neg) and actual_neg.strip().startswith('7F')
                result['结果'] = '通过' if is_negative else '失败（期望否定响应）'
                result['期望来源'] = '期望HEX（7F开头=要求否定响应）'

            elif comparison_mode == 'prefix':
                # ★★★ 新增：前缀匹配模式 ★★★
                # 用于22服务步骤2/4：只检查响应是否以指定前缀开头
                actual = result.get('肯定响应值', '')
                passed = actual.strip().startswith(exp_str_raw)
                result['结果'] = '通过' if passed else '失败'
                result['期望来源'] = f'期望前缀（{exp_str_raw}）'

            else:
                # 精确匹配
                exp_bytes = parse_hex_input(cfg['expected_hex_str'], is_request=False, config=DEFAULT_CAN_CONFIG)
                exp_str = ' '.join(f"{b:02X}" for b in exp_bytes).strip()
                actual = result.get('肯定响应值', '') or result.get('否定响应值', '')
                passed = (actual.strip() == exp_str)
                result['结果'] = '通过' if passed else '失败'
                result['期望来源'] = '期望HEX（严格匹配）'

        else:
            # 无期望值 = 只要求肯定响应
            actual = result.get('肯定响应值', '')
            is_positive = bool(actual) and not actual.strip().startswith('7F')
            result['结果'] = '通过' if is_positive else '失败（期望肯定响应）'
            result['期望来源'] = '无期望值（要求肯定响应）'

    # ★★★ 新增：分段保存Excel方法 ★★★
    def save_to_excel_segmented(self, base_filename: str) -> str:
        """
        分段保存Excel文件
        - 每达到 max_rows 行数，保存一个分段文件
        - 最后合并所有分段
        """
        max_rows = FILE_SEGMENT_CONFIG.get('excel_max_rows', 1000)
        
        if not self.messages:
            return base_filename

        # 检查是否需要分段
        total_rows = len(self.messages)
        if total_rows <= max_rows or not FILE_SEGMENT_CONFIG.get('enable_segment_write', True):
            # 数据量小，直接保存
            self._save_single_excel(base_filename)
            return base_filename

        # 分段保存
        segment_count = (total_rows // max_rows) + 1
        base_name = os.path.splitext(base_filename)[0]
        
        print(f"  数据共 {total_rows} 行，将保存为 {segment_count} 个分段文件")

        for seg_idx in range(segment_count):
            start_idx = seg_idx * max_rows
            end_idx = min((seg_idx + 1) * max_rows, total_rows)
            segment_data = self.messages[start_idx:end_idx]

            segment_file = f"{base_name}_part{seg_idx + 1}.xlsx"
            self._save_segment_excel(segment_file, segment_data, start_idx)
            self.segment_files.append(segment_file)
            print(f"  ✓ 分段 {seg_idx + 1}/{segment_count} 已保存: {segment_file}")

        # 合并所有分段到最终文件
        self._merge_segment_files(base_filename)
        return base_filename

    def _save_segment_excel(self, filename: str, messages: List[Dict], offset: int):
        """保存单个分段Excel"""
        df = pd.DataFrame(messages)
        df['测试用例ID'] = df['测试用例ID'].astype(str)
        df['测试项'] = df['测试项'].astype(str)
        
        # 添加结果列
        result_map = {(tid, tname): res for (tid, tname), res in self.test_results.items()}
        for col in ['结果', '肯定响应值', '否定响应值', '22服务内容(hex)',
                    '22服务内容(ascii)', '数据库获取值', '期望来源']:
            df[col] = ''
        
        for idx, row in df.iterrows():
            key = (row['测试用例ID'], row['测试项'])
            if key in result_map:
                res = result_map[key]
                for col in res:
                    df.at[idx, col] = res[col]

        columns_order = [
            '时间', '类型', 'CAN ID (十六进制)', '扩展帧', '数据长度',
            '数据 (十六进制)', '数据 (ASCII)', '测试用例ID', '测试项',
            '结果', '肯定响应值', '否定响应值', '22服务内容(hex)',
            '22服务内容(ascii)', '数据库获取值', '期望来源'
        ]
        existing_cols = [col for col in columns_order if col in df.columns]
        df = df[existing_cols]

        os.makedirs(os.path.dirname(filename), exist_ok=True)
        
        # ★★★ 使用 with 块确保文件关闭 ★★★
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=f'结果{offset//FILE_SEGMENT_CONFIG["excel_max_rows"]+1}', index=False)
            
            if self.service28_monitor_logs:
                df_monitor = pd.DataFrame(self.service28_monitor_logs)
                df_monitor.to_excel(writer, sheet_name='28服务监控', index=False)

    def _merge_segment_files(self, final_filename: str):
        """合并所有分段文件"""
        if not self.segment_files:
            return

        all_data = []
        for seg_file in self.segment_files:
            try:
                df = pd.read_excel(seg_file, sheet_name=0)
                all_data.append(df)
            except Exception as e:
                print(f"  ⚠ 读取分段文件失败: {seg_file}, {e}")

        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            
            with pd.ExcelWriter(final_filename, engine='openpyxl') as writer:
                merged_df.to_excel(writer, sheet_name='测试结果', index=False)
                
                if self.service28_monitor_logs:
                    df_monitor = pd.DataFrame(self.service28_monitor_logs)
                    df_monitor.to_excel(writer, sheet_name='28服务监控', index=False)

            print(f"  ✓ 合并完成: {final_filename}")

            # 清理分段文件
            for seg_file in self.segment_files:
                try:
                    os.remove(seg_file)
                except:
                    pass
            self.segment_files.clear()

    def _save_single_excel(self, filename: str):
        """保存单个Excel文件（不分段）"""
        if not self.messages:
            return
        df = pd.DataFrame(self.messages)
        df['测试用例ID'] = df['测试用例ID'].astype(str)
        df['测试项'] = df['测试项'].astype(str)
        result_map = {(tid, tname): res for (tid, tname), res in self.test_results.items()}
        for col in ['结果', '肯定响应值', '否定响应值', '22服务内容(hex)',
                    '22服务内容(ascii)', '数据库获取值', '期望来源']:
            df[col] = ''
        for idx, row in df.iterrows():
            key = (row['测试用例ID'], row['测试项'])
            if key in result_map:
                res = result_map[key]
                for col in res:
                    df.at[idx, col] = res[col]
        columns_order = [
            '时间', '类型', 'CAN ID (十六进制)', '扩展帧', '数据长度',
            '数据 (十六进制)', '数据 (ASCII)', '测试用例ID', '测试项',
            '结果', '肯定响应值', '否定响应值', '22服务内容(hex)',
            '22服务内容(ascii)', '数据库获取值', '期望来源'
        ]
        existing_cols = [col for col in columns_order if col in df.columns]
        df = df[existing_cols]
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='测试结果', index=False)

            # ★ 新增：28服务监控日志 Sheet
            if self.service28_monitor_logs:
                df_monitor = pd.DataFrame(self.service28_monitor_logs)
                df_monitor.to_excel(writer, sheet_name='28服务监控', index=False)


# ══════════════════════════════════════════════════════════════
# 第2层：类层
# ══════════════════════════════════════════════════════════════

class UDSServiceHandler:
    """UDS服务处理器基类"""

    def __init__(self, config: Dict):
        self.config = config

    def get_expanded_configs(self, cfg: Dict) -> List[Dict]:
        """默认不展开"""
        return [cfg]

    def send(self, bus, arb_id: int, cfg: Dict, logger: 'CanMessageLogger'):
        """发送诊断请求"""
        raw_data = parse_hex_input(cfg['request_data_str'], is_request=True, config=self.config)
        if not raw_data:
            return

        frames = raw_data if isinstance(raw_data, list) and len(raw_data) > 0 and isinstance(raw_data[0], list) else [
            raw_data]

        if len(frames) > 1:
            self._send_multi_frame(bus, logger, arb_id, frames)
        else:
            msg_data = frames[0]
            msg = can.Message(
                arbitration_id=arb_id, data=msg_data,
                is_extended_id=self.config['is_extended_id'],
                is_fd=self.config['can_mode'] > 0
            )
            with BUS_SEND_LOCK:
                bus.send(msg)
            logger.log_sent_message(msg, time.time())

    def _send_multi_frame(self, bus, logger: 'CanMessageLogger', arb_id: int, frames: List[List[int]]):
        """发送多帧（带流控）"""
        is_fd = self.config['can_mode'] > 0

        first_frame = frames[0]
        msg = can.Message(
            arbitration_id=arb_id, data=first_frame,
            is_extended_id=self.config['is_extended_id'], is_fd=is_fd
        )
        with BUS_SEND_LOCK:
            bus.send(msg)
        logger.log_sent_message(msg, time.time())

        fc_received = False
        start_wait = time.time()
        fc_response_id = self.config.get('security_response_id', 0x7BE)

        while time.time() - start_wait < self.config.get('fc_timeout', 1.0):
            resp = bus.recv(timeout=0.05)
            if resp and resp.arbitration_id == fc_response_id:
                data = list(resp.data)
                logger.log_received_message(resp, time.time())
                if data and (data[0] & 0xF0) == 0x30:
                    fc_received = True
                    time.sleep(0.02)
                    break

        if not fc_received:
            return

        for cf in frames[1:]:
            msg = can.Message(
                arbitration_id=arb_id, data=cf,
                is_extended_id=self.config['is_extended_id'], is_fd=is_fd
            )
            with BUS_SEND_LOCK:
                bus.send(msg)
            logger.log_sent_message(msg, time.time())
            time.sleep(self.config['multi_frame_gap'])

    def receive(self, bus, cfg: Dict, logger: 'CanMessageLogger'):
        """接收响应"""
        start_time = time.time()
        logger.current_response_frames.clear()
        response_id = self.config.get('security_response_id', 0x7BE)

        while time.time() - start_time < cfg['response_timeout']:
            msg = bus.recv(timeout=0.1)
            if msg and logger._should_log(msg.arbitration_id):
                logger.log_received_message(msg, time.time())
                # 只有响应ID的帧才重置计时器（支持多帧响应）
                if msg.arbitration_id == response_id:
                    start_time = time.time()

                    # 处理多帧流控
                    if list(msg.data)[0] & 0xF0 == 0x10:
                        fc_data = pad_data(self.config['flow_control_data'], self.config)
                        fc_msg = can.Message(
                            arbitration_id=self.config['security_request_id'],
                            data=fc_data,
                            is_extended_id=self.config['is_extended_id'],
                            is_fd=self.config['can_mode'] > 0
                        )
                        with BUS_SEND_LOCK:
                            bus.send(fc_msg)
                        logger.log_sent_message(fc_msg, time.time())

    def analyze_and_judge(self, cfg: Dict, logger: 'CanMessageLogger'):
        """分析响应并判断结果"""
        response_id = self.config.get('security_response_id', 0x7BE)
        logger.finalize_and_analyze_response(response_id=response_id)
        logger.apply_comparison(cfg)

    def execute_test(self, bus, arb_id: int, cfg: Dict, logger: 'CanMessageLogger'):
        """完整测试流程"""
        self.send(bus, arb_id, cfg, logger)
        self.receive(bus, cfg, logger)
        self.analyze_and_judge(cfg, logger)


class Service22Handler(UDSServiceHandler):
    """
    22服务：读取DID数据
    ★ 展开规则由顶部配置区 SERVICE22_EXPANSION_STEPS 控制
    """

    def get_expanded_configs(self, cfg: Dict) -> List[Dict]:
        """22服务展开规则：使用顶部配置区的 SERVICE22_EXPANSION_STEPS"""
        test_id = cfg['test_case_id']
        test_name = cfg['test_name']

        # 从原始配置提取 DID
        request_hex = cfg['request_data_str']  # 如 "22F190"
        did_hex = request_hex[2:] if request_hex.startswith('22') else request_hex  # "F190"

        # 构建22服务的期望前缀（SID + DID）
        did_prefix = f"62 {did_hex[:2]} {did_hex[2:]}" if len(did_hex) >= 4 else f"62 {did_hex}"

        expanded = []

        # ★ 使用顶部配置区的展开步骤
        for step_idx, step in enumerate(SERVICE22_EXPANSION_STEPS, 1):
            # 替换 {DID} 占位符
            request = step['request'].replace('{DID}', did_hex)
            expected = step['expected'].replace('{DID}', f"{did_hex[:2]} {did_hex[2:]}" if len(did_hex) >= 4 else did_hex)
            wait = step.get('wait', cfg.get('wait_after_request', 0.2))  # ★★★ 使用步骤配置的等待时间

            expanded.append({
                **cfg,
                'test_name': f"{test_name} - {step['name']}",
                'request_data_str': request,
                'expected_hex_str': expected,
                'wait_after_request': wait,  # ★★★ 独立等待时间
                '_comparison_mode': 'prefix',
                '_step_order': step_idx,
            })

        return expanded


class Service28Handler(UDSServiceHandler):
    """
    28服务：通信控制
    - 响应只判断肯定/否定
    - 监控数据记录到独立Sheet
    """

    def execute_test(self, bus, arb_id: int, cfg: Dict, logger: 'CanMessageLogger'):
        monitor_id = cfg.get('monitor_can_id')
        monitor_duration = cfg.get('monitor_duration', 1.0)
        response_id = self.config.get('security_response_id', 0x7BE)

        # 发送28服务请求
        self.send(bus, arb_id, cfg, logger)

        # ── 收集响应，判断肯定/否定 ──
        is_positive = False  # 是否收到肯定响应
        is_negative = False  # 是否收到否定响应
        response_detail = ''  # 响应详情

        start_time = time.time()
        while time.time() - start_time < 0.5:
            msg = bus.recv(timeout=0.02)
            if msg:
                logger.log_received_message(msg, time.time())
                if msg.arbitration_id == response_id:
                    data = list(msg.data)
                    if len(data) >= 2:
                        if data[1] == 0x68:
                            # 肯定响应 68 xx
                            is_positive = True
                            response_detail = f"68 {data[2]:02X}" if len(data) >= 3 else "68"
                        elif data[1] == 0x7F:
                            # 否定响应 7F 28 xx
                            is_negative = True
                            response_detail = f"7F {data[2]:02X} {data[3]:02X}" if len(data) >= 4 else "7F"

        # ── 监控指定ID（如果有） ──
        frame_count = 0
        if monitor_id is not None and monitor_duration > 0:
            monitor_start = time.time()
            while time.time() - monitor_start < monitor_duration:
                msg = bus.recv(timeout=0.02)
                if msg:
                    # ★ 记录到28服务监控日志
                    logger.log_service28_monitor(
                        cfg['test_case_id'], cfg['test_name'],
                        monitor_id, msg, time.time()
                    )
                    if msg.arbitration_id == monitor_id:
                        frame_count += 1

        # ── 判断结果 ──
        expected_min = cfg.get('expected_min_frames', 0)
        expected_max = cfg.get('expected_max_frames', float('inf'))
        frame_passed = expected_min <= frame_count <= expected_max if monitor_id else True
        passed = is_positive and frame_passed

        key = (cfg['test_case_id'], cfg['test_name'])
        logger.test_results[key] = {
            '肯定响应值': response_detail if is_positive else '',
            '否定响应值': response_detail if is_negative else '',
            '22服务内容(hex)': '',
            '22服务内容(ascii)': '',
            '数据库获取值': '',
            '结果': '通过' if passed else '失败',
            '期望来源': f'28响应(肯定)+监控{expected_min}-{expected_max}帧' if monitor_id else '28响应(肯定)'
        }
        logger.execution_order.append(key)

class Service10Handler(UDSServiceHandler):
    """10服务：会话控制（不展开）"""
    pass


class Service27Handler(UDSServiceHandler):
    """
    27服务：安全访问
    ★ 流程参数由构造时传入的 service27_config 控制
    ★ Excel触发值：KEY → SERVICE27_CONFIG, KEY1 → SERVICE27_CONFIG_1
    """

    def __init__(self, config: Dict, service27_config: Dict = None):
        super().__init__(config)
        self.service27_config = service27_config or SERVICE27_CONFIG

    def get_expanded_configs(self, cfg: Dict) -> List[Dict]:
        return [cfg]

    def execute_test(self, bus, arb_id: int, cfg: Dict, logger: CanMessageLogger):
        test_id = cfg['test_case_id']
        test_name = cfg['test_name']
        response_id = self.config.get('security_response_id', 0x7BE)
        svc_cfg = self.service27_config

        logger.current_test_case_id = test_id
        logger.current_test_name = test_name

        print(f"\n  [安全访问] 开始执行...")

        seed = None
        key = None
        response_str = ""
        success = False

        # ★ 使用构造时传入的配置参数
        session_req = svc_cfg['session_request']  # 如 "1003" 或 "1002"
        session_sid = int(session_req[:2], 16) + 0x40  # 10 → 50
        session_sub = int(session_req[2:4], 16)  # 03 或 02

        # ═══ 步骤0：切换会话 ═══
        print(f"  [27服务] 发送 {session_req}")
        self._send(bus, arb_id, [0x02, int(session_req[:2], 16), int(session_req[2:4], 16)], logger)
        time.sleep(0.2)

        # 等待会话响应
        if self._wait_positive(bus, logger, response_id, session_sid, session_sub):
            time.sleep(0.1)

            # ═══ 步骤1：请求种子 ═══
            seed_req = svc_cfg['seed_request']  # 如 "2701" 或 "2705"
            print(f"  [27服务] 发送 {seed_req}")
            self._send(bus, arb_id, [0x02, int(seed_req[:2], 16), int(seed_req[2:4], 16)], logger)
            time.sleep(0.2)

            # ═══ 步骤2：获取种子 ═══
            seed = self._wait_seed(bus, logger, response_id,
                                   svc_cfg['seed_expected_sid'],
                                   svc_cfg['seed_expected_sub'])

            if seed:
                print(f"  [27服务] 收到种子: {seed.hex().upper()}")

                # ═══ 步骤3：计算密钥 ═══
                key = VW_Seed2Key(seed)
                print(f"  [27服务] 计算密钥: {key.hex().upper()}")

                # ═══ 步骤4：发送密钥 ═══
                key_req = svc_cfg['key_request']  # 如 "2702" 或 "2706"
                print(f"  [27服务] 发送 {key_req} {key.hex().upper()}")
                self._send(bus, arb_id, [0x06, int(key_req[:2], 16), int(key_req[2:4], 16)] + list(key) + [0xCC], logger)
                time.sleep(0.2)

                # ═══ 步骤5：等待结果 ═══
                success, response_str = self._wait_key_result(bus, logger, response_id,
                                                              svc_cfg['key_expected_sid'],
                                                              svc_cfg['key_expected_sub'])
            else:
                response_str = "获取种子失败"
        else:
            response_str = f"{session_req} 切换失败"

        # ═══ 记录结果 ═══
        result_key = (test_id, test_name)
        detail = ""
        if seed:
            detail += f"种子:{seed.hex().upper()}"
        if key:
            detail += f" 密钥:{key.hex().upper()}"

        logger.test_results[result_key] = {
            '肯定响应值': response_str if success else '',
            '否定响应值': response_str if not success else '',
            '22服务内容(hex)': detail,
            '结果': '通过' if success else '失败',
            '期望来源': f'期望 {svc_cfg["key_expected_sid"]:02X} {svc_cfg["key_expected_sub"]:02X}（安全访问成功）'
        }
        logger.execution_order.append(result_key)

        print(f"  [安全访问] {'✓ 成功' if success else '✗ 失败'}")

    def _send(self, bus, arb_id: int, data: list, logger):
        """发送单帧"""
        padded = pad_data(data, self.config)
        msg = can.Message(
            arbitration_id=arb_id, data=padded,
            is_extended_id=self.config['is_extended_id'],
            is_fd=self.config['can_mode'] > 0
        )
        with BUS_SEND_LOCK:
            bus.send(msg)
        logger.log_sent_message(msg, time.time())

    def _wait_positive(self, bus, logger, resp_id, sid, sub) -> bool:
        """等待肯定响应"""
        start = time.time()
        while time.time() - start < 1.0:
            msg = bus.recv(timeout=0.1)
            if msg:
                logger.log_received_message(msg, time.time())
                if msg.arbitration_id == resp_id:
                    data = list(msg.data)
                    if len(data) >= 3 and data[1] == sid and data[2] == sub:
                        return True
                    elif data[1] == 0x7F:
                        return False
        return False

    def _wait_seed(self, bus, logger, resp_id, expected_sid, expected_sub) -> Optional[bytes]:
        """等待种子"""
        start = time.time()
        while time.time() - start < 2.0:
            msg = bus.recv(timeout=0.1)
            if msg:
                logger.log_received_message(msg, time.time())
                if msg.arbitration_id == resp_id:
                    data = list(msg.data)
                    if len(data) >= 6 and data[1] == expected_sid and data[2] == expected_sub:
                        return bytes(data[3:7])
                    elif data[1] == 0x7F:
                        return None
        return None

    def _wait_key_result(self, bus, logger, resp_id, expected_sid, expected_sub) -> Tuple[bool, str]:
        """等待密钥验证结果"""
        start = time.time()
        while time.time() - start < 2.0:
            msg = bus.recv(timeout=0.1)
            if msg:
                logger.log_received_message(msg, time.time())
                if msg.arbitration_id == resp_id:
                    data = list(msg.data)
                    if len(data) >= 3 and data[1] == expected_sid and data[2] == expected_sub:
                        return True, f"{expected_sid:02X} {expected_sub:02X}"
                    elif data[1] == 0x7F and data[2] == 0x27:
                        nrc = data[3] if len(data) >= 4 else 0
                        return False, f"7F 27 {nrc:02X}"
        return False, "超时"


class ServiceBA27Handler(UDSServiceHandler):
    """
    BA安全访问（非大众27）
    ★ 流程参数由顶部配置区 SERVICE_BA27_CONFIG 控制
    ★ Excel触发值：KEYBB
    """

    def get_expanded_configs(self, cfg: Dict) -> List[Dict]:
        return [cfg]

    def execute_test(self, bus, arb_id: int, cfg: Dict, logger: CanMessageLogger):
        test_id = cfg['test_case_id']
        test_name = cfg['test_name']
        response_id = self.config.get('security_response_id', 0x719)

        logger.current_test_case_id = test_id
        logger.current_test_name = test_name

        print(f"\n  [BA安全访问] 开始执行...")

        seed = None
        key = None
        response_str = ""
        success = False

        # ★ 使用顶部配置区的参数
        seed_req = SERVICE_BA27_CONFIG['seed_request']  # 如 "BA01"
        seed_sid = int(seed_req[:2], 16)  # BA
        seed_sub = int(seed_req[2:4], 16)  # 01

        # ═══ 步骤1：请求种子 ═══
        print(f"  [BA服务] 发送 {seed_req}")
        self._send(bus, arb_id, [0x02, seed_sid, seed_sub], logger)
        time.sleep(0.2)

        # ═══ 步骤2：等待种子 ═══
        seed = self._wait_seed(bus, logger, response_id,
                               SERVICE_BA27_CONFIG['seed_expected_sid'],
                               SERVICE_BA27_CONFIG['seed_expected_sub'])

        if seed:
            print(f"  [BA服务] 收到种子: {seed.hex().upper()}")

            # ═══ 步骤3：计算密钥 ═══
            key_bytes = calculate_key_level1(list(seed))
            if key_bytes:
                key = bytes(key_bytes)
                print(f"  [BA服务] 计算密钥: {key.hex().upper()}")

                # ═══ 步骤4：发送密钥 ═══
                key_req = SERVICE_BA27_CONFIG['key_request']  # 如 "BA02"
                print(f"  [BA服务] 发送 {key_req} {key.hex().upper()}")
                self._send(bus, arb_id, [0x06, int(key_req[:2], 16), int(key_req[2:4], 16)] + key_bytes, logger)
                time.sleep(0.2)

                success = True
                response_str = "安全访问成功"
            else:
                response_str = "密钥计算失败"
        else:
            response_str = "获取种子失败"

        # ═══ 记录结果 ═══
        result_key = (test_id, test_name)
        detail = ""
        if seed:
            detail += f"种子:{seed.hex().upper()}"
        if key:
            detail += f" 密钥:{key.hex().upper()}"

        logger.test_results[result_key] = {
            '肯定响应值': response_str if success else '',
            '否定响应值': response_str if not success else '',
            '22服务内容(hex)': detail,
            '结果': '通过' if success else '失败',
            '期望来源': 'BA安全访问成功'
        }
        logger.execution_order.append(result_key)

        print(f"  [BA安全访问] {'✓ 成功' if success else '✗ 失败'}")
        time.sleep(0.5)

    def _send(self, bus, arb_id: int, data: list, logger):
        padded = pad_data(data, self.config)
        msg = can.Message(
            arbitration_id=arb_id, data=padded,
            is_extended_id=self.config['is_extended_id'],
            is_fd=self.config['can_mode'] > 0
        )
        with BUS_SEND_LOCK:
            bus.send(msg)
        logger.log_sent_message(msg, time.time())

    def _wait_seed(self, bus, logger, resp_id, expected_sid, expected_sub) -> Optional[bytes]:
        start = time.time()
        while time.time() - start < 2.0:
            msg = bus.recv(timeout=0.1)
            if msg:
                logger.log_received_message(msg, time.time())
                if msg.arbitration_id == resp_id:
                    data = list(msg.data)
                    if len(data) >= 6 and data[1] == expected_sid and data[2] == expected_sub:
                        return bytes(data[3:7])
        return None


class Service31Handler(UDSServiceHandler):
    """31服务：例程控制（不展开）"""
    pass


class RelayHandler(UDSServiceHandler):
    """继电器控制"""
    pass


class PeriodicHandler(UDSServiceHandler):
    """周期发送"""
    pass


class TesterPresentHandler(UDSServiceHandler):
    """3E00 维持会话（Tester Present）
    
    ★ 功能：
      - TP3E：激活状态 + 立即发送首帧 3E00 + 记录 T1
      - STP3E：停止激活
    
    ★ 周期：
      - 独立函数 _send_3e00() 被复用（首帧 + 后续周期续帧）
      - 主循环每轮诊断后 / 分片等待中 调用 check_and_send() 检查 T2-T1≥3s
    
    ★ 约束：
      - 绝不在 execute_test() 内部（诊断请求-响应中间）发送 3E00
      - 仅在 execute_test() 后 / 等待间隙中发送
    """
    _active = False
    _last_send_time = 0.0

    def execute_test(self, bus: 'can.interface.Bus', arb_id: int,
                     cfg: Dict, logger: 'CanLogger'):
        """状态控制：TP3E 启动 + 首帧 / STP3E 停止"""
        request_hex = cfg.get('request_data_str', '').upper().strip()
        if request_hex == 'TP3E':
            TesterPresentHandler._active = True
            TesterPresentHandler._last_send_time = time.time()
            self._send_3e00(bus, logger)
        elif request_hex == 'STP3E':
            TesterPresentHandler._active = False

    def check_and_send(self, bus: 'can.interface.Bus', logger: 'CanLogger'):
        """主循环调用：检查 T2-T1≥3s 则发送续帧"""
        if not TesterPresentHandler._active:
            return
        interval = self.config.get('tp3e_interval', 3.0)
        if time.time() - TesterPresentHandler._last_send_time < interval:
            return
        self._send_3e00(bus, logger)
        TesterPresentHandler._last_send_time = time.time()

    def _send_3e00(self, bus: 'can.interface.Bus', logger: 'CanLogger'):
        """发送 3E00 + 获取响应 7E00

        ★ 发送复用基类 send() 方法（自动处理单帧/多帧/CAN模式）
        ★ 接收独立实现（不调用基类 receive，避免清空诊断响应帧）
        """
        try:
            config = self.config

            # 构造临时 cfg，复用基类 send 发送 3E 00
            
            fake_cfg = {'request_data_str': '3E 00'}
            print(fake_cfg)
            self.send(bus, config.get('tp3e_arb_id', 0x711), fake_cfg, logger)

            # ★★★ 等待并获取 7E00 响应（只接收 security_response_id 的帧）★★★
            response_id = config.get('security_response_id', 0x719)
            deadline = time.time() + config.get('response_timeout', 2.0)
            response = None
            while time.time() < deadline:
                msg = bus.recv(timeout=0.1)
                if msg and msg.arbitration_id == response_id:
                    response = msg
                    break
            if response:
                print(f"[3E00] 接收 {response}")
                logger.log_received_message(response, time.time())

            time.sleep(config.get('tp3e_wait_after', 0.1))
        except Exception as e:
            print(f"[3E00] 发送失败: {e}")


# ══════════════════════════════════════════════════════════════
# 第3层：调度层
# ══════════════════════════════════════════════════════════════

def get_handler(cfg: Dict, config: Dict) -> UDSServiceHandler:
    """根据配置选择对应的 Handler"""
    request_hex = cfg.get('request_data_str', '').upper().strip()
    # ★★★ 新增：包含BB05 → BA安全访问 ★★★
    if 'KEYBB' in request_hex:
        return ServiceBA27Handler(config)
    # ★★★ 触发值"KEY1" → VW安全访问（2705/2706） ★★★
    request_hex = cfg.get('request_data_str', '').upper().strip()
    if request_hex == 'KEY1':
        return Service27Handler(config, service27_config=SERVICE27_CONFIG_1)
    # ★★★ 触发值"KEY" → 安全访问（2701/2702） ★★★
    request_hex = cfg.get('request_data_str', '').upper().strip()
    if request_hex == 'KEY':
        return Service27Handler(config, service27_config=SERVICE27_CONFIG)

    if cfg.get('is_relay_command'):
        return RelayHandler(config)
    if cfg.get('is_periodic') or cfg.get('should_stop_periodic'):
        return PeriodicHandler(config)

    if request_hex.startswith('10'):
        return Service10Handler(config)
    elif request_hex.startswith('22'):
        return Service22Handler(config)
    elif request_hex.startswith('28'):
        return Service28Handler(config)
    # elif request_hex.startswith('27'):
    #     return Service27Handler(config)
    elif request_hex.startswith('31'):
        return Service31Handler(config)
    elif request_hex == 'TP3E' or request_hex == 'STP3E':
        return TesterPresentHandler(config)
    else:
        return UDSServiceHandler(config)


def load_send_configs_from_excel(path: str) -> List[Dict]:
    """从Excel加载原始配置"""
    if not os.path.exists(path):
        return []

    df = pd.read_excel(path)
    configs = []

    for idx, row in df.iterrows():
        if idx == 0 or row.isna().all():
            continue

        enable_str = str(row.get('是否启用', '')).strip().lower()
        is_enabled = enable_str not in ['0', 'false', 'no', '禁用', '跳过', '关', '关闭', '否']

        if not is_enabled:
            continue

        cfg: Dict = {}
        cfg['test_case_id'] = str(row.get('测试用例ID', f"case_{idx}"))
        cfg['test_name'] = row.get('测试项', f"测试项{idx}")

        can_id_val = row.get('CANID')
        if pd.isna(can_id_val):
            continue
        try:
            cfg['arbitration_id'] = int(str(can_id_val), 0)
        except ValueError:
            continue

        cfg['request_data_str'] = str(row.get('请求数据', '')).strip().upper().replace(' ', '')
        cfg['expected_db_id'] = int(row['期望DBID']) if pd.notna(row.get('期望DBID')) else None
        cfg['expected_hex_str'] = str(row.get('期望HEX', '')) if pd.notna(row.get('期望HEX')) else None
        cfg['response_timeout'] = float(row.get('响应超时时间', DEFAULT_CAN_CONFIG['response_timeout']))
        cfg['wait_after_request'] = float(row.get('等待间隔时间', DEFAULT_CAN_CONFIG['wait_after_request']))
        print(cfg['request_data_str'] ,cfg['wait_after_request'] )

        periodic_str = str(row.get('是否周期发送', '')).strip().lower()
        if periodic_str in ['是', 'true', '1', 'yes', '启用', 'start', '开启']:
            cfg['is_periodic'] = True
            cfg['should_stop_periodic'] = False
        elif periodic_str in ['否', '0', 'false', 'no', '禁用', 'stop', '暂停', '停止']:
            cfg['is_periodic'] = False
            cfg['should_stop_periodic'] = True
        else:
            cfg['is_periodic'] = False
            cfg['should_stop_periodic'] = False

        if pd.notna(row.get('周期时间(秒)')):
            try:
                interval = float(row.get('周期时间(秒)'))
                cfg['periodic_interval'] = max(0.02, interval)
            except:
                cfg['periodic_interval'] = cfg['wait_after_request']
        else:
            cfg['periodic_interval'] = cfg['wait_after_request']

        cfg['periodic_can_mode'] = None
        periodic_mode_raw = row.get('周期CAN模式')
        if pd.notna(periodic_mode_raw):
            mode_str = str(periodic_mode_raw).strip()
            try:
                mode_val = int(float(mode_str))
                if mode_val in (0, 1, 2):
                    cfg['periodic_can_mode'] = mode_val
            except:
                pass

        cfg['is_relay_command'] = cfg['request_data_str'] in ['KL30ON', 'KL30OFF', 'KL15ON', 'KL15OFF']
        cfg['relay_command'] = {
            'KL30ON': 'KL30 on', 'KL30OFF': 'KL30 off',
            'KL15ON': 'KL15 on', 'KL15OFF': 'KL15 off',
        }.get(cfg['request_data_str'])
        cfg['monitor_can_id'] = None
        if pd.notna(row.get('监控CANID')):
            try:
                cfg['monitor_can_id'] = int(str(row.get('监控CANID')), 0)
                print(f"  [配置] {cfg['test_name']} 监控ID: 0x{cfg['monitor_can_id']:X}")
            except ValueError:
                print(f"  [警告] 监控CANID格式错误: {row.get('监控CANID')}")
                pass

        cfg['monitor_duration'] = float(row.get('监控时长(秒)', 1.0)) if pd.notna(row.get('监控时长(秒)')) else 1.0
        cfg['expected_min_frames'] = 0
        cfg['expected_max_frames'] = float('inf')

        if pd.notna(row.get('期望报文数')):
            val = str(row.get('期望报文数')).strip()
            if val.startswith('>'):
                cfg['expected_min_frames'] = int(val[1:]) + 1
            elif val.startswith('<'):
                cfg['expected_max_frames'] = int(val[1:]) - 1
            else:
                try:
                    num = int(val)
                    cfg['expected_min_frames'] = num
                    cfg['expected_max_frames'] = num
                except ValueError:
                    pass

        # 调试输出
        if cfg['monitor_can_id']:
            print(f"  [配置] 监控时长: {cfg['monitor_duration']}s")
            print(f"  [配置] 期望报文: {cfg['expected_min_frames']}-{cfg['expected_max_frames']}")

        configs.append(cfg)

    return configs


def load_and_expand_configs(excel_path: str, config: Dict) -> List[Dict]:
    """加载Excel配置并展开"""
    raw_configs = load_send_configs_from_excel(excel_path)

    all_expanded = []
    for cfg in raw_configs:
        handler = get_handler(cfg, config)
        expanded = handler.get_expanded_configs(cfg)
        all_expanded.extend(expanded)

    return all_expanded


# ══════════════════════════════════════════════════════════════
# 第4层：异常退出保存机制 ★★★ 新增
# ══════════════════════════════════════════════════════════════

# 全局变量，用于紧急保存
_global_logger: Optional[CanMessageLogger] = None
_global_listener_thread: Optional[threading.Thread] = None
_global_listener_stop: Optional[threading.Event] = None
_global_output_dir: str = OUTPUT_DIR

def emergency_save(signum=None, frame=None):
    """
    紧急保存函数：在异常退出时保存已收集的数据
    """
    print("\n⚠ 检测到异常退出信号，正在紧急保存数据...")

    if _global_logger is not None:
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_emergency')
            excel_file = os.path.join(_global_output_dir, f"CAN测试结果_{timestamp}.xlsx")
            _global_logger.save_to_excel_segmented(excel_file)
            print(f"✓ 紧急保存Excel成功: {excel_file}")
        except Exception as e:
            print(f"✗ 紧急保存Excel失败: {e}")

    if _global_listener_stop is not None:
        _global_listener_stop.set()

    if _global_listener_thread is not None and _global_listener_thread.is_alive():
        _global_listener_thread.join(timeout=2.0)

    print("✓ 紧急保存完成")
    sys.exit(0)

def register_signal_handlers():
    """
    注册信号处理器
    """
    # Ctrl+C 信号
    signal.signal(signal.SIGINT, emergency_save)
    # 终止信号
    signal.signal(signal.SIGTERM, emergency_save)
    # 程序退出时注册
    atexit.register(emergency_save)


# ══════════════════════════════════════════════════════════════
# 第5层：执行层
# ══════════════════════════════════════════════════════════════

def send_and_receive_can_messages(send_configs: List[Dict],
                                  base_config: Dict = DEFAULT_CAN_CONFIG,
                                  round_num: int = 0,
                                  logger: CanMessageLogger = None):
    config = {**DEFAULT_CAN_CONFIG, **base_config}
    can_logger = logger if logger is not None else CanMessageLogger(config['allowed_ids'])
    can_logger.allowed_ids.update(config['allowed_ids'])
    relay_ser = None
    bus = None
    periodic_tasks: Dict[int, Dict] = {}
    listener_stop = threading.Event()
    listener_thread: Optional[threading.Thread] = None
    
    try:
        relay_ser = init_serial()
        bus_kwargs = {
            'bustype': config['bustype'],
            'channel': config['channel'],
            'bitrate': config['bitrate'],
            'tres': config.get('tres', True),
        }
        if config['can_mode'] > 0:
            bus_kwargs['is_fd'] = True
            bus_kwargs['data_bitrate'] = config['data_bitrate']
        bus = can.interface.Bus(**bus_kwargs)
        
        # ★★★ 启动分段写入的监听线程 ★★★
        listener_thread = threading.Thread(
            target=channel1_listener_segmented,
            args=(listener_stop, FILE_SEGMENT_CONFIG['log_max_size_mb']),
            daemon=True
        )
        listener_thread.start()
        
        # ★★★ 设置全局变量用于紧急保存 ★★★
        global _global_logger, _global_listener_thread, _global_listener_stop
        _global_logger = can_logger
        _global_listener_thread = listener_thread
        _global_listener_stop = listener_stop

        # ★★★ 创建 TP3E Handler 单例 ★★★
        tp3e_handler = TesterPresentHandler(config)

        total = len(send_configs)
        for idx, cfg in enumerate(send_configs, 1):
            print(f"[{idx}/{total}] {cfg['test_name']} (ID: {cfg['test_case_id']})")
            can_logger.current_test_case_id = cfg['test_case_id']
            can_logger.current_test_name = cfg['test_name']
            arb_id = cfg['arbitration_id']
            handler = get_handler(cfg, config)

            # 继电器命令
            if cfg.get('is_relay_command'):
                relay_cmd = cfg.get('relay_command')
                if relay_cmd and relay_ser:
                    execute_relay_command(relay_ser, relay_cmd, delay=0.6)
            # 停止周期发送
            elif cfg.get('should_stop_periodic', False) and arb_id in periodic_tasks:
                periodic_tasks[arb_id]['stop_event'].set()
                periodic_tasks[arb_id]['thread'].join(timeout=2.0)
                del periodic_tasks[arb_id]
                if arb_id in can_logger.allowed_ids:
                    can_logger.allowed_ids.remove(arb_id)
            # 周期发送
            elif cfg.get('is_periodic', False):
                hex_str = cfg['request_data_str'].replace(' ', '')
                try:
                    raw_bytes = [int(hex_str[i:i + 2], 16) for i in range(0, len(hex_str), 2)]
                except ValueError:
                    pass  # 异常不 continue，走统一等待
                else:
                    periodic_config = config.copy()
                    if cfg.get('periodic_can_mode') is not None:
                        periodic_config['can_mode'] = cfg['periodic_can_mode']
                    max_len = 8 if periodic_config['can_mode'] in (0, 1) else 64
                    if len(raw_bytes) > max_len:
                        raw_bytes = raw_bytes[:max_len]
                    msg_data = pad_data(raw_bytes, periodic_config)
                    if arb_id in periodic_tasks:
                        periodic_tasks[arb_id]['stop_event'].set()
                        periodic_tasks[arb_id]['thread'].join(timeout=2.0)
                        del periodic_tasks[arb_id]
                    stop_event = threading.Event()
                    t = threading.Thread(
                        target=periodic_send,
                        args=(bus, arb_id, msg_data, cfg['periodic_interval'],
                              stop_event, periodic_config, can_logger),
                        daemon=True
                    )
                    t.start()
                    periodic_tasks[arb_id] = {'thread': t, 'stop_event': stop_event}
                    can_logger.allowed_ids.add(arb_id)
                    is_fd = periodic_config['can_mode'] > 0
                    init_msg = can.Message(
                        arbitration_id=arb_id, data=msg_data,
                        is_extended_id=periodic_config['is_extended_id'], is_fd=is_fd
                    )
                    with BUS_SEND_LOCK:
                        bus.send(init_msg)
                    can_logger.log_sent_message(init_msg, time.time())
            # 普通诊断 / TP3E / STP3E
            else:
                handler.execute_test(bus, arb_id, cfg, can_logger)

            # ★★★ 统一检查 TP3E（execute_test 后 / 等待间隙） ★★★
            tp3e_handler.check_and_send(bus, can_logger)

            # ★★★ 统一分片等待：每0.5s检查一次是否到点发3E00 ★★★
            _w = cfg['wait_after_request']
            print(_w)
            while _w > 0:
                time.sleep(min(0.5, _w))
                _w -= 0.5
                tp3e_handler.check_and_send(bus, can_logger)
    except Exception as e:
        print(f"发生错误：{e}")
        import traceback
        traceback.print_exc()
    finally:
        for task in periodic_tasks.values():
            task['stop_event'].set()
            task['thread'].join(timeout=2.0)
        listener_stop.set()
        if listener_thread and listener_thread.is_alive():
            listener_thread.join(timeout=3.0)
        if bus:
            try:
                bus.shutdown()
            except:
                pass
        close_serial(relay_ser)


# ══════════════════════════════════════════════════════════════
# 入口
# ══════════════════════════════════════════════════════════════
def main():
    # ★★★ 注册信号处理器 ★★★
    register_signal_handlers()
    
    configs = load_and_expand_configs(EXCEL_PLAN_PATH, DEFAULT_CAN_CONFIG)
    if not configs:
        print("无有效测试用例，程序退出")
        return

    total_rounds = max(1, LOOP_COUNT)
    print(f"共加载 {len(configs)} 条测试用例（展开后）")
    print(f"共执行 {total_rounds} 轮\n")

    # ★ 一个 logger 跑所有轮次
    can_logger = CanMessageLogger(DEFAULT_CAN_CONFIG['allowed_ids'])

    for round_num in range(1, total_rounds + 1):
        if total_rounds > 1:
            print(f"\n{'='*60}")
            print(f"  第 {round_num}/{total_rounds} 轮")
            print(f"{'='*60}\n")

        can_logger._current_round = round_num  # ★ 标记当前轮次

        round_configs = load_and_expand_configs(EXCEL_PLAN_PATH, DEFAULT_CAN_CONFIG)
        send_and_receive_can_messages(round_configs, round_num=round_num, logger=can_logger)

        if round_num < total_rounds:
            print(f"\n  等待 {LOOP_GAP}s 后开始下一轮...")
            time.sleep(LOOP_GAP)

    # ★ 所有轮次跑完，统一保存
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_file = os.path.join(OUTPUT_DIR, f"CAN测试结果_{timestamp}.xlsx")
    can_logger.save_to_excel_segmented(excel_file)
    print(f"\n全部 {total_rounds} 轮执行完毕，结果已保存到：{excel_file}")


if __name__ == "__main__":
    main()
