#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CAN报文发送和接收工具 - 完整版（严格ISO-TP流控）
满足全部9项需求 + 多帧发送必须等待流控帧
V7: 移除trim_padding_bytes以避免删除有效数据；添加can_mode开关（0=Classic CAN dlc=8, 1=CAN FD dlc max=8, 2=CAN FD dlc max=64）
"""

import os
import subprocess
import can
import time
from datetime import datetime
import pandas as pd
from typing import List, Dict, Optional, Union
import sys  # 新增：用于处理命令行参数

# ========================== 默认配置 ==========================
DEFAULT_CAN_CONFIG = {
    'bustype': 'bmcan',
    'channel': 0,
    'bitrate': 500000,
    'data_bitrate': 2000000,
    'can_mode': 2,  # 0: Classic CAN (dlc=8), 1: CAN FD (dlc max=8), 2: CAN FD (dlc max=64)
    'is_extended_id': False,
    'dlc': 8,  # 默认dlc，会根据can_mode动态调整
    'wait_response': True,
    'response_timeout': 2.0,
    'allowed_ids': {0x711, 0x719},
    'wait_after_request': 2.0,
    'flow_control_data': [0x30, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],  # 默认FC
    'multi_frame_gap': 0.02,
    'fc_timeout': 1.0,  # 等待流控帧超时时间（秒）
    'security_request_id': 0x711,
    'security_response_id': 0x719,
    'security_timeout': 2.0,
    'keep_alive_id': 0x711,
    'keep_alive_data': [0x02, 0x3E, 0x80, 0x00, 0x00, 0x00, 0x00, 0x00],
    'keep_alive_interval': 3.0,
    'keep_alive_enabled_by_default': False,
}

EXCEL_PLAN_PATH = r"E:\SRwork_DQJBF\301.script\python-can-(BUSMUST)\python-can-4.0.0\examples\input\can_input_模板 - 副本.xlsx"
OUTPUT_DIR = r"E:\SRwork_DQJBF\301.script\python-can-(BUSMUST)\python-can-4.0.0\examples\output"

# ========================== CAN FD DLC 映射表 ==========================
FD_DLC_MAP = {
    0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8,
    9: 12, 10: 16, 11: 20, 12: 24, 13: 32, 14: 48, 15: 64
}

def get_fd_dlc(length: int) -> int:
    """根据数据长度获取CAN FD DLC值"""
    for dlc, size in sorted(FD_DLC_MAP.items(), key=lambda x: x[1]):
        if length <= size:
            return dlc
    raise ValueError(f"Data length {length} exceeds maximum 64 bytes for CAN FD")

def get_fd_padded_length(dlc: int) -> int:
    """根据DLC获取填充后的数据长度"""
    return FD_DLC_MAP.get(dlc, 0)

# ========================== 工具函数 ==========================
def pad_data(data: List[int], config: Dict) -> List[int]:
    """根据can_mode填充数据"""
    can_mode = config['can_mode']
    dlc_max = 8 if can_mode in (0, 1) else 64
    padded = list(data)
    current_len = len(padded)
    if current_len > dlc_max:
        raise ValueError(f"Data length {current_len} exceeds max {dlc_max} for can_mode {can_mode}")

    if can_mode == 0:
        # Classic CAN: pad to 8
        if current_len < 8:
            padded += [0x00] * (8 - current_len)
        return padded[:8]
    elif can_mode == 1:
        # CAN FD max 8: pad to nearest <=8, but dlc=len
        return padded  # No pad, dlc = current_len <=8
    else:  # can_mode == 2
        # CAN FD full: find dlc, pad to corresponding size
        dlc = get_fd_dlc(current_len)
        padded_len = get_fd_padded_length(dlc)
        if current_len < padded_len:
            padded += [0x00] * (padded_len - current_len)
        return padded

def parse_hex_input(hex_str: str, is_request: bool = True, config: Dict = DEFAULT_CAN_CONFIG) -> Union[List[int], List[List[int]]]:
    """解析用户输入HEX，请求自动补length+padding，期望HEX不补"""
    hex_str = hex_str.strip().upper().replace(' ', '')
    if not hex_str:
        return []

    bytes_list = [int(hex_str[i:i+2], 16) for i in range(0, len(hex_str), 2)]

    if not is_request:
        return bytes_list

    can_mode = config['can_mode']
    dlc_max = 8 if can_mode in (0, 1) else 64
    ff_prefix_len = 2 if can_mode in (0, 1) else 2  # FF: 0x1X len_low (12位len，高4位在0x10)
    cf_prefix_len = 1
    ff_data_len = dlc_max - ff_prefix_len
    cf_data_len = dlc_max - cf_prefix_len

    total_len = len(bytes_list)
    if total_len <= (dlc_max - 1):  # 单帧 (SF: PCI(1) + data <= dlc_max)
        return pad_data([total_len] + bytes_list, config)
    else:  # 多帧
        frames: List[List[int]] = []
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
            seq = (seq + 1) % 16  # 循环seq
        return frames

def bytes_to_ascii(data: List[int]) -> str:
    return ''.join(chr(b) if 32 <= b <= 126 else '.' for b in data)

def extract_isotp_payload(frames: List[List[int]]) -> List[int]:
    if not frames:
        return []

    first = frames[0]
    pci = first[0] & 0xF0
    payload: List[int] = []

    if pci == 0x00:  # 单帧
        length = first[0] & 0x0F
        payload = first[1:1 + length]
    elif pci == 0x10:  # 多帧
        total_len = ((first[0] & 0x0F) << 8) | first[1]
        payload.extend(first[2:])
        for cf in frames[1:]:
            if (cf[0] & 0xF0) == 0x20:
                payload.extend(cf[1:])
        payload = payload[:total_len]
    else:
        for fr in frames:
            payload.extend(fr)

    return payload  # 移除trim_padding_bytes，直接返回按length截取的payload

# ========================== ADB查询 ==========================
def execute_adb_command(command: str) -> Optional[str]:
    try:
        result = subprocess.run(['adb', 'shell', command], capture_output=True, text=True, check=True)
        return result.stdout.strip()
    except Exception:
        return None

def query_tbox_data_by_id(param_id: int) -> Optional[str]:
    cmd = f"cd /oemdata/parameters/ && sqlite3 paramsDb 'select * from paramstbl where id={param_id};'"
    result = execute_adb_command(cmd)
    if result:
        parts = result.split('|')
        return parts[2] if len(parts) >= 3 else None
    return None

# ========================== 输入配置加载 ==========================
def load_send_configs_from_excel(path: str) -> List[Dict]:
    if not os.path.exists(path):
        print(f"输入Excel不存在: {path}")
        return []

    df = pd.read_excel(path)
    configs = []
    print(f"加载输入Excel，共 {len(df)} 行（第1行为示例，已跳过）")

    for idx, row in df.iterrows():
        if idx == 0:  # 示例行跳过
            continue
        if row.isna().all():
            continue

        enable_val = row.get('是否启用', '是')
        if pd.isna(enable_val):
            enable_val = '是'
        if str(enable_val).strip().lower() in ['0', 'false', 'no', '禁用', '跳过']:
            print(f"跳过禁用行: {row.get('测试项', '未知')}")
            continue

        cfg: Dict = {}
        cfg['test_case_id'] = str(row.get('测试用例ID', f"case_{idx}"))
        cfg['test_name'] = row.get('测试项', f"测试项{idx}")
        cfg['arbitration_id'] = int(str(row['CANID']), 0)
        cfg['request_data_str'] = str(row.get('请求数据', '')).strip()
        cfg['expected_db_id'] = int(row['期望DBID']) if pd.notna(row.get('期望DBID')) else None
        cfg['expected_hex_str'] = str(row.get('期望HEX', '')).strip() if pd.notna(row.get('期望HEX')) else None
        cfg['response_timeout'] = float(row.get('响应超时时间', DEFAULT_CAN_CONFIG['response_timeout']))
        cfg['wait_after_request'] = float(row.get('等待间隔时间', DEFAULT_CAN_CONFIG['wait_after_request']))
        cfg['enable_keep_alive'] = str(row.get('是否发送保持活跃', '否')).strip().lower() in ['是', 'true', '1', 'yes', '启用', '全局']

        configs.append(cfg)

    print(f"有效测试用例: {len(configs)} 条")
    return configs

# ========================== 日志记录器 ==========================
class CanMessageLogger:
    def __init__(self, allowed_ids: Optional[set] = None):
        self.messages: List[Dict] = []
        self.test_results: Dict[str, Dict] = {}
        self.allowed_ids = allowed_ids or set()
        self.current_test_case_id: Optional[str] = None
        self.current_test_name: Optional[str] = None
        self.current_response_frames: List[List[int]] = []

    def _should_log(self, arbitration_id: int) -> bool:
        return not self.allowed_ids or arbitration_id in self.allowed_ids

    def log_sent_message(self, msg: can.Message, timestamp: float):
        if not self._should_log(msg.arbitration_id):
            return
        data_list = list(msg.data)
        self.messages.append({
            '时间': datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
            '类型': '发送',
            'CAN ID (十六进制)': f"0x{msg.arbitration_id:X}",
            '扩展帧': '是' if msg.is_extended_id else '否',
            '数据长度': msg.dlc,
            '数据 (十六进制)': ' '.join(f"{b:02X}" for b in data_list),
            '数据 (ASCII)': bytes_to_ascii(data_list),
            '测试用例ID': self.current_test_case_id or '',
            '测试项': self.current_test_name or '',
            '结果': '',
            '肯定响应值': '',
            '否定响应值': '',
            '22服务内容(hex)': '',
            '22服务内容(ascii)': '',
            '数据库获取值': '',
            '期望来源': '',
        })

    def log_received_message(self, msg: can.Message, timestamp: float):
        if not self._should_log(msg.arbitration_id):
            return
        self.current_response_frames.append(list(msg.data))
        data_list = list(msg.data)
        self.messages.append({
            '时间': datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
            '类型': '接收',
            'CAN ID (十六进制)': f"0x{msg.arbitration_id:X}",
            '扩展帧': '是' if msg.is_extended_id else '否',
            '数据长度': msg.dlc,
            '数据 (十六进制)': ' '.join(f"{b:02X}" for b in data_list),
            '数据 (ASCII)': bytes_to_ascii(data_list),
            '测试用例ID': self.current_test_case_id or '',
            '测试项': self.current_test_name or '',
            '结果': '',
            '肯定响应值': '',
            '否定响应值': '',
            '22服务内容(hex)': '',
            '22服务内容(ascii)': '',
            '数据库获取值': '',
            '期望来源': '',
        })

    def finalize_and_analyze_response(self):
        if not self.current_response_frames or not self.current_test_case_id:
            self.test_results[self.current_test_case_id] = {'结果': '超时无响应'}
            return

        payload = extract_isotp_payload(self.current_response_frames)
        if not payload:
            self.test_results[self.current_test_case_id] = {'结果': '空响应'}
            return

        result_dict = {
            '肯定响应值': '',
            '否定响应值': '',
            '22服务内容(hex)': '',
            '22服务内容(ascii)': '',
            '结果': '通过',  # 默认，后面比对可覆盖
            '数据库获取值': '',
            '期望来源': '',
        }

        sid = payload[0]
        if sid == 0x7F and len(payload) >= 3:  # 否定响应
            result_dict['否定响应值'] = f"7F {payload[1]:02X} {payload[2]:02X}"
            result_dict['结果'] = '否定响应'
        else:
            if sid == 0x62:  # 22服务正响应
                if len(payload) >= 3:
                    data_part = payload[3:]
                    result_dict['肯定响应值'] = f"62 {payload[1]:02X} {payload[2]:02X}"
                    result_dict['22服务内容(hex)'] = ' '.join(f"{b:02X}" for b in data_part)
                    result_dict['22服务内容(ascii)'] = bytes_to_ascii(data_part)
            else:
                result_dict['肯定响应值'] = ' '.join(f"{b:02X}" for b in payload)

        self.test_results[self.current_test_case_id] = result_dict
        self.current_response_frames.clear()

    def apply_comparison(self, cfg: Dict):
        tid = cfg['test_case_id']
        result = self.test_results.get(tid, {})
        if not result:
            return

        if cfg['expected_db_id'] is not None:
            db_val = query_tbox_data_by_id(cfg['expected_db_id']) or ''
            ascii_val = result.get('22服务内容(ascii)', '')
            passed = (ascii_val == db_val)
            result['结果'] = '通过' if passed else '失败'
            result['数据库获取值'] = db_val
            result['期望来源'] = f"DB id={cfg['expected_db_id']}"

        elif cfg['expected_hex_str']:
            exp_bytes = parse_hex_input(cfg['expected_hex_str'], is_request=False)
            exp_str = ' '.join(f"{b:02X}" for b in exp_bytes).strip()
            actual = result.get('肯定响应值', '') or result.get('否定响应值', '')
            passed = (actual.strip() == exp_str)
            result['结果'] = '通过' if passed else '失败'
            result['期望来源'] = '期望HEX'

    def save_to_excel(self, filename: str):
        if not self.messages:
            print("无数据保存")
            return

        df = pd.DataFrame(self.messages)
        for tid, res in self.test_results.items():
            for col, val in res.items():
                df.loc[df['测试用例ID'] == tid, col] = val

        columns_order = [
            '时间', '类型', 'CAN ID (十六进制)', '扩展帧', '数据长度',
            '数据 (十六进制)', '数据 (ASCII)', '测试用例ID', '测试项',
            '结果', '肯定响应值', '否定响应值', '22服务内容(hex)',
            '22服务内容(ascii)', '数据库获取值', '期望来源'
        ]
        df = df[columns_order]

        os.makedirs(os.path.dirname(filename), exist_ok=True)
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='测试结果', index=False)
        print(f"测试结果已保存: {filename}")

# ========================== 多帧发送（严格等待流控） ==========================
def send_multi_frame_with_flow_control(bus, logger: CanMessageLogger, arbitration_id: int,
                                       frames: List[List[int]], config: Dict):
    """发送多帧请求：发首帧 → 等待FC → 再发CF"""
    if len(frames) <= 1:
        return False

    is_fd = config['can_mode'] > 0

    # 发送首帧
    first_frame = frames[0]
    msg = can.Message(arbitration_id=arbitration_id, data=first_frame,
                      is_extended_id=config['is_extended_id'], is_fd=is_fd)
    ts = time.time()
    bus.send(msg)
    logger.log_sent_message(msg, ts)
    print("→ 已发送首帧，等待流控帧...")

    # 等待流控帧（FC）
    fc_received = False
    start_wait = time.time()
    fc_response_id = config.get('security_response_id', 0x7AD)  # 通常响应ID为请求ID+8

    while time.time() - start_wait < config.get('fc_timeout', 1.0):
        resp = bus.recv(timeout=0.05)
        if resp and resp.arbitration_id == fc_response_id:
            data = list(resp.data)
            logger.log_received_message(resp, time.time())
            if data and (data[0] & 0xF0) == 0x30:  # 流控帧
                print("✓ 收到流控帧，继续发送连续帧")
                fc_received = True
                time.sleep(0.02)  # 标准建议延迟
                break

    if not fc_received:
        print("✗ 超时未收到流控帧，停止发送后续帧")
        return False

    # 发送剩余连续帧
    for cf in frames[1:]:
        msg = can.Message(arbitration_id=arbitration_id, data=cf,
                          is_extended_id=config['is_extended_id'], is_fd=is_fd)
        ts = time.time()
        bus.send(msg)
        logger.log_sent_message(msg, ts)
        time.sleep(config['multi_frame_gap'])

    return True

# ========================== 主流程 ==========================
def send_and_receive_can_messages(send_configs: List[Dict], base_config: Dict = DEFAULT_CAN_CONFIG, shared_logger: Optional[CanMessageLogger] = None):
    config = {**DEFAULT_CAN_CONFIG, **base_config}
    logger = shared_logger if shared_logger else CanMessageLogger(config['allowed_ids'])

    print("=" * 60)
    print("开始执行CAN/UDS测试（严格流控）")
    print("=" * 60)

    try:
        is_fd = config['can_mode'] > 0
        bus_kwargs = {
            'bustype': config['bustype'],
            'channel': config['channel'],
            'bitrate': config['bitrate'],
        }
        if is_fd:
            bus_kwargs['is_fd'] = True
            bus_kwargs['data_bitrate'] = config['data_bitrate']
        bus = can.interface.Bus(**bus_kwargs)
        print(f"CAN总线打开成功: {bus.channel_info}")

        last_activity_time = time.time()
        KEEP_ALIVE_INTERVAL = config['keep_alive_interval']  # 通常 3.0 秒

        def send_keep_alive(bus, logger: CanMessageLogger, config: Dict, is_fd: bool):
            ka_data = pad_data(config['keep_alive_data'], config)
            msg = can.Message(
                arbitration_id=config['keep_alive_id'],
                data=ka_data,
                is_extended_id=config['is_extended_id'],
                is_fd=is_fd
            )
            ts = time.time()
            bus.send(msg)
            logger.log_sent_message(msg, ts)
            print("→ 发送保持活跃 3E 80")
            return ts  # 返回发送时间，用于更新活动时间

        for idx, cfg in enumerate(send_configs, 1):
            print(f"\n[{idx}/{len(send_configs)}] {cfg['test_name']} (ID: {cfg['test_case_id']})")
            logger.current_test_case_id = cfg['test_case_id']
            logger.current_test_name = cfg['test_name']

            # === 在发送请求前检查是否需要发送 keep-alive ===
            current_time = time.time()
            if cfg['enable_keep_alive'] and (current_time - last_activity_time >= KEEP_ALIVE_INTERVAL):
                send_keep_alive(bus, logger, config, is_fd)
                last_activity_time = time.time()   # 更新活动时间

            # 解析请求数据
            raw_data = parse_hex_input(cfg['request_data_str'], is_request=True, config=config)
            frames = raw_data if isinstance(raw_data, list) and len(raw_data) > 0 and isinstance(raw_data[0], list) else [raw_data]

            # 发送请求（单帧或多帧）
            if len(frames) > 1:
                success = send_multi_frame_with_flow_control(bus, logger, cfg['arbitration_id'], frames, config)
                if success:
                    last_activity_time = time.time()  # 多帧发送成功，视为活动
            else:
                msg_data = frames[0]
                msg = can.Message(
                    arbitration_id=cfg['arbitration_id'],
                    data=msg_data,
                    is_extended_id=config['is_extended_id'],
                    is_fd=is_fd
                )
                ts = time.time()
                bus.send(msg)
                logger.log_sent_message(msg, ts)
                last_activity_time = ts  # 单帧发送也算活动

            # 接收响应
            start_time = time.time()
            logger.current_response_frames.clear()
            received_something = False

            while time.time() - start_time < cfg['response_timeout']:
                msg = bus.recv(timeout=0.1)
                if msg and logger._should_log(msg.arbitration_id):
                    logger.log_received_message(msg, time.time())
                    start_time = time.time()  # 刷新超时
                    received_something = True

                    # 响应为多帧首帧时，自动发送流控帧
                    if msg.arbitration_id == config.get('security_response_id', 0x7AD) and list(msg.data)[0] & 0xF0 == 0x10:
                        fc_data = pad_data(config['flow_control_data'], config)
                        fc_msg = can.Message(
                            arbitration_id=config['security_request_id'],
                            data=fc_data,
                            is_extended_id=config['is_extended_id'],
                            is_fd=is_fd
                        )
                        bus.send(fc_msg)
                        logger.log_sent_message(fc_msg, time.time())
                        print("→ 检测到响应首帧，已发送流控帧")
                        last_activity_time = time.time()  # 发送流控也算活动

            if received_something:
                last_activity_time = time.time()  # 只要收到任何响应，刷新活动时间

            # 分析响应并比对
            logger.finalize_and_analyze_response()
            logger.apply_comparison(cfg)

            time.sleep(cfg['wait_after_request'])

        bus.shutdown()

        # 移除保存逻辑，由main调用

    except Exception as e:
        import traceback
        print(f"错误: {e}")
        traceback.print_exc()

    return logger  # 返回logger以便main使用

def main():
    # 新增：处理命令行参数，指定循环次数（默认1次）
    if len(sys.argv) > 1:
        try:
            num_loops = int(sys.argv[1])
            if num_loops < 1:
                raise ValueError
        except ValueError:
            print("无效的循环次数参数，使用默认1次")
            num_loops = 1
    else:
        num_loops = 1  # 默认1次；若想默认100次，可改为100

    print(f"将循环执行测试用例集 {num_loops} 次")

    send_configs = load_send_configs_from_excel(EXCEL_PLAN_PATH)
    if not send_configs:
        print("无有效测试用例，程序退出")
        return

    # 创建共享logger
    shared_logger = CanMessageLogger(DEFAULT_CAN_CONFIG['allowed_ids'])

    for loop_idx in range(1, num_loops + 1):
        print(f"\n=== 开始第 {loop_idx} 次循环执行 ===")
        
        # 为每个循环复制configs，并修改test_case_id以唯一
        loop_configs = []
        for cfg in send_configs:
            new_cfg = cfg.copy()
            new_cfg['test_case_id'] = f"{cfg['test_case_id']}_loop{loop_idx}"
            new_cfg['test_name'] = f"{cfg['test_name']} (Loop {loop_idx})"
            loop_configs.append(new_cfg)
        
        send_and_receive_can_messages(loop_configs, shared_logger=shared_logger)
        print(f"=== 第 {loop_idx} 次循环执行完成 ===")

    # 所有循环后保存到一个Excel
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_file = os.path.join(OUTPUT_DIR, f"CAN测试结果_全循环_{timestamp}.xlsx")
    shared_logger.save_to_excel(excel_file)

if __name__ == "__main__":
    main()