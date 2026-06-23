#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简单CAN测试脚本 - 快速测试发送和接收

这是一个简化版本，适合快速测试CAN通信是否正常
"""

import can
import time
from datetime import datetime
import pandas as pd


def simple_test():
    """简单的发送接收测试"""
    
    print("=" * 50)
    print("简单CAN测试")
    print("=" * 50)
    
    # 配置参数 - 你可以在这里修改
    CAN_ID = 0x7DF          # CAN ID（十六进制）
    CAN_DATA = [0x02, 0x10, 0x03]  # 数据
    CHANNEL = 0             # CAN通道
    BITRATE = 500000        # 波特率
    
    print(f"配置:")
    print(f"  CAN ID: 0x{CAN_ID:X} ({CAN_ID})")
    print(f"  数据: {[hex(b) for b in CAN_DATA]}")
    print(f"  通道: {CHANNEL}")
    print(f"  波特率: {BITRATE}")
    print("=" * 50)
    
    # 记录数据
    records = []
    
    try:
        # 打开CAN总线
        print("\n正在打开CAN总线...")
        bus = can.interface.Bus(
            bustype='bmcan',
            channel=CHANNEL,
            bitrate=BITRATE,
            data_bitrate=2000000,
            tres=True
        )
        print(f"✓ CAN总线已打开")
        
        # 创建并发送报文
        print("\n发送CAN报文...")
        msg = can.Message(
            arbitration_id=CAN_ID,
            data=CAN_DATA,
            is_extended_id=False
        )
        
        send_time = time.time()
        bus.send(msg)
        print(f"✓ 已发送: ID=0x{CAN_ID:X}, 数据={[hex(b) for b in CAN_DATA]}")
        
        # 记录发送的报文
        records.append({
            '时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
            '类型': '发送',
            'CAN ID': f"0x{CAN_ID:X}",
            '数据': ' '.join([f"{b:02X}" for b in CAN_DATA])
        })
        
        # 等待接收响应
        print("\n等待响应（5秒）...")
        timeout = 5.0
        start_time = time.time()
        response_count = 0
        
        while (time.time() - start_time) < timeout:
            received_msg = bus.recv(timeout=0.1)
            if received_msg is not None:
                receive_time = time.time()
                response_count += 1
                print(f"✓ 收到响应 #{response_count}: ID=0x{received_msg.arbitration_id:X}, "
                      f"数据={[hex(b) for b in received_msg.data]}")
                
                # 记录接收的报文
                records.append({
                    '时间': datetime.fromtimestamp(receive_time).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                    '类型': '接收',
                    'CAN ID': f"0x{received_msg.arbitration_id:X}",
                    '数据': ' '.join([f"{b:02X}" for b in received_msg.data])
                })
        
        if response_count == 0:
            print("⚠ 未收到任何响应")
        else:
            print(f"\n✓ 共收到 {response_count} 个响应")
        
        # 关闭总线
        bus.shutdown()
        print("\n✓ CAN总线已关闭")
        
        # 保存到Excel
        if records:
            timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            excel_filename = f"can_test_{timestamp_str}.xlsx"
            
            df = pd.DataFrame(records)
            df.to_excel(excel_filename, index=False, engine='openpyxl')
            print(f"\n✓ 数据已保存到: {excel_filename}")
        else:
            print("\n⚠ 没有数据需要保存")
        
    except Exception as e:
        print(f"\n✗ 发生错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    simple_test()

