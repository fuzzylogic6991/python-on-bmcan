# CAN报文发送接收工具 - 完整说明

## 📦 文件说明

本目录包含以下文件：

1. **can_send_receive_to_excel.py** - 完整功能脚本（推荐）
   - 支持发送多个CAN报文
   - 自动接收ECU响应
   - 导出到Excel（包含多个工作表）

2. **simple_can_test.py** - 简单测试脚本（适合新手）
   - 发送单个CAN报文
   - 接收响应
   - 导出到Excel

3. **使用说明-CAN发送接收工具.md** - 详细使用说明
4. **快速入门.md** - 快速入门指南
5. **requirements.txt** - 依赖库列表

## 🎯 使用流程

### 1. 安装依赖

```bash
pip install pandas openpyxl -i https://mirrors.aliyun.com/pypi/simple/
```

### 2. 选择脚本

- **新手**：使用 `simple_can_test.py`
- **需要完整功能**：使用 `can_send_receive_to_excel.py`

### 3. 配置参数

修改脚本中的CAN ID、数据、通道号、波特率等参数。

### 4. 运行脚本

```bash
python simple_can_test.py
# 或
python can_send_receive_to_excel.py
```

### 5. 查看结果

打开生成的Excel文件查看发送和接收的数据。

## 📚 详细文档

- **快速入门**：查看 `快速入门.md`
- **详细说明**：查看 `使用说明-CAN发送接收工具.md`

## 🔧 常见问题

### Q: 提示找不到pandas模块？
A: 运行 `pip install pandas openpyxl`

### Q: CAN总线打开失败？
A: 检查BUSMASTER是否已启动，BMAPI64.dll是否在正确位置

### Q: 没有收到响应？
A: 检查ECU是否正常工作，CAN ID是否正确，增加超时时间

### Q: 如何发送多个报文？
A: 使用 `can_send_receive_to_excel.py`，在 `send_configs` 列表中添加多个配置

## 💡 配置示例

### 标准帧示例
```python
CAN_ID = 0x7DF
CAN_DATA = [0x02, 0x10, 0x03]
is_extended_id = False
```

### 扩展帧示例
```python
CAN_ID = 0xC0FFEE
CAN_DATA = [0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08]
is_extended_id = True
```

## 📞 技术支持

如果遇到问题：
1. 检查BUSMASTER配置
2. 检查CAN硬件连接
3. 查看终端错误信息
4. 参考详细使用说明文档

