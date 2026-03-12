import pyvisa

# 获取IP地址
ip_address = input("请输入仪器IP地址: ")

# 获取要发送的指令
command = input("请输入要发送的指令: ")

# 初始化资源管理器
rm = pyvisa.ResourceManager()

# 打开资源
inst = rm.open_resource(f'TCPIP0::{ip_address}::INSTR')

# 设置终止符
inst.read_termination = '\n'
inst.write_termination = '\n'

# 发送指令
inst.write(command)

# 读取并打印返回结果
response = inst.read()
print(f"仪器返回结果: {response}")

# 关闭资源
inst.close()