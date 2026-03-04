import numpy as np
import pandas as pd

# 读取CSV文件
file_path = r"C:\PTS\zhongzi\SpectrumSNR\spectrum_curve.csv"

# 使用pandas读取CSV文件（更简单的方式）
data = pd.read_csv(file_path)

# 显示数据的基本信息
print("数据基本信息：")
print(data.head())
print("\n数据形状：", data.shape)

# 找出所有数值列的最大值
print("\n各列的最大值：")
for column in data.columns:
    if data[column].dtype in ['float64', 'int64']:
        max_value = data[column].max()
        print(f"{column} 的最大值: {max_value}")

# 使用numpy找出整个数据集的数值最大值
numeric_data = data.select_dtypes(include=[np.number])
overall_max = numeric_data.max().max()
print(f"\n整个数据集的数值最大值: {overall_max}")

# 找出最大值所在的行
max_row_index = numeric_data.max(axis=1).idxmax()
max_row = data.iloc[max_row_index]
print(f"\n最大值所在的行:")
print(max_row)

