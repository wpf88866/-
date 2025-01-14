import pandas as pd
import numpy as np
import random
import os
from time import sleep


def adjust_scores(original_score, target_avg, max_attempts=100):
    """
    生成随机调整后的分数
    添加最大尝试次数防止无限循环
    """
    for _ in range(max_attempts):
        try:
            # 生成前4个随机偏差值
            adjustments = [random.randint(-5, 5) for _ in range(4)]
            # 计算前4个新分数
            new_scores = [original_score + adj for adj in adjustments]

            # 计算第5个分数（保证平均值）
            last_score = round(target_avg * 5 - sum(new_scores))

            # 检查最后一个分数是否在合理范围内
            if abs(last_score - original_score) <= 5:
                new_scores.append(last_score)
                return new_scores
        except Exception as e:
            print(f"生成分数时出错: {str(e)}")
            continue

    # 如果无法生成合适的分数，返回原始分数
    return [original_score] * 5


try:
    # 读取Excel文件
    file_path = r"C:\Users\wpf\Desktop\机电随机成绩\2班.xlsx"
    df = pd.read_excel(file_path)

    print("Excel文件的列名：", df.columns.tolist())
    total_rows = len(df)
    print(f"总共需要处理 {total_rows} 行数据")
    print("开始处理数据...")

    # 创建新的列
    new_columns = ['随机1', '随机2', '随机3', '随机4', '随机5', '验证平均分']
    for col in new_columns:
        df[col] = 0

    # 处理每一行数据
    for index, row in df.iterrows():
        try:
            # 使用iloc按位置获取数据
            original_score = float(row.iloc[2])  # C列是第3列（索引为2）
            target_avg = float(row.iloc[7])  # H列是第8列（索引为7）

            # 生成新的分数
            new_scores = adjust_scores(original_score, target_avg)

            # 更新DataFrame中的值
            for i, score in enumerate(new_scores):
                df.at[index, f'随机{i + 1}'] = score

            # 显示进度
            if (index + 1) % 5 == 0 or index == total_rows - 1:
                print(f"已处理 {index + 1}/{total_rows} 行数据 ({round((index + 1) / total_rows * 100, 1)}%)")

        except Exception as e:
            print(f"处理第 {index + 1} 行时出错: {str(e)}")
            continue

    # 验证平均值
    df['验证平均分'] = df[['随机1', '随机2', '随机3', '随机4', '随机5']].mean(axis=1)

    # 生成输出文件名
    output_path = os.path.join(os.path.dirname(file_path), "机电随机成绩_2班_已调整.xlsx")

    # 保存结果
    df.to_excel(output_path, index=False)
    print(f"\n处理完成！")
    print(f"结果已保存至：{output_path}")

    # 验证结果
    print("\n验证结果：")
    verification = df.apply(lambda x: abs(x['验证平均分'] - x[6]) < 0.1, axis=1)
    print(f"所有行的平均值验证：{'全部通过' if verification.all() else '有误差'}")

except Exception as e:
    print(f"发生错误：{str(e)}")
    import traceback

    print(traceback.format_exc())

finally:
    print("\n程序执行结束")
