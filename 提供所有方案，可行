import pandas as pd
import numpy as np
import random
import os
from time import sleep


def adjust_scores(original_score, target_avg, original_scores=None):
    """
    使用预设的固定方案调整分数，确保调整幅度不超过±2
    """
    original_score = round(original_score)
    target_avg = round(target_avg, 1)

    # 预设的调整方案（调整值之和为0，且单个调整值不超过±2）
    adjustment_patterns = [
        [1, 1, -1, -1, 0],  # 两高两低
        [1, -1, 1, -1, 0],  # 交替波动
        [2, -1, -1, 0, 0],  # 一高两低
        [1, 1, 0, -1, -1],  # 两高两低
        [-1, -1, 1, 1, 0],  # 两低两高
        [1, 0, 0, -1, 0],  # 最小波动
        [0, 1, -1, 0, 0],  # 轻微波动
        [1, -1, 0, 0, 0],  # 简单波动
    ]

    # 随机选择一个方案
    selected_pattern = random.choice(adjustment_patterns)

    # 应用调整方案
    new_scores = [original_score + adj for adj in selected_pattern]

    # 验证分数是否在有效范围内且调整幅度不超过±2
    if (all(0 <= score <= 100 for score in new_scores) and
            all(abs(score - original_score) <= 2 for score in new_scores)):
        actual_avg = round(sum(new_scores) / 5, 1)
        if abs(actual_avg - target_avg) < 0.1:  # 允许0.1的误差
            return new_scores

    # 如果调整后的分数无效，返回原始分数或接近目标的分数
    if original_scores and len(original_scores) == 5:
        return original_scores

    # 如果没有原始分数，则返回原始分数的5个副本
    return [original_score] * 5


def main():
    try:
        # 文件路径
        file_path = r"C:\Users\wpf\Desktop\机电随机成绩\2班.xlsx"

        # 验证文件是否存在
        if not os.path.exists(file_path):
            print(f"错误：找不到文件 {file_path}")
            print("请检查文件路径是否正确")
            return

        print(f"正在读取文件：{file_path}")

        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"读取Excel文件时出错：{str(e)}")
            print("请确保文件格式正确且未被其他程序占用")
            return

        # 验证数据格式
        if len(df.columns) < 8:
            print("错误：Excel文件格式不正确")
            print("需要至少8列数据（包括学号、姓名、5个分数列和目标平均分）")
            return

        print("\n文件读取成功！")
        print("Excel文件的列名：", df.columns.tolist())
        total_rows = len(df)
        print(f"总共需要处理 {total_rows} 行数据")

        if total_rows == 0:
            print("错误：文件中没有数据")
            return

        print("\n开始处理数据...")

        # 创建新的列
        new_columns = ['随机1', '随机2', '随机3', '随机4', '随机5', '验证平均分']
        for col in new_columns:
            df[col] = 0

        # 处理每一行数据
        for index, row in df.iterrows():
            try:
                # 获取原始分数和目标平均分
                original_score = float(row.iloc[2])  # C列的分数
                target_avg = float(row.iloc[7])  # H列的目标平均分

                # 数据验证
                if not (0 <= original_score <= 100 and 0 <= target_avg <= 100):
                    print(f"警告：第 {index + 1} 行的分数数据异常")
                    print(f"原始分数：{original_score}，目标平均分：{target_avg}")
                    continue

                # 获取原始的5个分数
                original_scores = [float(row.iloc[i]) for i in range(2, 7)]

                # 调整分数
                new_scores = adjust_scores(original_score, target_avg, original_scores)

                # 验证调整幅度
                if any(abs(score - original_score) > 2 for score in new_scores):
                    print(f"警告：第 {index + 1} 行的调整超出±2范围")
                    print(f"原始分数：{original_score}")
                    print(f"生成的分数：{new_scores}")

                # 更新DataFrame
                for i, score in enumerate(new_scores):
                    df.at[index, f'随机{i + 1}'] = score

                # 计算并更新验证平均分
                df.at[index, '验证平均分'] = round(sum(new_scores) / 5, 1)

                # 显示进度
                if (index + 1) % 5 == 0 or index == total_rows - 1:
                    print(f"已处理 {index + 1}/{total_rows} 行数据 ({round((index + 1) / total_rows * 100, 1)}%)")

            except Exception as e:
                print(f"处理第 {index + 1} 行时出错: {str(e)}")
                continue

        # 生成输出文件名
        output_dir = os.path.dirname(file_path)
        output_path = os.path.join(output_dir, "机电随机成绩_2班_已调整.xlsx")

        # 尝试保存文件
        try:
            df.to_excel(output_path, index=False)
            print(f"\n处理完成！")
            print(f"结果已保存至：{output_path}")
        except Exception as e:
            print(f"\n保存文件时出错：{str(e)}")
            print("请确保输出文件未被占用且有写入权限")
            return

        # 验证结果
        print("\n验证结果：")
        verification = df.apply(lambda x: abs(x['验证平均分'] - x.iloc[7]) < 0.1, axis=1)
        print(f"所有行的平均值验证：{'全部通过' if verification.all() else '有误差'}")

        # 显示不符合要求的行
        if not verification.all():
            print("\n以下行的平均值与目标值有较大误差：")
            error_rows = df[~verification]
            for idx, row in error_rows.iterrows():
                print(f"第 {idx + 1} 行:")
                print(f"原始分数：{row.iloc[2]}")
                print(f"目标平均分：{row.iloc[7]}")
                print(f"实际平均分：{row['验证平均分']}")
                print(f"生成的分数：{[row[f'随机{i + 1}'] for i in range(5)]}")
                print()

    except Exception as e:
        print(f"程序执行出错：{str(e)}")
        import traceback
        print(traceback.format_exc())

    finally:
        print("\n程序执行结束")


if __name__ == "__main__":
    # 请确认文件路径
    print("程序开始执行...")
    print("请确认Excel文件路径是否正确：")
    file_path = r"C:\Users\wpf\Desktop\机电随机成绩\2班.xlsx"
    print(file_path)
    print()

    main()
