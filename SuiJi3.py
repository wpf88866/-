import pandas as pd
import numpy as np
import random
import os
from time import sleep


def grade_to_score(grade):
    """
    将等级转换为分数
    A -> 95
    B -> 85
    C -> 75
    D -> 65
    E -> 55
    """
    grade_map = {
        'A': 95,
        'B': 85,
        'C': 75,
        'D': 65,
        'E': 55
    }
    return grade_map.get(str(grade).strip().upper(), 0)


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
        file_path = r"C:\Users\wpf\Desktop\机电随机成绩\3班.xlsx"

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

        print("\n文件读取成功！")
        print("Excel文件的列名：", df.columns.tolist())
        total_rows = len(df)
        print(f"总共需要处理 {total_rows} 行数据")

        # 创建新的列
        new_columns = ['随机1', '随机2', '随机3', '随机4', '随机5', '验证平均分']
        for col in new_columns:
            df[col] = 0

        # 处理每一行数据
        for index, row in df.iterrows():
            try:
                # 将等级转换为分数
                grades = [row.iloc[i] for i in range(2, 7)]  # 获取C到G列的等级
                scores = [grade_to_score(grade) for grade in grades]

                # 计算原始平均分
                original_score = sum(scores) / len(scores)
                target_avg = original_score  # 使用原始平均分作为目标平均分

                # 调整分数
                new_scores = adjust_scores(original_score, target_avg, scores)

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
                print(f"当前行数据: {row.tolist()}")
                continue

        # 保存结果
        output_path = os.path.join(os.path.dirname(file_path), "机电随机成绩_3班_已调整.xlsx")
        df.to_excel(output_path, index=False)
        print(f"\n处理完成！")
        print(f"结果已保存至：{output_path}")

        # 验证结果
        print("\n验证结果：")
        for index, row in df.iterrows():
            grades = [row.iloc[i] for i in range(2, 7)]
            original_scores = [grade_to_score(grade) for grade in grades]
            original_avg = sum(original_scores) / len(original_scores)
            new_scores = [row[f'随机{i + 1}'] for i in range(5)]
            new_avg = sum(new_scores) / 5

            if abs(new_avg - original_avg) > 0.1:
                print(f"\n第 {index + 1} 行存在较大偏差：")
                print(f"原始等级：{grades}")
                print(f"原始分数：{original_scores}")
                print(f"原始平均分：{original_avg}")
                print(f"调整后分数：{new_scores}")
                print(f"调整后平均分：{new_avg}")

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
