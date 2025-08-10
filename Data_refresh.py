import os
from salary_check import get_last_month_str, get_salary_data, get_github_excel, get_github_excel1, merge_data_by_project
import pandas as pd
from datetime import datetime, timedelta

def refresh_df(pat):
    salary_month = get_last_month_str()
    print(f" - 获取工资数据：{salary_month}")
    salary_df = get_salary_data(salary_month)
    # 2. 获取GitHub上的核对人信息（PAT 应该来自安全来源）
    print(" - 获取 GitHub 信息")
    github_df = get_github_excel(pat)
    # 3. 合并 & 发送核对报告
    if not salary_df.empty:
        final_df = merge_data_by_project(salary_df, github_df)
        # 获取当前时间并加 8 小时
        current_time_utc8 = datetime.now() + timedelta(hours=8)

        # 添加到 DataFrame
        final_df['creation_time'] = current_time_utc8
        # 保存时标准化时间格式
        final_df.to_json('output.json', orient='records', date_format='iso')
    else:
        print("❗ 没有有效的工资数据可供处理")

    print("✅ 工资核对流程结束。")
# 使用示例
if __name__ == "__main__":
    github_pat = os.getenv("EXCEL_GITHUB_PAT")

    print("asdd")
    if not github_pat:
        print("asd")
        raise ValueError("请设置 GITHUB_PAT 环境变量")
    refresh_df(github_pat)