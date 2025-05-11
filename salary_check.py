# -*- coding: utf-8 -*-
import os
from email.utils import formataddr
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
github_pat = os.getenv("EXCEL_GITHUB_PAT")
smtp_user = os.getenv("SMTP_USER")
smtp_pass = os.getenv("SMTP_PASS")

def send_salary_reminder(to_email, content_table, subject='工资核对提醒'):
    """
    发送工资核对提醒邮件（支持HTML表格内容）

    参数:
        to_email (str/list): 收件人邮箱，可以是单个字符串或多个邮箱的列表
        content_table (str): HTML表格内容
        subject (str): 邮件主题，默认为'工资核对提醒'
    """
    # region SMTP认证配置
    username = smtp_user
    password = smtp_pass
    # 发件人设置
    From = formataddr(['工资核对提醒', username])
    replyto = 'hr@boyuegf.com'  # 设置回信地址
    #endregion
    # region 处理收件人格式（支持字符串或列表）
    if isinstance(to_email, str):
        to = to_email
    else:
        to = ','.join(to_email)
    #endregion
    # region 构建邮件
    msg = MIMEMultipart('alternative')
    msg['From'] = From
    msg['Reply-to'] = replyto
    msg['To'] = to
    msg['Subject'] = subject
    msg['Message-id'] = email.utils.make_msgid()
    msg['Date'] = email.utils.formatdate()

    # 添加HTML内容
    msg.attach(MIMEText(content_table, _subtype='html', _charset='UTF-8'))
    # endregion
    # region 连接SMTP服务器并发送
    try:
        # 尝试SSL连接
        client = smtplib.SMTP_SSL('smtp.qiye.aliyun.com', 465)
        print('SMTP_SSL连接成功')
    except Exception as e1:
        try:
            # 尝试普通连接
            client = smtplib.SMTP('smtp.qiye.aliyun.com', 25, timeout=5)
            print('SMTP连接成功')
        except Exception as e2:
            print('连接服务器失败:', str(e2))
            return False

    try:
        client.login(username, password)
        print('登录成功')
        client.sendmail(username, [to] if isinstance(to, str) else to.split(','), msg.as_string())
        print('邮件发送成功')
        return True
    except Exception as e:
        print('邮件发送失败:', str(e))
        return False
    finally:
        client.quit()
    # endregion

def get_salary_data(salary_month):
    """
    获取工资表数据（模拟Power Query功能）

    参数:
        salary_month (str): 工资月份，格式如'2023-05'

    返回:
        pd.DataFrame: 处理后的工资表数据
    """
    # 1. 设置请求URL和头部
    url = "http://121.28.192.238:8562/salary-bytx/saUploadPayroll/getExcel"
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36..."
    }

    # 2. 构建请求数据
    post_data = {
        "action": "",
        "header": {
            "requestId": "1",
            "timeStamp": "1",
            "accessToken": "",
            "appCode": "app_01",
            "appSecretKey": "1"
        },
        "body": {
            "saMonth": salary_month,
            "currentUserId": "0301592",
            "isLock": "0,2",
            "powerType": 875,
            "socketId": ""
        }
    }

    try:
        # 3. 发送POST请求
        response = requests.post(
            url,
            headers=headers,
            json=post_data,
            timeout=10
        )
        response.raise_for_status()

        # 4. 读取Excel数据（假设返回的是Excel二进制流）
        from io import BytesIO
        df = pd.read_excel(BytesIO(response.content), sheet_name="Sheet0")

        # 5. 数据清洗（对应Power Query步骤）
        # 重命名列（如果第一行是标题）
        df.columns = df.iloc[0] if 'BG' not in df.columns else df.columns
        df = df.iloc[1:] if 'BG' not in df.columns else df

        # 转换数据类型
        df = df.astype({
            'BG': 'str',
            '部门': 'str',
            '基地': 'str',
            '项目组': 'str',
            '上传人': 'str',
            '终版上传人': 'str',
            '备注': 'str'
        })
        df['工资月份'] = pd.to_datetime(df['工资月份'])
        df['上传时间'] = pd.to_datetime(df['上传时间'])
        df['终版上传时间'] = pd.to_datetime(df['终版上传时间'], errors='coerce')  # 处理可能的空值

        # 6. 筛选行（排除特定项目组）
        filter_condition = (
                ~df['项目组'].str.contains('共享中心|劳务派遣|招聘中台平台', na=False) &
                ~df['基地'].str.contains('总部职能', na=False)
        )
        df = df[filter_condition].reset_index(drop=True)

        return df

    except Exception as e:
        print(f"获取数据失败: {str(e)}")
        return pd.DataFrame()  # 返回空DataFrame

def get_github_excel(github_pat):
    """
    从GitHub仓库获取Excel文件

    参数:
        github_pat (str): GitHub个人访问令牌

    返回:
        pd.DataFrame: 包含项目与核对人关系的DataFrame
    """
    url = "https://api.github.com/repos/BYTX-YGJ/excel/contents/%E5%B7%A5%E8%B5%84%E6%A0%B8%E7%AE%97%E4%BA%BA%E7%BB%9F%E8%AE%A1.xlsx"

    headers = {
        "Authorization": f"token {github_pat}",
        "Accept": "application/vnd.github.v3.raw"
    }

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        # 读取Excel数据
        df = pd.read_excel(BytesIO(response.content), sheet_name="Sheet1")
        return df

    except Exception as e:
        print(f"从GitHub获取数据失败: {str(e)}")
        return pd.DataFrame()

def get_github_excel1(github_pat):
    """
    从GitHub仓库获取Excel文件

    参数:
        github_pat (str): GitHub个人访问令牌

    返回:
        pd.DataFrame: 包含项目与核对人关系的DataFrame
    """
    url = "https://api.github.com/repos/BYTX-YGJ/excel/contents/邮箱维护.xlsx"

    headers = {
        "Authorization": f"token {github_pat}",
        "Accept": "application/vnd.github.v3.raw"
    }

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        # 读取Excel数据
        df = pd.read_excel(BytesIO(response.content), sheet_name="Sheet1")
        return df

    except Exception as e:
        print(f"从GitHub获取数据失败: {str(e)}")
        return pd.DataFrame()

def merge_data_by_project(salary_df, checker_df):
    """
    以项目组为主键合并数据，保留salary_df所有项目，补充checker_df的核对人信息

    参数:
        salary_df (pd.DataFrame): 第一个查询的数据（含项目组，无核对人）
        checker_df (pd.DataFrame): 第二个查询的数据（含项目和核对人）

    返回:
        pd.DataFrame: 合并后的数据
    """
    # 标准化列名（确保都有"项目组"列）
    checker_df = checker_df.rename(columns={'项目': '项目组'})

    # 左连接合并（保留salary_df所有项目组）
    merged_df = pd.merge(
        left=salary_df,
        right=checker_df[['项目组', '工资核对人']],  # 只取需要的列
        on='项目组',
        how='left'
    )

    # 重命名列（保持一致性）
    merged_df = merged_df.rename(columns={'工资核对人': '核对人'})

    return merged_df

def send_complete_salary_report(final_df,github_df1,hours):
    """
    发送完整的工资核对报告，包含：
    - 最近1小时新提交记录（待核对）
    - 历史未核对记录（待核对）
    - 已完成核对记录（有终版上传时间）
    """
    # 检查必要列是否存在
    required_columns = ['BG', '部门', '基地', '项目组', '工资月份',
                        '上传时间', '终版上传时间', '核对人']
    if not all(col in final_df.columns for col in required_columns):
        print("数据中缺少必要列！")
        return False

    # 转换时间列
    final_df['上传时间'] = pd.to_datetime(final_df['上传时间'])
    final_df['终版上传时间'] = pd.to_datetime(final_df['终版上传时间'], errors='coerce')

    # 时间阈值：当前时间前一小时
    time_threshold = datetime.now() - timedelta(hours=hours)

    # 1. 最近1小时内的新提交记录
    recent_records = final_df[(final_df['上传时间'] >= time_threshold)&(final_df['终版上传时间'].isna())].copy()

    # 2. 历史未完成记录（无终版上传时间，且上传时间早于1小时以前）
    pending_records = final_df[
        (final_df['终版上传时间'].isna()) &
        (final_df['上传时间'] < time_threshold)
    ].copy()

    # 3. 已完成记录（有终版上传时间）
    completed_records = final_df[final_df['终版上传时间'].notna()].copy()

    # 没有待核对的记录则不发送
    if recent_records.empty and pending_records.empty:
        print("没有需要核对的工资记录")
        return False

    # 为每类记录添加状态标签（仅在非空时设置）
    if not recent_records.empty:
        recent_records.loc[:, '状态'] = '待核对（新提交）'
    if not pending_records.empty:
        pending_records.loc[:, '状态'] = '待核对（历史未完成）'
    if not completed_records.empty:
        completed_records.loc[:, '状态'] = '已完成'

    # 合并所有记录
    all_records = pd.concat([recent_records, pending_records, completed_records], ignore_index=True)

    # 格式化时间列
    time_format = '%Y-%m-%d %H:%M'
    all_records['上传时间'] = pd.to_datetime(all_records['上传时间'], errors='coerce').dt.strftime(time_format)
    all_records['终版上传时间'] = pd.to_datetime(all_records['终版上传时间'], errors='coerce').dt.strftime(time_format)
    all_records['工资月份'] = pd.to_datetime(all_records['工资月份'], errors='coerce').dt.strftime('%Y-%m')

    # 指定需要展示的列
    display_columns = ['BG', '部门', '基地', '项目组', '工资月份',
                       '上传时间', '终版上传时间', '状态', '核对人']
    all_records = all_records[display_columns]

    now = datetime.now()
    time_tolerance = timedelta(minutes=5)
    scheduled_times = [
        now.replace(hour=9, minute=0, second=0, microsecond=0),
        now.replace(hour=13, minute=30, second=0, microsecond=0)
    ]
    is_scheduled_time = any(abs(now - scheduled_time) <= time_tolerance for scheduled_time in scheduled_times)


    # 分组按核对人发送邮件（仅满足条件才发）
    for checker, group in all_records.groupby('核对人'):
        has_new = (group['状态'] == '待核对（新提交）').any()

        if has_new or is_scheduled_time:
            # 从github_df1中查找邮箱
            to_email = github_df1.loc[github_df1['核对人'] == checker, '邮箱'].values
            if len(to_email) == 0:
                print(f"未找到 {checker} 的邮箱地址，跳过发送")
                continue
            to_email = to_email[0]
            html_content = create_status_html(group)
            send_salary_reminder(
                to_email=to_email,
                content_table=html_content,
                subject=f"【您的待核对】{now.strftime('%m-%d')} "
            )
        else:
            print(f"{checker} 无需发送邮件（无新增，非定时）")
    return True

def create_status_html(df):
    """
    生成按状态分组的HTML报告

    参数:
        df (pd.DataFrame): 包含状态标签的完整数据

    返回:
        str: 美化后的HTML内容
    """
    # 按状态分组
    status_groups = {
        '待核对（新提交）': df[df['状态'] == '待核对（新提交）'],
        '待核对（历史未完成）': df[df['状态'] == '待核对（历史未完成）'],
        '已完成': df[df['状态'] == '已完成']
    }

    # 生成每个组的表格
    tables_html = ""
    for status, group_df in status_groups.items():
        if not group_df.empty:
            table_html = group_df.to_html(
                index=False,
                classes='salary-table',
                border=0,
                justify='center',
                na_rep=''
            )
            tables_html += f"""
            <div class="status-section">
                <h3>{status}（共{len(group_df)}条）</h3>
                {table_html}
            </div>
            """

    # 完整HTML结构
    beautiful_html = f"""
    <html>
        <head>
            <style>
                body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; }}
                .container {{ max-width: 1000px; margin: 0 auto; padding: 20px; }}
                .header {{ color: #333; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
                .status-section {{ margin-bottom: 30px; }}
                .status-section h3 {{
                    color: #1e88e5;
                    border-left: 4px solid #1e88e5;
                    padding-left: 10px;
                }}
                .salary-table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 10px 0;
                    font-size: 14px;
                }}
                .salary-table th {{
                    background-color: #f5f5f5;
                    padding: 12px;
                    text-align: center;
                    border-bottom: 2px solid #ddd;
                }}
                .salary-table td {{
                    padding: 10px;
                    border-bottom: 1px solid #eee;
                    text-align: center;
                }}
                .status-pending {{ color: #d32f2f; font-weight: bold; }}
                .status-completed {{ color: #388e3c; }}
                .footer {{
                    margin-top: 20px;
                    color: #777;
                    font-size: 12px;
                    text-align: center;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h2>工资核对进度</h2>
                </div>
                {tables_html}
                <div class="footer">
                    <p>本邮件由系统自动发送，请勿直接回复</p>
                    <p>生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                </div>
            </div>
        </body>
    </html>
    """
    return beautiful_html

def get_last_month_str():
    today = datetime.today().replace(day=1)
    last_month = today - timedelta(days=1)
    return last_month.strftime('%Y-%m')

def run_salary_check_process(pat):
    """
    工资核对流程执行函数：
    1. 获取上月工资数据
    2. 获取 GitHub 上的核对人信息
    3. 合并处理后发送邮件报告
    """
    print("▶ 开始工资核对流程...")

    # 1. 获取原始工资数据（上一个月）
    salary_month = get_last_month_str()
    print(f" - 获取工资数据：{salary_month}")
    salary_df = get_salary_data(salary_month)

    # 2. 获取GitHub上的核对人信息（PAT 应该来自安全来源）
    print(" - 获取 GitHub 信息")
    github_df = get_github_excel(pat)
    github_df1 = get_github_excel1(pat)  # 包含邮箱

    # 3. 合并 & 发送核对报告
    if not salary_df.empty:
        final_df = merge_data_by_project(salary_df, github_df)
        send_complete_salary_report(final_df, github_df1, 1.1)
    else:
        print("❗ 没有有效的工资数据可供处理")

    print("✅ 工资核对流程结束。")

# 使用示例
if __name__ == "__main__":
    github_pat = github_pat  # 推荐从环境变量或安全存储读取
    run_salary_check_process(github_pat)