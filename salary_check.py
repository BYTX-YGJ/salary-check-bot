# -*- coding: utf-8 -*-
import os
import smtplib
import email
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def send_salary_reminder(to_email, content_table, subject='工资核对提醒'):
    """
    发送工资核对提醒邮件（从环境变量读取SMTP配置）
    """
    # 从环境变量读取配置
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASS")
    smtp_server = os.getenv("SMTP_SERVER", "smtp.qiye.aliyun.com")
    smtp_port = os.getenv("SMTP_PORT", "465")
    reply_to = os.getenv("REPLY_TO", "hr@boyuegf.com")

    # 邮件配置
    msg = MIMEMultipart('alternative')
    msg['From'] = formataddr(['工资核对提醒', smtp_user])
    msg['Reply-to'] = reply_to
    msg['To'] = to_email if isinstance(to_email, str) else ','.join(to_email)
    msg['Subject'] = subject
    msg['Message-id'] = email.utils.make_msgid()
    msg['Date'] = email.utils.formatdate()
    msg.attach(MIMEText(content_table, _subtype='html', _charset='UTF-8'))

    # 发送邮件
    try:
        client = smtplib.SMTP_SSL(smtp_server, int(smtp_port))
        client.login(smtp_user, smtp_pass)
        client.sendmail(smtp_user, to_email.split(',') if isinstance(to_email, str) else to_email, msg.as_string())
        logger.info(f"邮件发送成功 -> {to_email}")
        return True
    except Exception as e:
        logger.error(f"邮件发送失败: {str(e)}")
        return False
    finally:
        client.quit()


def get_salary_data(salary_month):
    """
    从API获取工资数据（带错误重试机制）
    """
    url = "http://121.28.192.238:8562/salary-bytx/saUploadPayroll/getExcel"
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    payload = {
        "action": "",
        "header": {"requestId": "1", "timeStamp": "1", "accessToken": "", "appCode": "app_01", "appSecretKey": "1"},
        "body": {
            "saMonth": salary_month,
            "currentUserId": "0301592",
            "isLock": "0,2",
            "powerType": 875,
            "socketId": ""
        }
    }

    try:
        logger.info(f"正在获取 {salary_month} 工资数据...")
        response = requests.post(url, headers=headers, json=payload, timeout=15)
        response.raise_for_status()

        df = pd.read_excel(BytesIO(response.content), sheet_name="Sheet0")
        df.columns = df.iloc[0] if 'BG' not in df.columns else df.columns
        df = df.iloc[1:] if 'BG' not in df.columns else df

        # 数据清洗
        df = df.astype({
            'BG': 'str', '部门': 'str', '基地': 'str',
            '项目组': 'str', '上传人': 'str', '终版上传人': 'str', '备注': 'str'
        })
        df['工资月份'] = pd.to_datetime(df['工资月份'])
        df['上传时间'] = pd.to_datetime(df['上传时间'])
        df['终版上传时间'] = pd.to_datetime(df['终版上传时间'], errors='coerce')

        # 筛选有效数据
        df = df[
            (~df['项目组'].str.contains('共享中心|劳务派遣|招聘中台平台', na=False)) &
            (~df['基地'].str.contains('总部职能', na=False))
            ].reset_index(drop=True)

        logger.info(f"成功获取 {len(df)} 条工资记录")
        return df

    except Exception as e:
        logger.error(f"获取工资数据失败: {str(e)}")
        return pd.DataFrame()


def get_github_data(repo_path, sheet_name="Sheet1"):
    """从GitHub仓库获取Excel数据（通用方法）"""
    github_pat = os.getenv("GITHUB_PAT")
    url = f"https://api.github.com/repos/{repo_path}"
    headers = {
        "Authorization": f"token {github_pat}",
        "Accept": "application/vnd.github.v3.raw"
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content), sheet_name=sheet_name)
    except Exception as e:
        logger.error(f"从GitHub获取数据失败: {str(e)}")
        return pd.DataFrame()


def send_complete_salary_report(final_df, github_df1, hours=1.1):
    """发送完整的工资核对报告"""
    try:
        # 时间计算
        now = datetime.now()
        time_threshold = now - timedelta(hours=hours)

        # 数据分类
        recent_records = final_df[(final_df['上传时间'] >= time_threshold) & (final_df['终版上传时间'].isna())].copy()
        pending_records = final_df[(final_df['终版上传时间'].isna()) & (final_df['上传时间'] < time_threshold)].copy()
        completed_records = final_df[final_df['终版上传时间'].notna()].copy()

        # 添加状态标签
        if not recent_records.empty:
            recent_records.loc[:, '状态'] = '待核对（新提交）'
        if not pending_records.empty:
            pending_records.loc[:, '状态'] = '待核对（历史未完成）'
        if not completed_records.empty:
            completed_records.loc[:, '状态'] = '已完成'

        # 合并记录
        all_records = pd.concat([recent_records, pending_records, completed_records], ignore_index=True)

        # 格式化时间
        time_format = '%Y-%m-%d %H:%M'
        all_records['上传时间'] = pd.to_datetime(all_records['上传时间']).dt.strftime(time_format)
        all_records['终版上传时间'] = pd.to_datetime(all_records['终版上传时间'], errors='coerce').dt.strftime(
            time_format)
        all_records['工资月份'] = pd.to_datetime(all_records['工资月份']).dt.strftime('%Y-%m')

        # 检查是否有待核对记录
        if recent_records.empty and pending_records.empty:
            logger.info("没有需要核对的工资记录")
            return False

        # 分组发送邮件
        for checker, group in all_records.groupby('核对人'):
            to_email = github_df1.loc[github_df1['核对人'] == checker, '邮箱'].values
            if len(to_email) == 0:
                logger.warning(f"未找到 {checker} 的邮箱地址")
                continue

            html_content = create_status_html(group)
            send_salary_reminder(
                to_email=to_email[0],
                content_table=html_content,
                subject=f"【您的待核对】{now.strftime('%m-%d')}"
            )
        return True

    except Exception as e:
        logger.error(f"发送报告时出错: {str(e)}")
        return False


def create_status_html(df):
    """生成带样式的HTML报告"""
    status_groups = {
        '待核对（新提交）': df[df['状态'] == '待核对（新提交）'],
        '待核对（历史未完成）': df[df['状态'] == '待核对（历史未完成）'],
        '已完成': df[df['状态'] == '已完成']
    }

    tables_html = ""
    for status, group_df in status_groups.items():
        if not group_df.empty:
            table_html = group_df.to_html(index=False, classes='salary-table', border=0, justify='center', na_rep='')
            tables_html += f"""
            <div class="status-section">
                <h3>{status}（共{len(group_df)}条）</h3>
                {table_html}
            </div>
            """

    return f"""
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
                .footer {{ margin-top: 20px; color: #777; font-size: 12px; text-align: center; }}
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


def get_last_month_str():
    """获取上个月的年月字符串（格式：YYYY-MM）"""
    today = datetime.today().replace(day=1)
    last_month = today - timedelta(days=1)
    return last_month.strftime('%Y-%m')


def main():
    """主执行函数"""
    logger.info("▶ 开始工资核对流程")

    # 1. 获取数据
    salary_month = get_last_month_str()
    salary_df = get_salary_data(salary_month)
    if salary_df.empty:
        logger.error("获取工资数据失败，流程终止")
        return

    # 2. 获取核对人信息
    checker_df = get_github_data("BYTX-YGJ/excel/contents/工资核算人统计.xlsx")
    email_df = get_github_data("BYTX-YGJ/excel/contents/邮箱维护.xlsx")
    if checker_df.empty or email_df.empty:
        logger.error("获取核对人信息失败，流程终止")
        return

    # 3. 合并数据并发送报告
    final_df = pd.merge(
        salary_df,
        checker_df[['项目组', '工资核对人']].rename(columns={'工资核对人': '核对人'}),
        on='项目组',
        how='left'
    )
    send_complete_salary_report(final_df, email_df)

    logger.info("✅ 工资核对流程完成")


if __name__ == "__main__":
    main()