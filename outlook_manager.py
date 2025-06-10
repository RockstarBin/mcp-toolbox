# outlook_manager.py
# 此文件由 calendar_manager.py 重命名并替换内部标识而来
# 主要功能：通过 Outlook COM 接口提供日历、邮件、联系人等 MCP 工具

# 以下代码除名称替换外保持一致
from mcp.server.fastmcp import FastMCP
import sys
import logging
from datetime import datetime, timedelta
import os
import json
import re
from typing import Optional, List

# --- Outlook 依赖 ---
if sys.platform == 'win32':
    try:
        import win32com.client  # 需要安装 pywin32
        import pythoncom
        from pywintypes import com_error
    except ImportError:
        win32com = None
        com_error = None
        logging.getLogger("Outlook_Manager").warning("未找到 pywin32，Outlook 功能将不可用")
else:
    win32com = None  # 非 Windows 系统不支持 Outlook COM 接口

logger = logging.getLogger('Outlook_Manager')

# 修复Windows控制台的UTF-8编码
if sys.platform == 'win32':
    sys.stderr.reconfigure(encoding='utf-8')
    sys.stdout.reconfigure(encoding='utf-8')

# 创建MCP服务器
mcp = FastMCP("Outlook_Manager")

# MCP服务接口
@mcp.tool()
async def add_calendar_event(*args, **kwargs):
    """(已废弃) 本地日历功能已删除"""
    return {"success": False, "message": "本地日历功能已删除"}

@mcp.tool()
async def delete_calendar_event(*args, **kwargs):
    return {"success": False, "message": "本地日历功能已删除"}

@mcp.tool()
async def list_reminders(*args, **kwargs):
    return {"success": False, "message": "本地日历功能已删除"}

# ================== Outlook 日历 MCP 工具 ==================

# Outlook 日历文件夹常量
OL_FOLDER_CALENDAR = 9  # olFolderCalendar
OL_FOLDER_CONTACTS = 10  # olFolderContacts

# --- Outlook 实用函数 ---

def _dispatch_outlook_application():
    """尝试创建 Outlook.Application COM 对象，兼容不同版本注册名"""
    if win32com is None:
        raise RuntimeError("pywin32 未安装或当前系统不支持 Outlook")

    # 确保线程已初始化 COM
    pythoncom.CoInitialize()

    progids = [
        "Outlook.Application",
        "Outlook.Application.16",
        "Outlook.Application.15",
        "Outlook.Application.14",
    ]
    last_err = None
    for progid in progids:
        try:
            outlook = win32com.client.Dispatch(progid)
            return outlook
        except com_error as e:
            last_err = e
            continue

    raise RuntimeError(f"无法启动 Outlook，可能未安装或 COM 注册损坏: {last_err}")


def _get_outlook_calendar_folder():
    """返回 Outlook 默认日历文件夹；若不可用抛出异常"""
    outlook = _dispatch_outlook_application()
    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)


def _get_outlook_contacts_folder():
    outlook = _dispatch_outlook_application()
    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)


@mcp.tool()
async def add_outlook_event(title: str, start_time: str, end_time: str | None = None, description: str = "", reminder_minutes: int = 15, importance: str = "normal"):
    """添加 Outlook 日历提醒

    参数
    -----
    title:        主题
    start_time:   "YYYY-MM-DD HH:MM" 格式
    end_time:     若为空，默认为开始时间后 30 分钟
    description:  备注
    reminder_minutes: 提前多少分钟提醒
    importance: low / normal / high
    """
    try:
        # ---- 时间解析 ----
        if start_time:
            start_dt = datetime.strptime(start_time, "%Y-%m-%d %H:%M")
        else:
            text = title + " " + description

            start_dt = None

            # 1）使用 dateparser 解析自然语言（若安装）
            if dateparser:
                try:
                    start_dt = dateparser.parse(
                        text,
                        languages=["zh", "en"],
                        settings={"PREFER_DATES_FROM": "future", "RELATIVE_BASE": datetime.now()}
                    )
                except Exception:
                    start_dt = None

            # 2）如仍失败，手工解析常见中文表达
            if start_dt is None:
                cn_map = {"零":0,"〇":0,"一":1,"二":2,"两":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9}

                def cn_num_to_int(cn:str)->int:
                    cn = cn.strip()
                    if cn in cn_map:
                        return cn_map[cn]
                    if cn.startswith("十"):
                        return 10 + cn_map.get(cn[1:], 0) if len(cn) > 1 else 10
                    if "十" in cn:
                        parts = cn.split("十")
                        tens = cn_map.get(parts[0],0)
                        ones = cn_map.get(parts[1],0) if len(parts)>1 else 0
                        return tens*10 + ones
                    return 0

                # 判断日期关键词
                base_date = datetime.now().date()
                if "明天" in text:
                    base_date += timedelta(days=1)
                elif "后天" in text:
                    base_date += timedelta(days=2)

                # 识别时间段
                m1 = re.search(r"(\d{1,2})[:：](\d{2})", text)
                m2 = re.search(r"([零〇一二两三四五六七八九十]+)点半", text)
                m3 = re.search(r"([零〇一二两三四五六七八九十]+)点(一刻|三刻)", text)
                m4 = re.search(r"([零〇一二两三四五六七八九十]+)点", text)

                if m1:
                    hh = int(m1.group(1)); mm = int(m1.group(2))
                elif m2:
                    hh = cn_num_to_int(m2.group(1)); mm = 30
                elif m3:
                    hh = cn_num_to_int(m3.group(1)); mm = 15 if m3.group(2)=="一刻" else 45
                elif m4:
                    hh = cn_num_to_int(m4.group(1)); mm = 0
                else:
                    raise ValueError("无法解析时间，请提供 start_time 或明确时间表达")

                start_dt = datetime.combine(base_date, datetime.min.time()) + timedelta(hours=hh, minutes=mm)

        end_dt = datetime.strptime(end_time, "%Y-%m-%d %H:%M") if end_time else start_dt + timedelta(minutes=30)

        outlook, _ = _get_outlook_calendar_folder()
        appt = outlook.CreateItem(1)  # 1 = olAppointmentItem
        appt.Subject = title
        appt.Body = description
        appt.Start = start_dt
        appt.End = end_dt

        # ---- 重要性 ----
        importance_map = {"low": 0, "normal": 1, "high": 2, "低": 0, "普通": 1, "高": 2}
        appt.Importance = importance_map.get(importance.lower(), 1)

        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes

        appt.Save()

        return {"success": True, "message": f"已添加 Outlook 日历事件: {title}"}
    except Exception as e:
        logger.exception("添加 Outlook 事件失败")
        return {"success": False, "message": f"添加 Outlook 事件失败: {str(e)}"}


@mcp.tool()
async def delete_outlook_event(title: str, start_time: str):
    """删除匹配标题与开始时间的 Outlook 日历事件"""
    try:
        target_dt = datetime.strptime(start_time, "%Y-%m-%d %H:%M")

        _, calendar_folder = _get_outlook_calendar_folder()

        # Outlook 需要先排序后 Restrict 才能保证日期过滤正常
        items = calendar_folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        # Outlook Restrict 日期格式为 MM/DD/YYYY HH:MM AM/PM
        dt_str = target_dt.strftime("%m/%d/%Y %I:%M %p")
        restriction = f"[Subject] = '{title}' AND [Start] = '{dt_str}'"
        matches = items.Restrict(restriction)

        delete_count = 0
        for item in list(matches):
            item.Delete()
            delete_count += 1

        if delete_count:
            return {"success": True, "message": f"已删除 {delete_count} 条事件"}
        else:
            return {"success": False, "message": "未找到匹配事件"}
    except Exception as e:
        logger.exception("删除 Outlook 事件失败")
        return {"success": False, "message": f"删除 Outlook 事件失败: {str(e)}"}


@mcp.tool()
async def list_outlook_events(start_date: str = "", end_date: str = ""):
    """列出指定日期范围内的 Outlook 日历事件。若参数为空，则默认列出未来 7 天。"""
    try:
        if not start_date:
            start_dt = datetime.now()
        else:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d")

        if not end_date:
            end_dt = start_dt + timedelta(days=7)
        else:
            end_dt = datetime.strptime(end_date, "%Y-%m-%d")

        _, calendar_folder = _get_outlook_calendar_folder()

        items = calendar_folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        restriction = (
            f"[Start] >= '{start_dt.strftime('%m/%d/%Y %I:%M %p')}' AND "
            f"[End] <= '{(end_dt + timedelta(days=1)).strftime('%m/%d/%Y %I:%M %p')}'"
        )
        restricted = items.Restrict(restriction)

        events = []
        for itm in restricted:
            events.append({
                "subject": itm.Subject,
                "start": itm.Start.Format("yyyy-MM-dd HH:mm"),
                "end": itm.End.Format("yyyy-MM-dd HH:mm"),
                "body": itm.Body
            })

        return {"success": True, "message": events}
    except Exception as e:
        logger.exception("获取 Outlook 事件失败")
        return {"success": False, "message": f"获取 Outlook 事件失败: {str(e)}"}

# ---------------------------------------------------------------------
# Outlook 发送邮件功能
# ---------------------------------------------------------------------

# 用于自然语言时间解析（可选依赖）
try:
    import dateparser  # type: ignore
except ImportError:
    dateparser = None

# 在 dateparser 不可用时提示
if dateparser is None:
    logger.warning("未安装 dateparser，自然语言时间解析能力受限。建议 pip install dateparser")

@mcp.tool()
async def send_outlook_mail(to: str, subject: str, body: str, cc: str = "", importance: str = "normal", attachments: Optional[List[str]] = None):
    """发送 Outlook 邮件

    参数
    -----
    to: 收件人地址，多个地址用分号分隔
    subject: 邮件主题
    body: 邮件正文（纯文本）
    cc: 抄送地址，可为空
    importance: low / normal / high
    attachments: 附件文件路径列表
    """
    try:
        outlook = _dispatch_outlook_application()
        mail = outlook.CreateItem(0)  # olMailItem

        # ---- 收件人解析 ----
        def _add_recipients(addr_str:str, field:str):
            addr_str = addr_str.strip()
            if not addr_str:
                return
            for addr in addr_str.split(";"):
                addr = addr.strip()
                if addr:
                    mail.Recipients.Add(addr)

        _add_recipients(to, "To")
        if cc:
            _add_recipients(cc, "CC")

        # ---- 添加统一签名 ----
        signature = "\n\n--\n该邮件发送自MCP服务，如有疑问请回信联系。"
        mail.Body = body + signature

        importance_map = {"low": 0, "normal": 1, "high": 2, "低": 0, "普通": 1, "高": 2}
        mail.Importance = importance_map.get(importance.lower(), 1)

        if attachments:
            for path in attachments:
                if os.path.isfile(path):
                    mail.Attachments.Add(os.path.abspath(path))
                else:
                    logger.warning(f"附件不存在: {path}")

        # 尝试解析收件人名称
        if not mail.Recipients.ResolveAll():
            unresolved = [r.Name for r in mail.Recipients if not r.Resolved]
            return {"success": False, "message": f"收件人无法解析: {unresolved}"}

        mail.Send()
        return {"success": True, "message": "邮件已发送"}
    except Exception as e:
        logger.exception("发送邮件失败")
        return {"success": False, "message": f"发送邮件失败: {str(e)}"}

# ---------------------------------------------------------------------
# Outlook 新建联系人功能
# ---------------------------------------------------------------------

@mcp.tool()
async def add_outlook_contact(name: str, email: str = "", company: str = "", job_title: str = "", phone: str = "", address: str = ""):
    """创建/添加联系人

    参数
    -----
    name:      联系人姓名（必填）
    email:     电子邮箱，可省略
    company:   公司，可省略
    job_title: 职位，可省略
    phone:     电话，可省略
    address:   地址，可省略
    """
    try:
        outlook = _dispatch_outlook_application()
        contact = outlook.CreateItem(2)  # 2 = olContactItem

        contact.FullName = name
        if email:
            contact.Email1Address = email

        if company:
            contact.CompanyName = company
        if job_title:
            contact.JobTitle = job_title
        if phone:
            contact.BusinessTelephoneNumber = phone
        if address:
            contact.BusinessAddress = address
            contact.MailingAddress = address

        contact.Save()

        return {"success": True, "message": f"已创建联系人: {name}"}
    except Exception as e:
        logger.exception("创建联系人失败")
        return {"success": False, "message": f"创建联系人失败: {str(e)}"}

# ---------------------------------------------------------------------
# 删除 / 更新联系人功能
# ---------------------------------------------------------------------

@mcp.tool()
async def delete_outlook_contact(identifier: str):
    """按姓名或邮箱删除联系人

    identifier: 联系人姓名或 Email1Address
    """
    try:
        _, contacts_folder = _get_outlook_contacts_folder()
        items = contacts_folder.Items

        deleted = 0
        for it in list(items):  # 转为 list 避免迭代时修改集合
            if (it.FullName and it.FullName.lower() == identifier.lower()) or (
                it.Email1Address and it.Email1Address.lower() == identifier.lower()
            ):
                it.Delete()
                deleted += 1

        if deleted:
            return {"success": True, "message": f"已删除 {deleted} 个联系人"}
        else:
            return {"success": False, "message": "未找到匹配联系人"}
    except Exception as e:
        logger.exception("删除联系人失败")
        return {"success": False, "message": f"删除联系人失败: {str(e)}"}

@mcp.tool()
async def update_outlook_contact(identifier: str, name: str = "", email: str = "", company: str = "", job_title: str = "", phone: str = "", address: str = ""):
    """修改联系人信息

    identifier: 现有联系人姓名或邮箱（用于查找）
    其余参数：若提供则更新相应字段
    """
    try:
        _, contacts_folder = _get_outlook_contacts_folder()
        items = contacts_folder.Items

        updated = 0
        for it in items:
            if (it.FullName and it.FullName.lower() == identifier.lower()) or (
                it.Email1Address and it.Email1Address.lower() == identifier.lower()
            ):
                if name:
                    it.FullName = name
                if email:
                    it.Email1Address = email
                if company:
                    it.CompanyName = company
                if job_title:
                    it.JobTitle = job_title
                if phone:
                    it.BusinessTelephoneNumber = phone
                if address:
                    it.BusinessAddress = address
                    it.MailingAddress = address
                it.Save()
                updated += 1

        if updated:
            return {"success": True, "message": f"已更新 {updated} 个联系人"}
        else:
            return {"success": False, "message": "未找到匹配联系人"}
    except Exception as e:
        logger.exception("更新联系人失败")
        return {"success": False, "message": f"更新联系人失败: {str(e)}"}

# ---------------------------------------------------------------------
# 批量删除日历事件（改进删除功能）
# ---------------------------------------------------------------------

@mcp.tool()
async def delete_outlook_events(subject_keyword: str = "", start_date: str = "", end_date: str = "", delete_all: bool = False):
    """删除 Outlook 日历事件

    参数说明：
    subject_keyword  若提供，则仅删除主题包含该关键字的事件
    start_date/end_date YYYY-MM-DD，可限定时间范围；空则不限
    delete_all        True 时忽略其它过滤条件，直接删除全部事件（危险）
    """
    try:
        _, calendar_folder = _get_outlook_calendar_folder()
        items = calendar_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # 构造过滤日期范围
        if start_date:
            sd = datetime.strptime(start_date, "%Y-%m-%d")
        else:
            sd = datetime(1900,1,1)
        if end_date:
            ed = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
        else:
            ed = datetime(9999,12,31)

        count = 0

        for it in list(items):
            if not delete_all:
                if not (sd <= it.Start <= ed):
                    continue
                if subject_keyword and subject_keyword not in it.Subject:
                    continue
            it.Delete()
            count += 1

        return {"success": True, "message": f"已删除 {count} 条事件"}
    except Exception as e:
        logger.exception("批量删除事件失败")
        return {"success": False, "message": f"批量删除事件失败: {str(e)}"}

# 启动服务器
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    try:
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"服务启动失败: {e}")
        sys.exit(1) 