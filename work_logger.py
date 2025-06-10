# work_logger.py
from mcp.server.fastmcp import FastMCP
import sys
import logging
import os
from datetime import datetime
import json

logger = logging.getLogger('Work_Logger')

# Fix UTF-8 encoding for Windows console
if sys.platform == 'win32':
    sys.stderr.reconfigure(encoding='utf-8')
    sys.stdout.reconfigure(encoding='utf-8')

# Create an MCP server
mcp = FastMCP("Work_Logger")

@mcp.tool()
def add_work_log(content: str):
    """Add content to today's work log.添加工作日志内容
    
    Args:
        content: 要添加的工作日志内容
    """
    try:
        # 获取桌面路径
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        # 创建工作日志文件夹
        log_folder = os.path.join(desktop, "工作日志")
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
            
        # 生成今天的日期作为文件名
        today = datetime.now().strftime("%Y-%m-%d")
        log_file = os.path.join(log_folder, f"{today}.txt")
        
        # 获取当前时间
        current_time = datetime.now().strftime("%H:%M:%S")
        
        # 准备要写入的内容
        log_entry = f"\n[{current_time}]\n{content}\n"
        log_entry += "-" * 50  # 添加分隔线
        
        # 写入文件
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(log_entry)
            
        logger.info(f"Added work log: {content[:50]}...")
        return {"success": True, "message": f'成功添加工作日志: {content[:50]}...'}
        
    except Exception as e:
        logger.error(f"Error adding work log: {str(e)}")
        return {"success": False, "message": f"添加工作日志失败: {str(e)}"}

# Start the server
if __name__ == "__main__":
    mcp.run(transport="stdio") 