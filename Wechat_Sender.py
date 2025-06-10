# server.py
from mcp.server.fastmcp import FastMCP
import sys
import logging
import wxauto
from wxauto import WeChat
logger = logging.getLogger('Wechat_Sender')

# Fix UTF-8 encoding for Windows console
if sys.platform == 'win32':
    sys.stderr.reconfigure(encoding='utf-8')
    sys.stdout.reconfigure(encoding='utf-8')

# Create an MCP server
mcp = FastMCP("Wechat_Sender")

@mcp.tool()
def send_message_to_wechat(message: str, username: str):
    """Send a message to a specific WeChat user.给微信用户发送消息"""
    wx = WeChat()
    wx.SendMsg(f'{message}', username)
    logger.info(f"Sending message to WeChat: {message} to {username}")
    return {"success": True, "message": '发送成功'}

# Start the server
if __name__ == "__main__":
    mcp.run(transport="stdio") 