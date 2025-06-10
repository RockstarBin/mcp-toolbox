# app_launcher.py
from mcp.server.fastmcp import FastMCP
import sys
import logging
import subprocess
import os
import webbrowser
from urllib.parse import quote
import psutil
import asyncio

logger = logging.getLogger('App_Launcher')

# Fix UTF-8 encoding for Windows console
if sys.platform == 'win32':
    sys.stderr.reconfigure(encoding='utf-8')
    sys.stdout.reconfigure(encoding='utf-8')

# Create an MCP server
mcp = FastMCP("App_Launcher")

# 常用应用程序路径（多个可能的路径）
APP_PATHS = {
    "wechat": [
        r"D:\WeChat\WeChat.exe",
        r"C:\Program Files\Tencent\WeChat\WeChat.exe",
        os.path.expanduser(r"~\AppData\Local\Tencent\WeChat\WeChat.exe")
    ],
    "qqmusic": [
        r"D:\Program Files (x86)\Tencent\QQMusic\QQMusic2145.01.44.46\QQMusic.exe",
        r"C:\Program Files\Tencent\QQMusic\QQMusic.exe",
        os.path.expanduser(r"~\AppData\Local\Tencent\QQMusic\QQMusic.exe")
    ],
    "cursor": [
        r"C:\Users\{}\AppData\Local\Programs\Cursor\Cursor.exe".format(os.getenv('USERNAME')),
        r"C:\Program Files\Cursor\Cursor.exe"
    ],
    "chrome": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ],
    "wps": [
        r"C:\Users\DGG\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe",
        r"C:\Program Files\WPS Office\11.1.0.11664\office6\wps.exe"
    ],
    "wemeet": [
        r"I:\腾讯会议\WeMeet\wemeetapp.exe",
        r"C:\Program Files\Tencent\WeMeet\WeMeet.exe"
    ],
    "steam": [
        r"D:\Program Files (x86)\Steam\Steam.exe",
        r"C:\Program Files\Steam\Steam.exe"
    ],
    "dota2": [
        r"C:\Program Files (x86)\Steam\steamapps\common\dota 2 beta\game\bin\win64\dota2.exe",
        r"C:\Program Files\Steam\steamapps\common\dota 2 beta\game\bin\win64\dota2.exe"
    ],
    "cmd": "cmd.exe"  # CMD是系统命令，不需要完整路径
}

# 进程名称映射（用于关闭程序）
PROCESS_NAMES = {
    "wechat": "WeChat.exe",
    "qqmusic": "QQMusic.exe",
    "cursor": "Cursor.exe",
    "chrome": "chrome.exe",
    "wps": "wps.exe",
    "wemeet": "wemeetapp.exe",
    "steam": "Steam.exe",
    "dota2": "dota2.exe",
    "cmd": "cmd.exe"
}

def find_app_path(app_name: str) -> str:
    """查找应用程序的实际路径"""
    if app_name not in APP_PATHS:
        return None
        
    paths = APP_PATHS[app_name]
    if isinstance(paths, str):
        return paths if os.path.exists(paths) else None
        
    for path in paths:
        if os.path.exists(path):
            return path
    return None

@mcp.tool()
def open_application(app_name: str):
    """Open a Windows application.打开Windows应用程序
    
    Args:
        app_name: 应用程序名称 (wechat/qqmusic/cursor/chrome/wps/wemeet/steam/dota2/cmd)
    """
    try:
        app_name = app_name.lower()
        app_path = find_app_path(app_name)
        
        if not app_path:
            return {"success": False, "message": f"找不到应用程序: {app_name}"}
            
        subprocess.Popen(app_path)
        logger.info(f"Opened application: {app_name}")
        return {"success": True, "message": f'成功打开: {app_name}'}
        
    except Exception as e:
        logger.error(f"Error opening application: {str(e)}")
        return {"success": False, "message": f"打开应用程序失败: {str(e)}"}

@mcp.tool()
def close_application(app_name: str):
    """Close a Windows application.关闭Windows应用程序
    
    Args:
        app_name: 应用程序名称 (wechat/qqmusic/cursor/chrome/wps/wemeet/steam/dota2/cmd)
    """
    try:
        app_name = app_name.lower()
        if app_name not in PROCESS_NAMES:
            return {"success": False, "message": f"不支持的应用程序: {app_name}"}
            
        process_name = PROCESS_NAMES[app_name]
        closed = False
        
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == process_name.lower():
                    proc.terminate()
                    closed = True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
                
        if closed:
            logger.info(f"Closed application: {app_name}")
            return {"success": True, "message": f'成功关闭: {app_name}'}
        else:
            return {"success": False, "message": f'未找到运行中的: {app_name}'}
            
    except Exception as e:
        logger.error(f"Error closing application: {str(e)}")
        return {"success": False, "message": f"关闭应用程序失败: {str(e)}"}

@mcp.tool()
def search_google(query: str):
    """Search Google in Chrome browser.使用Chrome浏览器搜索Google
    
    Args:
        query: 搜索关键词
    """
    try:
        # 构建Google搜索URL
        search_url = f"https://www.google.com/search?q={quote(query)}"
        
        # 尝试使用Chrome打开
        chrome_path = find_app_path("chrome")
        if chrome_path:
            subprocess.Popen([chrome_path, search_url])
        else:
            # 如果找不到Chrome，使用默认浏览器
            webbrowser.open(search_url)
            
        logger.info(f"Searched Google for: {query}")
        return {"success": True, "message": f'成功搜索: {query}'}
        
    except Exception as e:
        logger.error(f"Error searching Google: {str(e)}")
        return {"success": False, "message": f"搜索失败: {str(e)}"}

@mcp.tool()
def search_baidu(query: str):
    """Search Baidu in Chrome browser.使用Chrome浏览器搜索百度
    
    Args:
        query: 搜索关键词
    """
    try:
        # 构建百度搜索URL
        search_url = f"https://www.baidu.com/s?wd={quote(query)}"
        
        # 尝试使用Chrome打开
        chrome_path = find_app_path("chrome")
        if chrome_path:
            subprocess.Popen([chrome_path, search_url])
        else:
            # 如果找不到Chrome，使用默认浏览器
            webbrowser.open(search_url)
            
        logger.info(f"Searched Baidu for: {query}")
        return {"success": True, "message": f'成功搜索: {query}'}
        
    except Exception as e:
        logger.error(f"Error searching Baidu: {str(e)}")
        return {"success": False, "message": f"搜索失败: {str(e)}"}

@mcp.tool()
async def launch_dota2():
    """Launch Dota 2 through Steam.通过Steam启动Dota 2
    
    Returns:
        dict: 包含操作结果的字典
    """
    try:
        # 首先检查Steam是否运行
        steam_running = False
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == 'steam.exe':
                    steam_running = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        # 如果Steam没有运行，先启动Steam
        if not steam_running:
            steam_path = find_app_path("steam")
            if not steam_path:
                return {"success": False, "message": "找不到Steam程序"}
            
            subprocess.Popen(steam_path)
            # 等待Steam启动
            await asyncio.sleep(10)  # 给Steam一些启动时间

        # 使用Steam协议启动Dota 2
        dota2_url = "steam://rungameid/570"  # Dota 2的Steam App ID是570
        webbrowser.open(dota2_url)
        
        logger.info("Launched Dota 2 through Steam")
        return {"success": True, "message": "成功启动Dota 2"}
        
    except Exception as e:
        logger.error(f"Error launching Dota 2: {str(e)}")
        return {"success": False, "message": f"启动Dota 2失败: {str(e)}"}

# Start the server
if __name__ == "__main__":
    mcp.run(transport="stdio") 