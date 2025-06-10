import subprocess
import os
import sys
import time
import signal
import logging
from pathlib import Path

"""start_all_services.py
一键启动本项目中的所有 MCP 工具脚本。

脚本会为每个工具脚本创建一个新的子进程：
    python mcp_pipe.py <tool_script>
从而让它们全部通过 WebSocket 与远端 AI 保持连接。

支持功能：
1. 统一日志：stdout/stderr 重定向到 `<tool_name>.log`。
2. 进程监控：若任一子进程意外退出，将自动尝试重启（带 5 秒延迟）。
3. 优雅退出：捕获 Ctrl-C 或终止信号，关闭所有子进程并退出。

使用方法：
    python start_all_services.py

如需只启动部分服务，可修改 SERVICE_SCRIPTS 列表或使用命令行参数：
    python start_all_services.py app_launcher calendar_manager
"""

# 默认需要启动的工具脚本（文件名）
SERVICE_SCRIPTS = [
    "app_launcher.py",
    "outlook_manager.py",
    "work_logger.py",
    "Wechat_Sender.py",
]

# ---------------------------- 内部实现 ----------------------------

def build_cmd(base_dir: Path, script_name: str) -> list[str]:
    """构造启动命令: python mcp_pipe.py <tool_script>"""
    return [
        sys.executable,  # 当前 Python 解释器
        str(base_dir / "mcp_pipe.py"),
        str(base_dir / script_name),
    ]


def start_service(cmd: list[str], log_path: Path) -> subprocess.Popen:
    """启动单个服务，并把输出写入 log_path"""
    log_file = open(log_path, "a", encoding="utf-8", buffering=1)  # 行缓冲
    process = subprocess.Popen(
        cmd,
        stdout=log_file,
        stderr=subprocess.STDOUT,
        text=True,
    )
    return process


def main():
    base_dir = Path(__file__).resolve().parent

    # 如果用户通过命令行传入服务名，则仅启动指定服务
    requested = sys.argv[1:] if len(sys.argv) > 1 else None
    scripts_to_run = [f"{name if name.endswith('.py') else name + '.py'}" for name in (requested or SERVICE_SCRIPTS)]

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [MAIN] %(levelname)s: %(message)s")
    logger = logging.getLogger("SERVICE_MANAGER")

    processes: dict[str, subprocess.Popen] = {}

    def launch_all():
        """启动所有脚本进程"""
        for script in scripts_to_run:
            cmd = build_cmd(base_dir, script)
            log_path = base_dir / f"{Path(script).stem}.log"
            proc = start_service(cmd, log_path)
            processes[script] = proc
            logger.info(f"启动 {script} (PID={proc.pid})，日志→ {log_path.name}")

    def shutdown_all():
        """尝试优雅终止所有子进程"""
        logger.info("收到退出信号，正在关闭所有子进程 …")
        for script, proc in processes.items():
            if proc.poll() is None:  # 仍在运行
                logger.info(f"终止 {script} (PID={proc.pid}) …")
                proc.terminate()
        # 再等待一会
        deadline = time.time() + 10
        for script, proc in processes.items():
            if proc.poll() is None and time.time() < deadline:
                try:
                    proc.wait(timeout=max(0, deadline - time.time()))
                except subprocess.TimeoutExpired:
                    pass
            if proc.poll() is None:  # 依旧没退出
                logger.warning(f"强制杀死 {script} (PID={proc.pid})")
                proc.kill()
        logger.info("已退出。")

    # 注册信号处理
    signal.signal(signal.SIGINT, lambda sig, frame: shutdown_all() or sys.exit(0))
    if os.name != "nt":  # Windows 仅支持部分信号
        signal.signal(signal.SIGTERM, lambda sig, frame: shutdown_all() or sys.exit(0))

    launch_all()

    # 主循环：检测子进程存活，自动重启
    try:
        while True:
            for script, proc in list(processes.items()):
                retcode = proc.poll()
                if retcode is not None:  # 已退出
                    logger.warning(f"{script} 意外退出 (code={retcode})，5 秒后重启 …")
                    time.sleep(5)
                    cmd = build_cmd(base_dir, script)
                    log_path = base_dir / f"{Path(script).stem}.log"
                    processes[script] = start_service(cmd, log_path)
                    logger.info(f"已重启 {script} (PID={processes[script].pid})")
            time.sleep(3)
    except KeyboardInterrupt:
        # 再保险：捕获不到信号时仍能关闭
        shutdown_all()


if __name__ == "__main__":
    main() 