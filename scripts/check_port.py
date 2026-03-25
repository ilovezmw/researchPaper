#!/usr/bin/env python3
"""
检测本机 TCP 端口是否可被绑定（用于部署前确认不与现有服务冲突）。
用法: python scripts/check_port.py 8765
"""
from __future__ import annotations

import socket
import sys


def main() -> None:
    if len(sys.argv) < 2:
        print("用法: python scripts/check_port.py <端口号>")
        sys.exit(2)
    port = int(sys.argv[1])
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind(("0.0.0.0", port))
        print(f"端口 {port} 当前可被绑定（很可能空闲）。")
        sys.exit(0)
    except OSError as e:
        print(f"端口 {port} 无法绑定: {e}")
        sys.exit(1)
    finally:
        s.close()


if __name__ == "__main__":
    main()
