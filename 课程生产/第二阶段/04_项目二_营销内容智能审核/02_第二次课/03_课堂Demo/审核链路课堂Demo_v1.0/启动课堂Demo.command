#!/bin/bash
cd -- "$(dirname -- "$0")" || exit 1
runtime_python='/Users/didi/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3'
if [ -x "$runtime_python" ]; then
  exec "$runtime_python" server.py --open
elif command -v python3 >/dev/null 2>&1; then
  exec python3 server.py --open
else
  echo '需要Python 3才能启动真实模型接口；也可直接打开index.html使用离线演示。'
  read -r -p '按回车关闭。'
fi
