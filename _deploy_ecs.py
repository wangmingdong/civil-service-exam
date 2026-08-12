# -*- coding: utf-8 -*-
# ECS 部署：用 SSH 通道 `cat >` 流式直传（绕开 paramiko SFTP 16MB 静默失败），
# 传完 wc -c 校验字节，再 systemctl restart cse_http。
# 密码从环境变量 CSE_ECS_PASS 读取（不落文件）。
import os, sys, paramiko
HOST='121.40.208.54'; PORT=22; USER='root'
REMOTE_DIR='/root/civil-service-exam'
PASS=os.environ.get('CSE_ECS_PASS')
if not PASS:
    print("ERROR: 请先设置环境变量 CSE_ECS_PASS"); sys.exit(2)
LOCAL=[('web_dist/index.html','index.html'),('web_dist/data.js','data.js')]

def stream_upload(sftp_or_ssh, ssh, local, remote):
    size=os.path.getsize(local)
    tmp=remote+'.tmp'
    # 用 exec_command 走 shell 重定向，分块写 stdin
    cmd="cat > %s" % tmp
    stdin,stdout,stderr=ssh.exec_command(cmd)
    with open(local,'rb') as f:
        while True:
            buf=f.read(1024*1024)
            if not buf: break
            stdin.write(buf)
        stdin.flush()
    stdin.channel.shutdown_write()
    exit_code=stdout.channel.recv_exit_status()
    if exit_code!=0:
        err=stderr.read().decode('utf-8','replace')
        raise RuntimeError("cat 写失败 exit=%d %s" % (exit_code, err))
    # 原子替换 + 校验字节
    ssh.exec_command("mv -f %s %s" % (tmp, remote))
    _,o,e=ssh.exec_command("wc -c < %s" % remote)
    rc=o.channel.recv_exit_status()
    remote_size=int(o.read().decode().strip())
    if remote_size!=size:
        raise RuntimeError("字节校验不符 local=%d remote=%d (%s)" % (size, remote_size, remote))
    print("  ✓ %s -> %s  (%d 字节 一致)" % (local, remote, size))

ssh=paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(HOST, port=PORT, username=USER, password=PASS, timeout=30, look_for_keys=False, allow_agent=False)
print("SSH 连接 %s 成功" % HOST)
# 确保目录存在（首次）
ssh.exec_command("mkdir -p %s" % REMOTE_DIR)
for local, name in LOCAL:
    stream_upload(None, ssh, local, REMOTE_DIR+'/'+name)
# 重启服务
print("重启 cse_http ...")
_,o,e=ssh.exec_command("systemctl restart cse_http")
rc=o.channel.recv_exit_status()
if rc!=0:
    # 兜底：直接 pkill http.server 8765（systemctl 失败也能恢复）
    print("  systemctl 返回 %d，尝试 pkill 兜底" % rc)
    ssh.exec_command("pkill -f 'http.server 8765'; sleep 1; (cd %s && nohup python3 -m http.server 8765 --directory %s --bind 0.0.0.0 >/dev/null 2>&1 &)" % (REMOTE_DIR, REMOTE_DIR))
    print("  pkill+重启 已执行")
else:
    print("  systemctl restart cse_http 成功")
ssh.close()
print("ECS 部署完成：http://121.40.208.54/gk/")
