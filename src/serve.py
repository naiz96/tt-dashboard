# -*- coding: utf-8 -*-
"""把看板共享到局域网，同事用浏览器打开就能看。

用法：
  python src/serve.py            前台托管，关窗口或 Ctrl+C 即停止
  python src/serve.py 9000       指定起始端口（默认 8000，被占用就往后试）
  python src/serve.py --ensure   保证共享在跑（已经在跑就复用），打印地址后退出
  python src/serve.py --stop     停掉后台运行的共享

只放出看板本体和 echarts，其他路径一律 404。项目根目录里有 config.json
（含飞书群 webhook 地址）和 data/ 下的原始投放表，整个目录端出去等于把
这些东西摊给同网段所有人，所以这里用白名单，不用 directory=BASE_DIR。

看板是实时读文件的，重新生成 dashboard.html 后同事刷新即最新。
另外提供 /api/data：现场调飞书取数（config.json 配了 source=feishu 时），
页面打开/刷新时自己会来拿，拿到就是最新的；同一套接口云函数里也有一份。
"""
import json
import os
import signal
import socket
import subprocess
import sys
import time
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feishu_sheet

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')
OUT_DIR = os.path.join(BASE_DIR, 'output')
# 运行状态放 output/：打包.py 只把 data/ 和 output/ 排除在外，放这儿不会跟着便携包发给同事
STATE_PATH = os.path.join(OUT_DIR, '共享状态.json')
LOG_PATH = os.path.join(OUT_DIR, '共享日志.txt')
DASH_PATH = os.path.join(BASE_DIR, 'dashboard.html')

# 白名单：URL 路径 -> 项目内的文件。dashboard.html 里只有 echarts 这一处外部引用
# （见 make_html.py 的 <script src="assets/echarts.min.js">），相对路径正好落在 /assets/ 下。
ROUTES = {
    '/': 'dashboard.html',
    '/index.html': 'dashboard.html',
    '/assets/echarts.min.js': 'assets/echarts.min.js',
}

PORT_START = 8000
PORT_TRIES = 20
START_TIMEOUT = 3.0  # --ensure 等后台实例起来的最长时间

API_CACHE_SEC = 30  # /api/data 结果缓 30 秒，防多人连点把飞书接口打爆
_api_cache = {'time': 0.0, 'body': None}


def _err(msg):
    return json.dumps({'error': msg}, ensure_ascii=False).encode('utf-8')


def api_payload():
    """现场调飞书取数 + 清洗，返回 (HTTP 状态码, JSON bytes)。

    成功的结果 30 秒内复用；失败不缓存（改好配置马上就能生效）。
    """
    now = time.time()
    if _api_cache['body'] and now - _api_cache['time'] < API_CACHE_SEC:
        return 200, _api_cache['body']
    with open(CONFIG_PATH, encoding='utf-8') as f:  # config.json 是唯一配置源
        cfg = json.load(f)
    if feishu_sheet.source_of(cfg) != 'feishu':
        return 404, _err('config.json 的 source 不是 feishu，没有实时数据接口')
    try:
        body = json.dumps(feishu_sheet.build_payload(*feishu_sheet.creds_of(cfg)),
                          ensure_ascii=False).encode('utf-8')
    except Exception as e:
        return 503, _err(str(e))
    _api_cache['time'], _api_cache['body'] = now, body
    return 200, body


class Handler(SimpleHTTPRequestHandler):
    # Windows 注册表常把 .js 认成 text/plain，浏览器就不执行 echarts 了，这里写死
    extensions_map = dict(SimpleHTTPRequestHandler.extensions_map, **{
        '.js': 'application/javascript',
        '.html': 'text/html; charset=utf-8',
    })

    def do_GET(self):
        if self.path.split('?', 1)[0].split('#', 1)[0] == '/api/data':
            self.api_data()
            return
        super().do_GET()

    def api_data(self):
        """实时数据接口（页面脚本自己来拿，拿不到就静默用内置快照）。"""
        try:
            code, body = api_payload()
        except Exception as e:  # config.json 读不了之类
            code, body = 503, _err(str(e))
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')  # file:// 打开的页面也来拿
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def translate_path(self, path):
        """路径只当字典的键用，从不参与拼接文件名，所以走不出白名单。"""
        rel = ROUTES.get(path.split('?', 1)[0].split('#', 1)[0])
        if rel is None:
            return os.path.join(BASE_DIR, '__404__')  # 不存在的路径 -> 404
        return os.path.join(BASE_DIR, *rel.split('/'))


def port_busy(port):
    """先连一下看有没有人在听。

    Windows 上 allow_reuse_address 的语义和 Unix 不同，bind 到别人已占的端口
    也会成功（等于抢端口），光靠 bind 失败判断不出来。
    """
    s = socket.socket()
    try:
        s.settimeout(0.2)
        s.connect(('127.0.0.1', port))
        return True
    except OSError:
        return False
    finally:
        s.close()


def lan_ips():
    """本机在局域网里的地址：先按默认路由取主网卡，再补上主机名解析出来的。"""
    ips = []
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # 不发包，只是让系统按路由表挑一块对外的网卡
        s.connect(('223.5.5.5', 80))
        ips.append(s.getsockname()[0])
    except OSError:
        pass
    finally:
        s.close()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith('127.') and ip not in ips:
                ips.append(ip)
    except OSError:
        pass
    return ips


def urls_for(port):
    """同事能用的地址，IP 在前、主机名在后。

    主机名排在后面不是次选：机器上有多块网卡（比如挂着 VPN）时，IP 可能取到
    同事连不上的那块，而主机名由名称服务器解析，往往指向真正的局域网地址。
    """
    ips = lan_ips()
    urls = ['http://%s:%d' % (ip, port) for ip in ips]
    host = socket.gethostname()
    if host not in ips:
        urls.append('http://%s:%d' % (host, port))
    return urls


def read_state():
    """读运行状态，返回 dict 或 None。"""
    try:
        with open(STATE_PATH, encoding='utf-8') as f:
            st = json.load(f)
    except (OSError, ValueError):
        return None
    return st if isinstance(st, dict) and st.get('port') else None


def write_state(port):
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(STATE_PATH, 'w', encoding='utf-8') as f:
        json.dump({'port': port, 'pid': os.getpid()}, f)


def clear_state(mine=False):
    """删状态文件。mine=True 时只在记的是自己时才删，
    免得自己退出时把后来那个实例的记录一并抹掉。"""
    if mine:
        st = read_state()
        if not st or st.get('pid') != os.getpid():
            return
    try:
        os.remove(STATE_PATH)
    except OSError:
        pass


# 绕过系统代理：公司机器上常有 http_proxy，设了的话连 127.0.0.1 都会被导去代理
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def is_ours(port):
    """这个端口上跑的是不是我们的看板：拉一次首页，跟磁盘上的 dashboard.html 逐字节比。

    只看端口通不通不够——我们的进程被杀掉、端口被别的程序占走时状态文件还留着，
    会把同事的链接指向别人的服务。
    """
    try:
        with open(DASH_PATH, 'rb') as f:
            local = f.read()
    except OSError:
        return False
    try:
        with _OPENER.open('http://127.0.0.1:%d/' % port, timeout=1) as r:
            return r.read() == local
    except (OSError, ValueError):
        return False


def ensure_running():
    """保证共享在跑，返回同事能用的地址列表；起不来返回 None。

    已经在跑的实例（上次留下的、或者手动开的）直接复用，不重复起一个。
    调用方即使拿不到地址也必须继续干活——共享起不来不能影响发日报。
    """
    st = read_state()
    if st and port_busy(st['port']) and is_ours(st['port']):
        return urls_for(st['port'])

    os.makedirs(OUT_DIR, exist_ok=True)
    env = dict(os.environ, PYTHONUTF8='1')  # 日志固定 UTF-8，跟控制台代码页无关
    with open(LOG_PATH, 'ab') as log:
        if sys.platform == 'win32':
            # DETACHED_PROCESS：脱离父进程的控制台，父进程（计划任务/主流程）退出后继续活着
            subprocess.Popen([sys.executable, os.path.abspath(__file__)],
                             stdin=subprocess.DEVNULL, stdout=log, stderr=log,
                             cwd=BASE_DIR, env=env, close_fds=True,
                             creationflags=0x00000008)
        else:
            subprocess.Popen([sys.executable, os.path.abspath(__file__)],
                             stdin=subprocess.DEVNULL, stdout=log, stderr=log,
                             cwd=BASE_DIR, env=env, close_fds=True,
                             start_new_session=True)

    deadline = time.time() + START_TIMEOUT
    while time.time() < deadline:
        time.sleep(0.2)
        st = read_state()
        # 端口被别人占着时 is_ours 为假，那是上一轮留下的记录，等子进程写出新的
        if st and port_busy(st['port']) and is_ours(st['port']):
            return urls_for(st['port'])
    return None


def stop():
    """停掉后台跑着的那个共享。"""
    st = read_state()
    if not st:
        print('[!] 没有正在运行的共享')
        return 0
    port, pid = st['port'], st.get('pid')
    # 先确认服务真的在跑再动手：PID 会被系统复用，照着文件里的数字杀可能误伤别的进程
    if isinstance(pid, int) and port_busy(port) and is_ours(port):
        if sys.platform == 'win32':
            subprocess.run(['taskkill', '/F', '/PID', str(pid)], capture_output=True)
        else:
            os.kill(pid, signal.SIGTERM)
        time.sleep(0.3)
    clear_state()
    print('[OK] 已停止共享')
    return 0


def main():
    args = sys.argv[1:]

    if args[:1] == ['--ensure']:
        urls = ensure_running()
        if not urls:
            print('[X] 共享没起来，看日志：%s' % LOG_PATH)
            return 1
        print('[OK] 看板共享已开启，同事在浏览器里打开：')
        for u in urls:
            print('      %s' % u)
        return 0

    if args[:1] == ['--stop']:
        return stop()

    port_start = PORT_START
    if args:
        try:
            port_start = int(args[0])
        except ValueError:
            print('[X] 端口要填数字，比如：python src/serve.py 9000')
            return 2

    if not os.path.exists(DASH_PATH):
        print('[X] 没找到 dashboard.html，先双击 启动.bat 生成一次看板')
        return 1

    httpd = None
    for port in range(port_start, port_start + PORT_TRIES):
        if port_busy(port):
            continue
        try:
            httpd = ThreadingHTTPServer(('0.0.0.0', port), Handler)
            break
        except OSError:
            continue
    if httpd is None:
        print('[X] %d-%d 端口都被占用了，换个起始端口：python src/serve.py %d'
              % (port_start, port_start + PORT_TRIES - 1, port_start + PORT_TRIES))
        return 1

    write_state(port)
    print('[OK] 看板已共享，同事在浏览器里打开：')
    for u in urls_for(port):
        print('      %s' % u)
    print('      (IP 变了就用带电脑名的那个；你自己预览 http://127.0.0.1:%d)' % port)
    print('[!] 第一次运行如果弹防火墙提示，要点"允许访问"')
    print('[!] 关掉这个窗口 = 停止共享；后台跑的用 停止共享.bat 停')
    print('[>] 按 Ctrl+C 停止')
    sys.stdout.flush()  # 输出被重定向时是块缓冲，地址要立刻能看到

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print()
    finally:
        httpd.server_close()
        clear_state(mine=True)
    print('[OK] 已停止共享')
    return 0


if __name__ == '__main__':
    sys.exit(main())
