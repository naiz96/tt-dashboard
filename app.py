# -*- coding: utf-8 -*-
"""TT投放看板 - 云端部署入口

每次打开网页自动从飞书拉最新数据，30秒缓存避免频繁调用。
支持通过环境变量配置飞书凭证（优先）或直接读取 config.json。

部署后设置以下环境变量：
  FEISHU_APP_ID         飞书应用 ID
  FEISHU_APP_SECRET     飞书应用密钥
  FEISHU_SPREADSHEET    飞书电子表格链接或 token
"""
import json
import os
import sys
import time

# 确保 src 在 import 路径中
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))

from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import feishu_sheet

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')
DASH_PATH = os.path.join(BASE_DIR, 'dashboard.html')
API_CACHE_SEC = 30

_api_cache = {'time': 0.0, 'body': None}


def _err(msg):
    return json.dumps({'error': msg}, ensure_ascii=False).encode('utf-8')


def get_config():
    """从环境变量或 config.json 构建配置"""
    app_id = os.environ.get('FEISHU_APP_ID', '')
    app_secret = os.environ.get('FEISHU_APP_SECRET', '')
    spreadsheet = os.environ.get('FEISHU_SPREADSHEET', '')

    if app_id and app_secret and spreadsheet:
        return {
            'source': 'feishu',
            'feishu': {
                'app_id': app_id,
                'app_secret': app_secret,
                'spreadsheet_token': spreadsheet,
            }
        }

    # 回退到 config.json
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)

    raise RuntimeError(
        '未配置飞书凭证。请设置环境变量: '
        'FEISHU_APP_ID, FEISHU_APP_SECRET, FEISHU_SPREADSHEET'
    )


def api_payload():
    """构建 API 响应数据 (HTTP status code, JSON bytes)"""
    now = time.time()
    if _api_cache['body'] and now - _api_cache['time'] < API_CACHE_SEC:
        return 200, _api_cache['body']

    cfg = get_config()
    if feishu_sheet.source_of(cfg) != 'feishu':
        return 404, _err('source 不是 feishu，没有实时数据接口')

    try:
        app_id, app_secret, token = feishu_sheet.creds_of(cfg)
        body = json.dumps(
            feishu_sheet.build_payload(app_id, app_secret, token),
            ensure_ascii=False
        ).encode('utf-8')
    except Exception as e:
        return 503, _err(str(e))

    _api_cache['time'], _api_cache['body'] = now, body
    return 200, body


ROUTES = {
    '/': 'dashboard.html',
    '/index.html': 'dashboard.html',
    '/assets/echarts.min.js': 'assets/echarts.min.js',
}


class Handler(SimpleHTTPRequestHandler):
    extensions_map = dict(SimpleHTTPRequestHandler.extensions_map, **{
        '.js': 'application/javascript',
        '.html': 'text/html; charset=utf-8',
    })

    def do_GET(self):
        clean_path = self.path.split('?', 1)[0].split('#', 1)[0]
        if clean_path == '/api/data':
            self.api_data()
            return
        super().do_GET()

    def api_data(self):
        try:
            code, body = api_payload()
        except Exception as e:
            code, body = 503, _err(str(e))
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def translate_path(self, path):
        rel = ROUTES.get(path.split('?', 1)[0].split('#', 1)[0])
        if rel is None:
            return os.path.join(BASE_DIR, '__404__')
        return os.path.join(BASE_DIR, *rel.split('/'))

    def log_message(self, format, *args):
        """云端日志输出到 stdout"""
        sys.stdout.write("%s - - [%s] %s\n" % (
            self.client_address[0],
            self.log_date_time_string(),
            format % args))
        sys.stdout.flush()


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8080'))
    httpd = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    print(f'[OK] TT投放看板已启动，端口 {port}')
    print(f'[OK] 访问地址: http://0.0.0.0:{port}')
    sys.stdout.flush()
    httpd.serve_forever()
