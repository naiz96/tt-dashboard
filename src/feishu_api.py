# -*- coding: utf-8 -*-
"""飞书开放平台 API 封装（零依赖，urllib + json）。

- tenant_access_token 获取与 2 小时文件缓存
- 多维表格（Bitable）：建表、列记录、批量写入、批量删除
- 群自定义机器人 webhook（含可选签名校验）
"""
import base64
import hashlib
import hmac
import json
import os
import time
import urllib.request

BASE = 'https://open.feishu.cn/open-apis'
_CACHE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.cache')


def _http(method, url, headers=None, body=None):
    data = json.dumps(body).encode('utf-8') if body is not None else None
    req = urllib.request.Request(url, method=method, headers=headers or {}, data=data)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode('utf-8'))


_MEM = {}  # app_id -> (token, expire)，进程内缓存


def get_tenant_token(app_id, app_secret):
    """获取 tenant_access_token（有效期约 2 小时）。

    两级缓存：先查内存，再查 .cache/token.json。云函数里代码目录只读、容器重启就丢
    文件缓存，所以缓存目录建不出来/写不进去时静默跳过，不影响取 token。
    """
    now = time.time()
    token, expire = _MEM.get(app_id, ('', 0))
    if token and expire > now + 60:
        return token
    cache_file = os.path.join(_CACHE_DIR, 'token.json')
    try:
        with open(cache_file, encoding='utf-8') as f:
            cache = json.load(f)
        if cache.get('app_id') == app_id and cache.get('expire', 0) > now + 60:
            _MEM[app_id] = (cache['token'], cache['expire'])
            return cache['token']
    except (OSError, ValueError):
        pass
    r = _http('POST', BASE + '/auth/v3/tenant_access_token/internal',
              {'Content-Type': 'application/json'},
              {'app_id': app_id, 'app_secret': app_secret})
    if r.get('code') != 0:
        raise RuntimeError('获取飞书 token 失败: %s' % r)
    token = r['tenant_access_token']
    expire = now + r.get('expire', 7200) - 300
    _MEM[app_id] = (token, expire)
    try:
        os.makedirs(_CACHE_DIR, exist_ok=True)
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump({'app_id': app_id, 'token': token, 'expire': expire}, f)
    except OSError:
        pass  # 只读文件系统（云函数）里照样能用，内存缓存兜着
    return token


def call(app_id, app_secret, method, path, body=None):
    """带 token 调用 open API，返回 data 部分；code != 0 抛异常。"""
    token = get_tenant_token(app_id, app_secret)
    r = _http(method, BASE + path,
              {'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'}, body)
    if r.get('code') != 0:
        raise RuntimeError('飞书 API 错误 %s %s: %s' % (method, path, r))
    return r.get('data') or {}


# ---------- 知识库（wiki） ----------

def wiki_get_node(app_id, app_secret, node_token):
    """知识库节点信息：拿 /wiki/<编号> 里的编号换出真实文档 token（obj_token）。

    知识库里的表格，地址栏里的是知识库节点编号、不是 spreadsheet_token，调表格接口前
    先用这个换算。需要应用开通"查看知识库"权限、且应用能访问该知识库节点。
    """
    return call(app_id, app_secret, 'GET',
                '/wiki/v2/spaces/get_node?token=%s&obj_type=wiki' % node_token)


# ---------- Bitable ----------

def bitable_list_tables(app_id, app_secret, app_token):
    r = call(app_id, app_secret, 'GET', '/bitable/v1/apps/%s/tables' % app_token)
    return r.get('items') or []


def bitable_create_table(app_id, app_secret, app_token, name, fields):
    """建表，fields: list[dict]，如 {'field_name': '消耗', 'type': 2}。返回 table_id。"""
    r = call(app_id, app_secret, 'POST', '/bitable/v1/apps/%s/tables' % app_token,
             {'table': {'name': name, 'fields': fields}})
    return r['table_id']


def bitable_list_records(app_id, app_secret, app_token, table_id):
    """读全表记录（自动分页）。返回 list[dict]，每项含 record_id 与 fields。"""
    items = []
    page_token = None
    while True:
        path = '/bitable/v1/apps/%s/tables/%s/records?page_size=500' % (app_token, table_id)
        if page_token:
            path += '&page_token=' + page_token
        r = call(app_id, app_secret, 'GET', path)
        items.extend(r.get('items') or [])
        if r.get('has_more'):
            page_token = r['page_token']
        else:
            return items


def bitable_create_records(app_id, app_secret, app_token, table_id, records):
    """批量插入记录（单批 500 条上限，内部分批）。

    records: list[dict]，字段名 -> 值
    """
    for i in range(0, len(records), 500):
        chunk = [{'fields': rec} for rec in records[i:i + 500]]
        call(app_id, app_secret, 'POST',
             '/bitable/v1/apps/%s/tables/%s/records/batch_create' % (app_token, table_id),
             {'records': chunk})


def bitable_delete_records(app_id, app_secret, app_token, table_id, record_ids):
    """批量删除记录（单批 500 条上限，内部分批）。"""
    for i in range(0, len(record_ids), 500):
        call(app_id, app_secret, 'POST',
             '/bitable/v1/apps/%s/tables/%s/records/batch_delete' % (app_token, table_id),
             {'records': record_ids[i:i + 500]})


# ---------- 群自定义机器人 webhook ----------

def send_webhook(url, secret, body):
    """向群自定义机器人发消息（body 为完整消息体，如 text/post 结构）；secret 非空时加签。"""
    body = dict(body)
    if secret:
        ts = str(int(time.time()))
        sign = base64.b64encode(
            hmac.new(secret.encode('utf-8'),
                     (ts + '\n' + secret).encode('utf-8'),
                     hashlib.sha256).digest()
        ).decode('utf-8')
        body['timestamp'] = ts
        body['sign'] = sign
    r = _http('POST', url, {'Content-Type': 'application/json'}, body)
    ok = r.get('code') == 0 or r.get('StatusCode') == 0
    if not ok:
        raise RuntimeError('webhook 发送失败: %s' % r)
