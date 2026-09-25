#!/usr/bin/env python3
"""
本地数据提交与驾驶舱系统
- 零依赖，纯 Python 标准库
- 支持内网多电脑访问
- Excel (.xlsx) 文件上传
- 驾驶舱数据可视化
"""

import json
import sqlite3
import os
import zipfile
import xml.etree.ElementTree as ET
import html
import re
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
import io
import uuid
import time

# ===================== 配置 =====================
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "team_data.db")
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
HOST = "0.0.0.0"  # 监听所有网卡，允许内网访问
PORT = 8765
NS = {
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ===================== Excel 解析（纯标准库）=====================
def parse_xlsx(file_bytes):
    """解析 .xlsx 文件，返回列名列表 + 数据行列表"""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            # 1. 读取共享字符串表
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_xml = z.read('xl/sharedStrings.xml')
                root = ET.fromstring(ss_xml)
                for si in root.findall('ss:si', NS):
                    texts = []
                    for t in si.findall('.//ss:t', NS):
                        if t.text:
                            texts.append(t.text)
                    shared_strings.append(''.join(texts))

            # 2. 读取第一个工作表
            sheet_files = [f for f in z.namelist() if f.startswith('xl/worksheets/') and f.endswith('.xml')]
            if not sheet_files:
                return [], []
            
            sheet_xml = z.read(sheet_files[0])
            root = ET.fromstring(sheet_xml)
            
            rows_data = []
            for row in root.findall('ss:sheetData/ss:row', NS):
                cells = row.findall('ss:c', NS)
                row_dict = {}
                for cell in cells:
                    cell_ref = cell.get('r', '')
                    col_match = re.match(r'([A-Z]+)', cell_ref)
                    if not col_match:
                        continue
                    col_letter = col_match.group(1)
                    
                    cell_type = cell.get('t', '')
                    val = ''
                    v_elem = cell.find('ss:v', NS)
                    if v_elem is not None and v_elem.text:
                        if cell_type == 's':  # 共享字符串
                            idx = int(v_elem.text)
                            val = shared_strings[idx] if idx < len(shared_strings) else ''
                        elif cell_type == 'inlineStr':
                            is_elem = cell.find('ss:is/ss:t', NS)
                            val = is_elem.text if is_elem is not None and is_elem.text else ''
                        else:
                            val = v_elem.text
                    
                    row_dict[col_letter] = val
                if row_dict:
                    rows_data.append(row_dict)
            
            if not rows_data:
                return [], []
            
            # 第一行作为列名
            headers = []
            all_cols = sorted(rows_data[0].keys(), key=lambda x: (len(x), x))
            for col in all_cols:
                headers.append(rows_data[0].get(col, ''))
            
            # 其余行作为数据
            data = []
            for row in rows_data[1:]:
                record = {}
                for i, col in enumerate(all_cols):
                    if i < len(headers):
                        record[headers[i]] = row.get(col, '')
                data.append(record)
            
            return headers, data
    except Exception as e:
        return [], []

# ===================== multipart 解析 =====================
def parse_multipart(content_type, body):
    """解析 multipart/form-data，正确处理二进制数据"""
    if not content_type or 'boundary=' not in content_type:
        return {}
    
    boundary = content_type.split('boundary=')[1].strip()
    if isinstance(boundary, bytes):
        boundary = boundary.decode('utf-8')
    
    parts = {}
    
    # 使用 bytes 处理，保留二进制数据
    if isinstance(body, str):
        body = body.encode('utf-8', errors='replace')
    
    boundary_bytes = ('--' + boundary).encode('utf-8')
    crlf = b'\r\n'
    double_crlf = b'\r\n\r\n'
    
    # 分割
    raw_parts = body.split(boundary_bytes)
    
    for raw in raw_parts:
        raw = raw.strip()
        if not raw or raw == b'--':
            continue
        
        # 分离 header 和 body
        if double_crlf in raw:
            idx = raw.index(double_crlf)
            header_block = raw[:idx].decode('utf-8', errors='replace')
            part_body = raw[idx + 4:]  # 跳过 \r\n\r\n
            
            # 去掉尾部的 \r\n--boundary 或 \r\n--
            if part_body.endswith(b'\r\n--'):
                part_body = part_body[:-4]
            elif part_body.endswith(b'--'):
                part_body = part_body[:-2]
            
            # 解析 header
            name_match = re.search(r'name="([^"]+)"', header_block)
            filename_match = re.search(r'filename="([^"]+)"', header_block)
            
            if name_match:
                name = name_match.group(1)
                if filename_match:
                    filename = filename_match.group(1)
                    parts[name] = {'filename': filename, 'data': part_body}
                else:
                    # 文本字段，去掉可能的 \r\n 结尾
                    text = part_body.rstrip(b'\r\n').decode('utf-8', errors='replace')
                    parts[name] = {'filename': None, 'data': text}
    
    return parts

# ===================== 数据库 =====================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute('''CREATE TABLE IF NOT EXISTS records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        upload_id TEXT,
        source_file TEXT,
        row_num INTEGER,
        category TEXT DEFAULT '',
        title TEXT DEFAULT '',
        content TEXT DEFAULT '',
        extra_json TEXT DEFAULT '{}',
        uploaded_by TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now', '+8 hours'))
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS uploads (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        upload_id TEXT UNIQUE,
        filename TEXT,
        uploaded_by TEXT,
        row_count INTEGER,
        created_at TEXT DEFAULT (datetime('now', '+8 hours'))
    )''')
    conn.commit()
    return conn

# ===================== HTML 页面 =====================
DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>团队数据驾驶舱</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, "Microsoft YaHei", "Segoe UI", sans-serif; background: #0f172a; color: #e2e8f0; }
.header { background: linear-gradient(135deg, #1e3a5f 0%, #0f172a 100%); padding: 20px 30px; border-bottom: 1px solid #1e3a5f; display: flex; justify-content: space-between; align-items: center; }
.header h1 { font-size: 1.5rem; color: #60a5fa; }
.header .info { font-size: 0.85rem; color: #94a3b8; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; padding: 20px 30px; }
.stat-card { background: #1e293b; border-radius: 10px; padding: 20px; border: 1px solid #334155; }
.stat-card .label { font-size: 0.8rem; color: #94a3b8; margin-bottom: 5px; }
.stat-card .value { font-size: 2rem; font-weight: 700; color: #60a5fa; }
.charts { display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 15px; padding: 0 30px 20px; }
.chart-box { background: #1e293b; border-radius: 10px; padding: 15px; border: 1px solid #334155; }
.chart-box h3 { font-size: 0.9rem; color: #94a3b8; margin-bottom: 10px; }
.chart { width: 100%; height: 300px; }
.data-section { padding: 0 30px 30px; }
.data-section h2 { font-size: 1.1rem; color: #60a5fa; margin-bottom: 15px; }
table { width: 100%; border-collapse: collapse; background: #1e293b; border-radius: 10px; overflow: hidden; }
th, td { padding: 10px 15px; text-align: left; border-bottom: 1px solid #334155; font-size: 0.85rem; }
th { background: #1e3a5f; color: #60a5fa; font-weight: 600; }
td { color: #cbd5e1; }
tr:hover td { background: #253347; }
.empty { text-align: center; color: #475569; padding: 40px; }
.tabs { display: flex; gap: 10px; padding: 0 30px; margin-bottom: 15px; }
.tab { padding: 8px 20px; border-radius: 6px; cursor: pointer; font-size: 0.9rem; border: 1px solid #334155; background: transparent; color: #94a3b8; }
.tab.active { background: #1e3a5f; color: #60a5fa; border-color: #60a5fa; }
</style>
</head>
<body>
<div class="header">
  <h1>团队数据驾驶舱</h1>
  <div class="info" id="serverTime"></div>
</div>

<div class="stats" id="stats"></div>

<div class="tabs">
  <button class="tab active" onclick="switchTab('charts')">图表</button>
  <button class="tab" onclick="switchTab('table')">明细</button>
  <button class="tab" onclick="switchTab('upload')">上传</button>
</div>

<div id="view-charts">
  <div class="charts">
    <div class="chart-box"><h3>类别分布</h3><div class="chart" id="chart-category"></div></div>
    <div class="chart-box"><h3>每日提交趋势</h3><div class="chart" id="chart-trend"></div></div>
    <div class="chart-box"><h3>上传来源分布</h3><div class="chart" id="chart-source"></div></div>
  </div>
</div>

<div id="view-table" style="display:none">
  <div class="data-section">
    <h2>数据明细</h2>
    <table id="dataTable"><thead id="dtHead"></thead><tbody id="dtBody"></tbody></table>
  </div>
</div>

<div id="view-upload" style="display:none">
  <div class="data-section">
    <h2>上传 Excel 文件</h2>
    <div style="background:#1e293b;border-radius:10px;padding:30px;border:2px dashed #334155;text-align:center;">
      <p style="color:#94a3b8;margin-bottom:20px;">支持 .xlsx 格式，第一行为列名</p>
      <form id="uploadForm" enctype="multipart/form-data">
        <div style="margin-bottom:15px;">
          <input type="text" name="uploader" placeholder="上传者姓名" style="padding:10px 15px;border-radius:6px;border:1px solid #334155;background:#0f172a;color:#e2e8f0;width:200px;">
        </div>
        <div style="margin-bottom:15px;">
          <input type="file" name="file" accept=".xlsx" required style="color:#94a3b8;">
        </div>
        <button type="submit" style="padding:10px 30px;background:#3b82f6;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:1rem;">上传并解析</button>
      </form>
      <div id="uploadStatus" style="margin-top:15px;color:#94a3b8;"></div>
    </div>
    <h2 style="margin-top:30px;">上传历史</h2>
    <table id="uploadTable"><thead><tr><th>时间</th><th>上传者</th><th>文件名</th><th>数据行数</th></tr></thead><tbody id="uploadBody"></tbody></table>
  </div>
</div>

<script>
function switchTab(name) {
  document.querySelectorAll('[id^="view-"]').forEach(el => el.style.display = 'none');
  document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
  document.getElementById('view-' + name).style.display = 'block';
  event.target.classList.add('active');
  if (name === 'charts') renderCharts();
  if (name === 'table') renderTable();
  if (name === 'upload') renderUploads();
}

function renderStats() {
  fetch('/api/stats').then(r => r.json()).then(s => {
    document.getElementById('stats').innerHTML = [
      ['总记录数', s.total_records, '#60a5fa'],
      ['上传文件数', s.total_uploads, '#34d399'],
      ['数据类别数', s.total_categories, '#fbbf24'],
      ['最近上传', s.last_upload || '暂无', '#a78bfa']
    ].map(([l, v, c]) => '<div class="stat-card"><div class="label">'+l+'</div><div class="value" style="color:'+c+';font-size:'+(typeof v==='number'?'2rem':'1rem')+'">'+v+'</div></div>').join('');
  });
}

function renderCharts() {
  fetch('/api/chart-data').then(r => r.json()).then(d => {
    // 类别饼图
    const catChart = echarts.init(document.getElementById('chart-category'), 'dark');
    catChart.setOption({
      tooltip: { trigger: 'item' },
      series: [{ type: 'pie', radius: ['40%', '70%'], data: d.categories.map(c => ({name: c.name, value: c.value})), label: { color: '#e2e8f0' } }]
    });
    
    // 趋势折线图
    const trendChart = echarts.init(document.getElementById('chart-trend'), 'dark');
    trendChart.setOption({
      tooltip: { trigger: 'axis' },
      xAxis: { type: 'category', data: d.trend.map(t => t.date), axisLabel: { color: '#94a3b8' } },
      yAxis: { type: 'value', axisLabel: { color: '#94a3b8' } },
      series: [{ type: 'line', data: d.trend.map(t => t.count), smooth: true, areaStyle: { opacity: 0.3 }, itemStyle: { color: '#3b82f6' } }]
    });
    
    // 来源柱状图
    const srcChart = echarts.init(document.getElementById('chart-source'), 'dark');
    srcChart.setOption({
      tooltip: { trigger: 'axis' },
      xAxis: { type: 'category', data: d.sources.map(s => s.name), axisLabel: { color: '#94a3b8', rotate: 30 } },
      yAxis: { type: 'value', axisLabel: { color: '#94a3b8' } },
      series: [{ type: 'bar', data: d.sources.map(s => s.value), itemStyle: { color: '#34d399' } }]
    });
    
    window.addEventListener('resize', () => { catChart.resize(); trendChart.resize(); srcChart.resize(); });
  });
}

function renderTable() {
  fetch('/api/records').then(r => r.json()).then(data => {
    const thead = document.getElementById('dtHead');
    const tbody = document.getElementById('dtBody');
    if (data.length === 0) {
      thead.innerHTML = '';
      tbody.innerHTML = '<tr><td colspan="6" class="empty">暂无数据</td></tr>';
      return;
    }
    const cols = Object.keys(data[0]);
    thead.innerHTML = '<tr>' + cols.map(c => '<th>'+c+'</th>').join('') + '</tr>';
    tbody.innerHTML = data.map(row => '<tr>' + cols.map(c => '<td>'+(row[c]||'')+'</td>').join('') + '</tr>').join('');
  });
}

function renderUploads() {
  fetch('/api/uploads').then(r => r.json()).then(data => {
    const tbody = document.getElementById('uploadBody');
    if (data.length === 0) {
      tbody.innerHTML = '<tr><td colspan="4" class="empty">暂无上传记录</td></tr>';
      return;
    }
    tbody.innerHTML = data.map(d => '<tr><td>'+d.created_at+'</td><td>'+(d.uploaded_by||'-')+'</td><td>'+d.filename+'</td><td>'+d.row_count+'</td></tr>').join('');
  });
}

document.getElementById('uploadForm').addEventListener('submit', function(e) {
  e.preventDefault();
  const status = document.getElementById('uploadStatus');
  status.textContent = '上传中...';
  
  const form = new FormData(this);
  fetch('/api/upload', { method: 'POST', body: form })
    .then(r => r.json())
    .then(result => {
      if (result.error) {
        status.textContent = '错误: ' + result.error;
        status.style.color = '#f87171';
      } else {
        status.textContent = '成功! 解析 ' + result.row_count + ' 行数据';
        status.style.color = '#34d399';
        renderStats();
        this.reset();
      }
    })
    .catch(err => {
      status.textContent = '上传失败: ' + err.message;
      status.style.color = '#f87171';
    });
});

document.getElementById('serverTime').textContent = new Date().toLocaleString('zh-CN');
renderStats();
renderCharts();
setInterval(() => { document.getElementById('serverTime').textContent = new Date().toLocaleString('zh-CN'); }, 60000);
</script>
</body>
</html>"""

# ===================== HTTP 处理器 =====================
class Handler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip('/')
        
        if path == '' or path == '/index.html':
            self._send_html(DASHBOARD_HTML)
        elif path == '/api/records':
            conn = init_db()
            rows = conn.execute('SELECT * FROM records ORDER BY id DESC').fetchall()
            conn.close()
            cols = ['id', 'upload_id', 'source_file', 'row_num', 'category', 'title', 'content', 'extra_json', 'uploaded_by', 'created_at']
            data = [dict(zip(cols, r)) for r in rows]
            self._send_json(data)
        elif path == '/api/stats':
            conn = init_db()
            total_records = conn.execute('SELECT COUNT(*) FROM records').fetchone()[0]
            total_uploads = conn.execute('SELECT COUNT(*) FROM uploads').fetchone()[0]
            total_categories = conn.execute('SELECT COUNT(DISTINCT category) FROM records WHERE category != ""').fetchone()[0]
            last = conn.execute('SELECT created_at FROM uploads ORDER BY id DESC LIMIT 1').fetchone()
            conn.close()
            self._send_json({
                'total_records': total_records,
                'total_uploads': total_uploads,
                'total_categories': total_categories,
                'last_upload': last[0] if last else None
            })
        elif path == '/api/chart-data':
            conn = init_db()
            # 类别分布
            cats = conn.execute('SELECT category, COUNT(*) as cnt FROM records WHERE category != "" GROUP BY category ORDER BY cnt DESC').fetchall()
            categories = [{'name': c[0], 'value': c[1]} for c in cats]
            
            # 每日趋势（近30天）
            trends = conn.execute('''
                SELECT DATE(created_at) as d, COUNT(*) as cnt FROM records 
                GROUP BY d ORDER BY d DESC LIMIT 30
            ''').fetchall()
            trend = [{'date': t[0], 'count': t[1]} for t in trends[::-1]]
            
            # 上传来源
            sources = conn.execute('''
                SELECT COALESCE(uploaded_by, '未知') as name, SUM(row_count) as cnt FROM uploads 
                GROUP BY name ORDER BY cnt DESC
            ''').fetchall()
            sources_data = [{'name': s[0], 'value': s[1]} for s in sources]
            
            conn.close()
            self._send_json({'categories': categories, 'trend': trend, 'sources': sources_data})
        elif path == '/api/uploads':
            conn = init_db()
            rows = conn.execute('SELECT * FROM uploads ORDER BY id DESC').fetchall()
            conn.close()
            cols = ['id', 'upload_id', 'filename', 'uploaded_by', 'row_count', 'created_at']
            data = [dict(zip(cols, r)) for r in rows]
            self._send_json(data)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip('/')
        
        if path == '/api/upload':
            self._handle_upload()
        else:
            self.send_response(404)
            self.end_headers()

    def _handle_upload(self):
        content_type = self.headers.get('Content-Type', '')
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        
        parts = parse_multipart(content_type, body)
        
        uploader = parts.get('uploader', {}).get('data', '') if isinstance(parts.get('uploader'), dict) else ''
        if isinstance(uploader, bytes):
            uploader = uploader.decode('utf-8', errors='replace')
        if isinstance(uploader, dict):
            uploader = uploader.get('data', '')
            if isinstance(uploader, bytes):
                uploader = uploader.decode('utf-8', errors='replace')
        
        file_part = parts.get('file', {})
        if not isinstance(file_part, dict) or 'data' not in file_part:
            self._send_json({'error': '未找到文件'})
            return
        
        filename = file_part.get('filename', 'unknown.xlsx')
        file_data = file_part['data']
        
        # 解析 Excel
        headers, data = parse_xlsx(file_data)
        if not headers or not data:
            self._send_json({'error': 'Excel 解析失败，请检查格式'})
            return
        
        # 保存文件
        upload_id = str(uuid.uuid4())[:8]
        safe_name = f"{upload_id}_{filename}"
        save_path = os.path.join(UPLOAD_DIR, safe_name)
        with open(save_path, 'wb') as f:
            f.write(file_data)
        
        # 存入数据库
        conn = init_db()
        # 尝试自动识别 category 和 title 列
        cat_col = None
        title_col = None
        content_col = None
        for h in headers:
            hl = h.lower()
            if '类' in hl or 'cat' in hl or 'type' in hl or '类型' in hl:
                cat_col = h
            if '题' in hl or 'title' in hl or '名称' in hl or 'name' in hl:
                title_col = h
            if '内' in hl or '容' in hl or 'content' in hl or 'desc' in hl or '描述' in hl or '备注' in hl:
                content_col = h
        
        count = 0
        for row in data:
            extra = {k: v for k, v in row.items() if k not in (cat_col, title_col, content_col)}
            conn.execute('''INSERT INTO records (upload_id, source_file, row_num, category, title, content, extra_json, uploaded_by)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                        (upload_id, filename, count + 1,
                         row.get(cat_col, '') if cat_col else '',
                         row.get(title_col, '') if title_col else '',
                         row.get(content_col, '') if content_col else '',
                         json.dumps(extra, ensure_ascii=False),
                         uploader))
            count += 1
        
        conn.execute('INSERT INTO uploads (upload_id, filename, uploaded_by, row_count) VALUES (?, ?, ?, ?)',
                    (upload_id, filename, uploader, count))
        conn.commit()
        conn.close()
        
        self._send_json({'ok': True, 'row_count': count, 'headers': headers})

    def _send_html(self, html_str):
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(html_str.encode('utf-8'))

    def _send_json(self, data):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def log_message(self, format, *args):
        print(f"[{self.address_string()}] {format % args}")

# ===================== 获取本机IP =====================
def get_local_ip():
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return '127.0.0.1'

# ===================== 启动 =====================
if __name__ == '__main__':
    init_db()
    local_ip = get_local_ip()
    print("=" * 50)
    print("  团队数据驾驶舱已启动")
    print("=" * 50)
    print(f"  本机访问: http://127.0.0.1:{PORT}")
    print(f"  内网访问: http://{local_ip}:{PORT}")
    print(f"  数据库: {DB_PATH}")
    print(f"  上传目录: {UPLOAD_DIR}")
    print("=" * 50)
    print(f"\n  告诉同事访问: http://{local_ip}:{PORT}")
    print("  按 Ctrl+C 停止\n")
    
    server = HTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
        server.shutdown()
