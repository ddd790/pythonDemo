# -*- coding: utf-8 -*-
"""商务英语情景剧大赛 · 在线评分系统

启动： python app.py
评委入口： http://<本机IP>:8000/
后台入口： http://<本机IP>:8000/admin  （默认密码 drama2026，可用环境变量 ADMIN_PASSWORD 修改）
"""
import io
import os
import socket
import sqlite3
import datetime
from statistics import mean

from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, session, url_for)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'scores.db')

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'drama2026')
SECRET_KEY = os.environ.get('SECRET_KEY', 'business-english-drama-contest-2026')

WEIGHT_PRO = 0.7          # 专业评委总分权重
WEIGHT_PUB = 0.3          # 大众评委总分权重
PRO_JUDGES = ['张三', '李四', '王五', '赵六', '杨七']

DIMENSIONS = [
    {
        'key': 'english',
        'name': '一、商务英语表达',
        'max': 37,
        'items': [
            {'code': 'q1', 'label': '发音语调', 'desc': '发音清晰，语调自然，商务沟通节奏合适', 'max': 10},
            {'code': 'q2', 'label': '用词句式', 'desc': '外贸函电/谈判表达地道，术语准确（询盘、发盘、还盘、L/C、装运、索赔等）', 'max': 12},
            {'code': 'q3', 'label': '流畅度与应变', 'desc': '台词连贯、卡壳少，临场接话与应答自然', 'max': 10},
            {'code': 'q4', 'label': '语法准确性', 'desc': '语法错误少，不影响业务信息传递', 'max': 5},
        ],
    },
    {
        'key': 'business',
        'name': '二、外贸业务逻辑与专业度',
        'max': 28,
        'items': [
            {'code': 'q5', 'label': '场景匹配', 'desc': '剧本贴合所抽外贸场景，业务链条完整不跑偏', 'max': 10},
            {'code': 'q6', 'label': '业务要点', 'desc': '贸易条款、交付、付款、品质、索赔等核心规则无硬伤', 'max': 10},
            {'code': 'q7', 'label': '冲突与沟通', 'desc': '有真实商务矛盾点，沟通方式符合真实外贸谈判习惯', 'max': 8},
        ],
    },
    {
        'key': 'team',
        'name': '三、团队参与度',
        'max': 15,
        'items': [
            {'code': 'q8', 'label': '全员出镜', 'desc': '≥6 人，人人有有效台词，无人只站背景', 'max': 9},
            {'code': 'q9', 'label': '角色分工', 'desc': '角色贴合业务岗位（业务员、客户、QC、单证、物流等），分工均衡', 'max': 6},
        ],
    },
    {
        'key': 'stage',
        'name': '四、舞台呈现与现场表现力',
        'max': 15,
        'items': [
            {'code': 'q10', 'label': '互动状态', 'desc': '眼神交流、肢体自然，不生硬念稿，脱稿优先', 'max': 7},
            {'code': 'q11', 'label': '道具辅助', 'desc': 'PPT、单据、样品等道具恰当，服务业务场景', 'max': 4},
            {'code': 'q12', 'label': '时间把控', 'desc': '15 分钟内完成；超时或过短酌情扣分', 'max': 4},
        ],
    },
]

ITEMS = [item for dim in DIMENSIONS for item in dim['items']]
ITEM_MAP = {item['code']: item for item in ITEMS}
BASE_MAX = sum(item['max'] for item in ITEMS)

app = Flask(__name__)
app.secret_key = SECRET_KEY


# ---------------------------------------------------------------- 数据库
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS teams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            order_no INTEGER
        );
        CREATE TABLE IF NOT EXISTS scores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            team_id INTEGER NOT NULL,
            team_name TEXT NOT NULL,
            judge_name TEXT NOT NULL,
            judge_role TEXT NOT NULL,
            items TEXT NOT NULL,
            base_total REAL NOT NULL,
            adjustment REAL NOT NULL,
            final_total REAL NOT NULL,
            remark TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE (team_id, judge_name)
        );
        """
    )
    conn.commit()
    conn.close()


# ---------------------------------------------------------------- 工具函数
def norm_name(raw):
    """去掉所有空白（含全角空格），用于评委姓名比对与去重。"""
    return ''.join((raw or '').split()).replace('\u3000', '')


def judge_role(name):
    return 'pro' if norm_name(name) in PRO_JUDGES else 'public'


def role_label(role):
    return '专业评委' if role == 'pro' else '大众评委'


def now_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def parse_items(form):
    """读取并校验 12 个评分项，返回 (items_dict, error)。"""
    items = {}
    for item in ITEMS:
        raw = (form.get(item['code']) or '').strip()
        if raw == '':
            return None, '请完成「%s」的评分' % item['label']
        try:
            value = float(raw)
        except ValueError:
            return None, '「%s」的分数格式不正确' % item['label']
        if value < 0 or value > item['max']:
            return None, '「%s」的分数需在 0 ~ %s 之间' % (item['label'], item['max'])
        items[item['code']] = round(value, 2)
    return items, None


def summarize(rows):
    """按规则汇总单个队伍。

    rows: 该队伍的评分记录列表（sqlite3.Row）
    - 专业评委总分：去掉一个最高分和一个最低分后取平均（不足 3 人时不剔除，直接平均）
    - 大众评委总分：全部取平均
    - 合计总分：专业 * 0.7 + 大众 * 0.3
    """
    pro = sorted([r for r in rows if r['judge_role'] == 'pro'], key=lambda r: r['final_total'])
    pub = [r for r in rows if r['judge_role'] == 'public']

    dropped_ids = []
    kept = pro
    if len(pro) >= 3:
        dropped_ids = [pro[0]['id'], pro[-1]['id']]
        kept = pro[1:-1]

    pro_avg = round(mean([r['final_total'] for r in kept]), 2) if kept else None
    pub_avg = round(mean([r['final_total'] for r in pub]), 2) if pub else None

    if pro_avg is not None and pub_avg is not None:
        total = round(pro_avg * WEIGHT_PRO + pub_avg * WEIGHT_PUB, 2)
    elif pro_avg is not None:
        total = pro_avg
    elif pub_avg is not None:
        total = pub_avg
    else:
        total = None

    return {
        'pro_avg': pro_avg,
        'pub_avg': pub_avg,
        'total': total,
        'pro_count': len(pro),
        'pub_count': len(pub),
        'dropped_ids': dropped_ids,
    }


def build_results(conn):
    teams = conn.execute(
        'SELECT * FROM teams ORDER BY COALESCE(order_no, 99999), id'
    ).fetchall()
    scores = conn.execute('SELECT * FROM scores').fetchall()
    grouped = {}
    for row in scores:
        grouped.setdefault(row['team_id'], []).append(row)
    for rows in grouped.values():
        rows.sort(key=lambda r: (r['judge_role'] != 'pro', -r['final_total']))

    results = []
    for team in teams:
        rows = grouped.get(team['id'], [])
        summary = summarize(rows)
        results.append({
            'team': dict(team),
            'rows': [dict(r) for r in rows],
            'stat': summary,
        })

    ranked = sorted(
        results,
        key=lambda item: (item['stat']['total'] is None, -(item['stat']['total'] or 0)),
    )
    for index, item in enumerate(ranked, start=1):
        item['rank'] = index if item['stat']['total'] is not None else '-'
    return ranked


def lan_ip():
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(('8.8.8.8', 80))
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except OSError:
        return '127.0.0.1'


def public_root():
    return os.environ.get('PUBLIC_URL') or request.url_root


@app.template_filter('adj')
def fmt_adj(value):
    """加减分显示：正数带 +，零显示 0。"""
    num = float(value or 0)
    if abs(num) < 1e-9:
        return '0'
    return ('+%g' % num) if num > 0 else ('%g' % num)


# ---------------------------------------------------------------- 评委端
@app.route('/')
def index():
    conn = get_db()
    teams = conn.execute(
        'SELECT id, name, order_no FROM teams ORDER BY COALESCE(order_no, 99999), id'
    ).fetchall()
    conn.close()
    return render_template('index.html', teams=teams, pro_judges=PRO_JUDGES)


@app.route('/start', methods=['POST'])
def start():
    name = (request.form.get('judge_name') or '').strip()
    if not norm_name(name):
        flash('请填写评委姓名', 'error')
        return redirect(url_for('index'))

    team_id = request.form.get('team_id') or ''
    other_name = (request.form.get('team_other') or '').strip()
    conn = get_db()

    if team_id == 'other' or not team_id:
        if not other_name:
            conn.close()
            flash('请选择或填写参赛队伍', 'error')
            return redirect(url_for('index'))
        row = conn.execute('SELECT id FROM teams WHERE name = ?', (other_name,)).fetchone()
        if row:
            tid = row['id']
        else:
            order_no = conn.execute('SELECT COALESCE(MAX(order_no), 0) + 1 AS n FROM teams').fetchone()['n']
            cur = conn.execute('INSERT INTO teams (name, order_no) VALUES (?, ?)', (other_name, order_no))
            conn.commit()
            tid = cur.lastrowid
    else:
        row = conn.execute('SELECT id FROM teams WHERE id = ?', (team_id,)).fetchone()
        if not row:
            conn.close()
            flash('参赛队伍不存在，请重新选择', 'error')
            return redirect(url_for('index'))
        tid = row['id']

    conn.close()
    return redirect(url_for('score', team_id=tid, judge=norm_name(name)))


@app.route('/score')
def score():
    team_id = request.args.get('team_id', type=int)
    judge = norm_name(request.args.get('judge', ''))
    if not team_id or not judge:
        return redirect(url_for('index'))

    conn = get_db()
    team = conn.execute('SELECT * FROM teams WHERE id = ?', (team_id,)).fetchone()
    if not team:
        conn.close()
        flash('参赛队伍不存在，请重新选择', 'error')
        return redirect(url_for('index'))

    existing = conn.execute(
        'SELECT * FROM scores WHERE team_id = ? AND judge_name = ?', (team_id, judge)
    ).fetchone()
    conn.close()

    old_items = {}
    if existing:
        import json
        old_items = json.loads(existing['items'])

    return render_template(
        'score.html',
        team=team,
        judge=judge,
        role=judge_role(judge),
        role_text=role_label(judge_role(judge)),
        dimensions=DIMENSIONS,
        base_max=BASE_MAX,
        existing=existing,
        old_items=old_items,
    )


@app.route('/submit', methods=['POST'])
def submit():
    team_id = request.form.get('team_id', type=int)
    judge = norm_name(request.form.get('judge_name', ''))
    if not team_id or not judge:
        flash('参数不完整，请重新进入', 'error')
        return redirect(url_for('index'))

    items, error = parse_items(request.form)
    if error:
        flash(error, 'error')
        return redirect(url_for('score', team_id=team_id, judge=judge))

    try:
        adjustment = float((request.form.get('adjustment') or '0').strip() or 0)
    except ValueError:
        flash('浮动加减分格式不正确', 'error')
        return redirect(url_for('score', team_id=team_id, judge=judge))
    if adjustment < -5 or adjustment > 5:
        flash('浮动加减分需在 -5 ~ +5 之间', 'error')
        return redirect(url_for('score', team_id=team_id, judge=judge))

    base_total = round(sum(items.values()), 2)
    final_total = round(max(0.0, base_total + adjustment), 2)
    remark = (request.form.get('remark') or '').strip()

    import json
    conn = get_db()
    team = conn.execute('SELECT * FROM teams WHERE id = ?', (team_id,)).fetchone()
    if not team:
        conn.close()
        flash('参赛队伍不存在', 'error')
        return redirect(url_for('index'))

    exists = conn.execute(
        'SELECT id FROM scores WHERE team_id = ? AND judge_name = ?', (team_id, judge)
    ).fetchone()
    payload = (
        json.dumps(items, ensure_ascii=False),
        base_total,
        round(adjustment, 2),
        final_total,
        remark,
        now_str(),
        judge_role(judge),
        team['name'],
    )
    if exists:
        conn.execute(
            """UPDATE scores SET items = ?, base_total = ?, adjustment = ?, final_total = ?,
                   remark = ?, updated_at = ?, judge_role = ?, team_name = ? WHERE id = ?""",
            payload + (exists['id'],),
        )
    else:
        conn.execute(
            """INSERT INTO scores (items, base_total, adjustment, final_total, remark,
                   updated_at, judge_role, team_name, team_id, judge_name, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            payload + (team_id, judge, now_str()),
        )
    conn.commit()
    conn.close()

    return redirect(url_for('done', team_id=team_id, judge=judge))


@app.route('/done')
def done():
    team_id = request.args.get('team_id', type=int)
    judge = norm_name(request.args.get('judge', ''))
    conn = get_db()
    row = conn.execute(
        'SELECT * FROM scores WHERE team_id = ? AND judge_name = ?', (team_id, judge)
    ).fetchone()
    conn.close()
    if not row:
        return redirect(url_for('index'))
    return render_template(
        'done.html',
        row=row,
        role_text=role_label(row['judge_role']),
        team_id=team_id,
        judge=judge,
    )


# ---------------------------------------------------------------- 后台
def require_admin():
    if not session.get('is_admin'):
        return redirect(url_for('admin_login', next=request.path))
    return None


@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        if request.form.get('password', '') == ADMIN_PASSWORD:
            session['is_admin'] = True
            return redirect(request.args.get('next') or url_for('admin'))
        flash('密码不正确', 'error')
    return render_template('admin_login.html')


@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin', None)
    return redirect(url_for('admin_login'))


@app.route('/admin')
def admin():
    guard = require_admin()
    if guard:
        return guard
    conn = get_db()
    results = build_results(conn)
    teams = conn.execute('SELECT * FROM teams ORDER BY COALESCE(order_no, 99999), id').fetchall()
    total_scores = conn.execute('SELECT COUNT(*) AS c FROM scores').fetchone()['c']
    pro_scores = conn.execute("SELECT COUNT(*) AS c FROM scores WHERE judge_role = 'pro'").fetchone()['c']
    pub_scores = conn.execute("SELECT COUNT(*) AS c FROM scores WHERE judge_role = 'public'").fetchone()['c']
    conn.close()
    return render_template(
        'admin.html',
        results=results,
        teams=teams,
        total_scores=total_scores,
        pro_scores=pro_scores,
        pub_scores=pub_scores,
        entry_url=public_root(),
        lan_url='http://%s:%s/' % (lan_ip(), os.environ.get('PORT', '8000')),
        weight_pro=WEIGHT_PRO,
        weight_pub=WEIGHT_PUB,
    )


@app.route('/admin/team/<int:team_id>')
def admin_team(team_id):
    guard = require_admin()
    if guard:
        return guard
    conn = get_db()
    results = {item['team']['id']: item for item in build_results(conn)}
    conn.close()
    if team_id not in results:
        abort(404)
    detail = results[team_id]
    import json
    for row in detail['rows']:
        row['scores'] = json.loads(row['items'])
    return render_template(
        'team_detail.html',
        detail=detail,
        dimensions=DIMENSIONS,
        items=ITEMS,
        weight_pro=WEIGHT_PRO,
        weight_pub=WEIGHT_PUB,
    )


@app.route('/admin/teams/add', methods=['POST'])
def admin_add_team():
    guard = require_admin()
    if guard:
        return guard
    name = (request.form.get('name') or '').strip()
    order_no = request.form.get('order_no', type=int)
    if name:
        conn = get_db()
        try:
            conn.execute('INSERT INTO teams (name, order_no) VALUES (?, ?)', (name, order_no))
        except sqlite3.IntegrityError:
            flash('队伍「%s」已存在' % name, 'error')
        conn.commit()
        conn.close()
    return redirect(url_for('admin') + '#teams')


@app.route('/admin/teams/<int:team_id>/delete', methods=['POST'])
def admin_delete_team(team_id):
    guard = require_admin()
    if guard:
        return guard
    conn = get_db()
    conn.execute('DELETE FROM scores WHERE team_id = ?', (team_id,))
    conn.execute('DELETE FROM teams WHERE id = ?', (team_id,))
    conn.commit()
    conn.close()
    flash('队伍及其评分已删除', 'ok')
    return redirect(url_for('admin') + '#teams')


@app.route('/admin/reset', methods=['POST'])
def admin_reset():
    guard = require_admin()
    if guard:
        return guard
    conn = get_db()
    conn.execute('DELETE FROM scores')
    conn.commit()
    conn.close()
    flash('已清空全部评分记录', 'ok')
    return redirect(url_for('admin'))


@app.route('/admin/qr.png')
def admin_qr():
    guard = require_admin()
    if guard:
        return guard
    import qrcode
    img = qrcode.make(public_root(), box_size=8, border=2)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')


@app.route('/admin/export.csv')
def admin_export():
    guard = require_admin()
    if guard:
        return guard
    import csv
    import json
    conn = get_db()
    results = build_results(conn)
    conn.close()

    buf = io.StringIO()
    writer = csv.writer(buf)
    header = ['排名', '参赛队伍', '专业评委总分', '大众评委总分', '合计总分', '评委', '评委类型',
              '基础总分', '浮动加减分', '最终得分', '备注/亮点'] + [i['label'] for i in ITEMS]
    writer.writerow(header)
    for item in results:
        stat = item['stat']
        if not item['rows']:
            writer.writerow([item['rank'], item['team']['name'], '', '', '', '', '', '', '', '', ''] +
                            [''] * len(ITEMS))
        for row in item['rows']:
            scores = json.loads(row['items'])
            writer.writerow([
                item['rank'], item['team']['name'],
                '' if stat['pro_avg'] is None else stat['pro_avg'],
                '' if stat['pub_avg'] is None else stat['pub_avg'],
                '' if stat['total'] is None else stat['total'],
                row['judge_name'], role_label(row['judge_role']),
                row['base_total'], row['adjustment'], row['final_total'], row['remark'],
            ] + [scores.get(i['code'], '') for i in ITEMS])
    data = '\ufeff' + buf.getvalue()
    return send_file(
        io.BytesIO(data.encode('utf-8')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='商务英语情景剧大赛_成绩明细.csv',
    )


init_db()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8000'))
    print('评委入口: http://%s:%s/' % (lan_ip(), port))
    print('后台入口: http://%s:%s/admin' % (lan_ip(), port))
    app.run(host='0.0.0.0', port=port, debug=False)
