# -*- coding: utf-8 -*-
"""
根据 URL 批量下载页面中的文件。

流程：
1. 登录 vendors.peerless-clothing.com 获取授权；
2. 访问用户输入的 URL，按页面上的总页数逐页(追加 &page=i)读取文件列表表格；
3. 筛选 Modified 列为日期输入框日期的前一天的文件，逐个下载；
4. 文件保存到指定目录下以日期输入框日期（YYMMDD）命名的文件夹中，已存在的文件不覆盖。
"""

import datetime
import os
import re
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from urllib.parse import unquote, urljoin

import requests
from bs4 import BeautifulSoup

BASE_URL = 'https://vendors.peerless-clothing.com'
LOGIN_URL = BASE_URL + '/index.php?dir='
LOGIN_USERNAME = 'motivesc'
LOGIN_PASSWORD = '5yhg3d'
DEFAULT_SAVE_DIR = r'\\192.168.0.3\01-业务一部资料\nina\123'

# 英文月份简写（避免依赖系统区域设置）
MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}

# 匹配形如 2026-Oct-19 09:10 AM 的日期
DATE_PATTERN = re.compile(
    r'(\d{4})\s*-\s*([A-Za-z]{3})\s*-\s*(\d{1,2})\s+(\d{1,2}):(\d{2})\s*([AaPp][Mm])'
)

# 匹配分页信息 Page << 1 >> of 2
PAGE_COUNT_PATTERN = re.compile(r'Page\s*<<\s*\d+\s*>>\s*of\s*(\d+)')


def parse_input_date(text):
    """将日期输入框的文本转换为 date，格式不正确时返回 None"""
    value = (text or '').strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y%m%d', '%Y年%m月%d日'):
        try:
            return datetime.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def build_page_url(url, page):
    """在原 URL 后面拼接页码参数"""
    separator = '&' if '?' in url else '?'
    return '%s%spage=%d' % (url, separator, page)


def get_filename_prefix(url):
    """取 URL 中倒数第二个 / 后（即最后一段目录名）的第一个字母作为文件名前缀，如 T_"""
    text = (url or '').strip().rstrip('/')
    last_slash = text.rfind('/')
    if last_slash < 0:
        return ''
    for char in text[last_slash + 1:]:
        if char.isalnum():
            return char.upper() + '_'
    return ''


def get_page_count(html):
    """读取页面中 Page << 1 >> of N 的 N，未找到时按 1 页处理"""
    text = BeautifulSoup(html, 'html.parser').get_text(' ', strip=True)
    match = PAGE_COUNT_PATTERN.search(text)
    if not match:
        return 1
    return max(1, int(match.group(1)))


def parse_modified(text):
    """将 Modified 列文本转换为 datetime，无法解析时返回 None"""
    match = DATE_PATTERN.search(text or '')
    if not match:
        return None
    year, month_abbr, day, hour, minute, ampm = match.groups()
    month = MONTH_MAP.get(month_abbr.lower())
    if month is None:
        return None
    hour = int(hour) % 12
    if ampm.upper() == 'PM':
        hour += 12
    try:
        return datetime.datetime(int(year), month, int(day), hour, int(minute))
    except ValueError:
        return None


def filename_from_disposition(disposition):
    """从 Content-Disposition 中提取文件名"""
    if not disposition:
        return None
    match = re.search(r"filename\*\s*=\s*([^;]+)", disposition, re.IGNORECASE)
    if match:
        value = match.group(1).strip().strip('"')
        if "''" in value:
            value = value.split("''", 1)[1]
        return unquote(value)
    match = re.search(r'filename\s*=\s*"?([^";]+)"?', disposition, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


def find_file_table(soup):
    """找到包含 File / Modified 列的表格（即文件列表表格）"""
    for table in soup.find_all('table'):
        headers = [th.get_text(strip=True).lower() for th in table.find_all('th')]
        if 'file' in headers and 'modified' in headers:
            return table
    return None


def collect_file_links(html, page_url, target_date):
    """解析页面，返回目标日期修改过的文件 [(文件名, 下载地址), ...]"""
    soup = BeautifulSoup(html, 'html.parser')
    table = find_file_table(soup)
    if table is None:
        raise RuntimeError('未找到文件列表表格，请检查输入的URL是否正确')

    rows = table.find_all('tr')
    if not rows:
        return []

    headers = [cell.get_text(' ', strip=True).lower() for cell in rows[0].find_all(['th', 'td'])]
    if 'file' not in headers or 'modified' not in headers:
        raise RuntimeError('文件列表表格中缺少 File 或 Modified 列')
    file_index = headers.index('file')
    modified_index = headers.index('modified')

    files = []
    for row in rows[1:]:
        cells = row.find_all('td')
        if len(cells) <= max(file_index, modified_index):
            continue
        link = cells[file_index].find('a', href=True)
        if link is None:
            continue
        href = link['href']
        # 跳过目录行 / 上级目录行
        if 'file=' not in href and (href.endswith('/') or 'dir=' in href):
            continue
        modified = parse_modified(cells[modified_index].get_text(' ', strip=True))
        if modified is None or modified.date() != target_date:
            continue
        name = os.path.basename(link.get_text(' ', strip=True).replace('\\', '/'))
        if not name:
            name = os.path.basename(unquote(href.split('file=', 1)[-1]))
        files.append((name, urljoin(page_url or BASE_URL, href)))
    return files


def login(session):
    """登录站点，失败时抛出异常"""
    session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    response = session.post(
        LOGIN_URL,
        data={'username': LOGIN_USERNAME, 'password': LOGIN_PASSWORD},
        timeout=30,
    )
    response.raise_for_status()
    if BeautifulSoup(response.text, 'html.parser').find('input', {'name': 'password'}):
        raise RuntimeError('授权登录失败，请检查用户名或密码')


def download_file(session, file_url, target_dir, fallback_name, prefix=''):
    """下载单个文件到目标目录，文件名前追加前缀，已存在时跳过并返回 None"""
    fallback_name = os.path.basename(str(fallback_name).replace('\\', '/')) or 'download'
    if os.path.exists(os.path.join(target_dir, prefix + fallback_name)):
        return None
    with session.get(file_url, timeout=300, stream=True) as response:
        response.raise_for_status()
        disposition = response.headers.get('Content-Disposition', '')
        content_type = response.headers.get('Content-Type', '')
        if 'text/html' in content_type.lower() and not disposition:
            raise RuntimeError('返回的是网页而不是文件，可能登录已失效')
        name = filename_from_disposition(disposition) or fallback_name
        name = os.path.basename(str(name).replace('\\', '/')) or fallback_name
        path = os.path.join(target_dir, prefix + name)
        if os.path.exists(path):
            return None
        with open(path, 'wb') as file_obj:
            for chunk in response.iter_content(chunk_size=64 * 1024):
                if chunk:
                    file_obj.write(chunk)
    return path


class DownloadApp(object):
    def __init__(self, root):
        self.root = root
        root.title('页面文件批量下载')
        root.geometry('800x230')
        root.minsize(640, 230)

        main = ttk.Frame(root, padding=16)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(1, weight=1)

        # 上部：URL 地址
        ttk.Label(main, text='URL地址：').grid(row=0, column=0, sticky=tk.W, pady=(0, 12))
        self.url_var = tk.StringVar()
        url_entry = ttk.Entry(main, textvariable=self.url_var)
        url_entry.grid(row=0, column=1, sticky=tk.EW, pady=(0, 12))

        # 中部：保存路径
        ttk.Label(main, text='保存路径：').grid(row=1, column=0, sticky=tk.W, pady=(0, 12))
        self.path_var = tk.StringVar(value=DEFAULT_SAVE_DIR)
        path_entry = ttk.Entry(main, textvariable=self.path_var)
        path_entry.grid(row=1, column=1, sticky=tk.EW, pady=(0, 12))

        # 下部：日期 + 下载按钮 + 状态
        bottom = ttk.Frame(main)
        bottom.grid(row=2, column=0, columnspan=2, sticky=tk.EW)
        ttk.Label(bottom, text='日期：').pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.date.today().strftime('%Y-%m-%d'))
        ttk.Entry(bottom, textvariable=self.date_var, width=14).pack(side=tk.LEFT)
        self.download_button = ttk.Button(
            bottom, text='下载', width=12, command=self.start_download)
        self.download_button.pack(side=tk.LEFT, padx=(12, 0))
        self.status_var = tk.StringVar(value='等待下载')
        ttk.Label(bottom, textvariable=self.status_var, foreground='#555555').pack(
            side=tk.LEFT, padx=(12, 0))

        url_entry.focus_set()

    def set_status(self, text):
        self.root.after(0, lambda: self.status_var.set(text))

    def start_download(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning('提示', '请输入URL地址')
            return
        if not url.lower().startswith(('http://', 'https://')):
            url = 'https://' + url
        save_dir = self.path_var.get().strip() or DEFAULT_SAVE_DIR
        reference_date = parse_input_date(self.date_var.get())
        if reference_date is None:
            messagebox.showwarning('提示', '日期格式不正确，请按 2026-11-11 的格式输入')
            return

        self.download_button.config(state=tk.DISABLED)
        self.status_var.set('正在登录...')
        threading.Thread(
            target=self.worker, args=(url, save_dir, reference_date), daemon=True).start()

    def worker(self, url, save_dir, reference_date):
        try:
            count, skipped, errors = self.run(url, save_dir, reference_date)
        except Exception as exc:
            message = str(exc) or exc.__class__.__name__
            self.root.after(0, lambda: self.finish_error(message))
            return
        self.root.after(0, lambda: self.finish(count, skipped, errors))

    def run(self, url, save_dir, reference_date):
        session = requests.Session()

        self.set_status('正在登录...')
        login(session)

        # 下载日期输入框日期前一天的文件，文件夹按日期输入框的日期命名
        target_date = reference_date - datetime.timedelta(days=1)
        session_dir = os.path.join(save_dir, reference_date.strftime('%y%m%d'))
        prefix = get_filename_prefix(url)

        self.set_status('正在读取文件列表...')
        response = session.get(build_page_url(url, 1), timeout=60)
        response.raise_for_status()
        page_count = get_page_count(response.text)

        count = 0
        skipped = 0
        errors = []
        for page in range(1, page_count + 1):
            if page == 1:
                html = response.text
            else:
                self.set_status('正在读取第 %d/%d 页...' % (page, page_count))
                response = session.get(build_page_url(url, page), timeout=60)
                response.raise_for_status()
                html = response.text

            files = collect_file_links(html, build_page_url(url, page), target_date)
            if not files:
                continue
            os.makedirs(session_dir, exist_ok=True)
            for index, (name, file_url) in enumerate(files, 1):
                self.set_status('第 %d/%d 页 %d/%d：%s'
                                % (page, page_count, index, len(files), name))
                try:
                    if download_file(session, file_url, session_dir, name, prefix) is None:
                        skipped += 1
                    else:
                        count += 1
                except Exception as exc:
                    errors.append('%s：%s' % (name, str(exc) or exc.__class__.__name__))
        return count, skipped, errors

    def finish(self, count, skipped, errors):
        self.download_button.config(state=tk.NORMAL)
        if count == 0 and skipped == 0 and not errors:
            self.status_var.set('没有找到需要下载的文件')
            messagebox.showinfo('提示', '没有找到修改日期为所选日期前一天的文件。')
            return
        if count == 0 and not skipped:
            self.status_var.set('下载失败')
            messagebox.showerror(
                '下载失败', '没有文件下载成功。\n\n' + '\n'.join(errors[:5]))
            return
        if errors:
            self.status_var.set('下载完成：成功 %d 个，跳过 %d 个，失败 %d 个'
                                % (count, skipped, len(errors)))
            messagebox.showwarning(
                '下载完成',
                '下载成功 %d 个文件，跳过已存在 %d 个，失败 %d 个。\n\n%s'
                % (count, skipped, len(errors), '\n'.join(errors[:5])))
            return
        self.status_var.set('下载成功，共 %d 个文件（跳过已存在 %d 个）' % (count, skipped))
        messagebox.showinfo(
            '下载成功',
            '下载成功！共下载 %d 个文件。%s'
            % (count, ('\n跳过已存在文件 %d 个。' % skipped) if skipped else ''))

    def finish_error(self, message):
        self.download_button.config(state=tk.NORMAL)
        self.status_var.set('下载失败')
        messagebox.showerror('下载失败', message)


if __name__ == '__main__':
    root = tk.Tk()
    DownloadApp(root)
    root.mainloop()
