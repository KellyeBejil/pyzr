#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tungee CRM Crawler - Python GUI Version
基于 crawler.js 转换的 Python 版本，带有 GUI 界面
使用 conda 环境: pyzx
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import json
import os
import re
import asyncio
import sys
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl
from openpyxl import Workbook

# Playwright 导入
try:
    from playwright.async_api import async_playwright
except ImportError:
    print("请先安装 playwright: pip install playwright && playwright install chromium")
    exit(1)


class Config:
    """配置管理类"""
    CONFIG_FILE = "crawler_config.json"
    ENTERPRISE_IDS_FILE = "enterprise_ids.json"
    OUTPUT_FILE = "tungee_contacts_python.xlsx"
    
    def __init__(self):
        self.accounts = []
        self.load()
    
    def load(self):
        """加载配置"""
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.accounts = data.get('accounts', [])
            except Exception as e:
                print(f"加载配置失败: {e}")
                self.accounts = []
    
    def save(self):
        """保存配置"""
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({'accounts': self.accounts}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def add_account(self, user: str, password: str):
        """添加账号"""
        self.accounts.append({'user': user, 'pass': password})
        self.save()
    
    def remove_account(self, index: int):
        """删除账号"""
        if 0 <= index < len(self.accounts):
            self.accounts.pop(index)
            self.save()
    
    @staticmethod
    def load_enterprise_ids() -> list:
        """加载企业ID列表"""
        if os.path.exists(Config.ENTERPRISE_IDS_FILE):
            try:
                with open(Config.ENTERPRISE_IDS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return []
        return []
    
    @staticmethod
    def save_enterprise_ids(ids: list):
        """保存企业ID列表"""
        with open(Config.ENTERPRISE_IDS_FILE, 'w', encoding='utf-8') as f:
            json.dump(ids, f)
    
    @staticmethod
    def remove_enterprise_id(id_to_remove: str):
        """从列表中移除已处理的企业ID"""
        ids = Config.load_enterprise_ids()
        if id_to_remove in ids:
            ids.remove(id_to_remove)
            Config.save_enterprise_ids(ids)


class ExcelManager:
    """Excel 文件管理"""
    
    @staticmethod
    def save_to_excel(data: list, output_file: str = Config.OUTPUT_FILE):
        """保存数据到Excel，只保存有有效手机号的记录"""
        if not data:
            return
        
        # 过滤出有有效手机号的数据（11位数字）
        valid_data = []
        for row in data:
            phone = row.get('Phone', '')
            # 检查是否是有效的手机号（11位数字，以1开头）
            if phone and re.match(r'^1\d{10}$', str(phone)):
                valid_data.append(row)
        
        if not valid_data:
            return
        
        if os.path.exists(output_file):
            # 追加到现有文件
            wb = openpyxl.load_workbook(output_file)
            ws = wb.active
        else:
            # 创建新文件
            wb = Workbook()
            ws = wb.active
            ws.title = "Contacts"
            # 写入表头
            headers = ['Company', 'EnterpriseID', 'ContactName', 'Operator', 'Region', 'Phone']
            ws.append(headers)
        
        # 写入数据
        for row in valid_data:
            ws.append([
                row.get('Company', ''),
                row.get('EnterpriseID', ''),
                row.get('ContactName', ''),
                row.get('Operator', ''),
                row.get('Region', ''),
                row.get('Phone', '')
            ])
        
        wb.save(output_file)


class TungeeCrawler:
    """Tungee CRM 爬虫核心类"""
    
    def __init__(self, config: Config, log_callback=None, time_begin: int = None, time_end: int = None):
        self.config = config
        self.log_callback = log_callback
        self.current_account_index = 0
        self.is_running = False
        self.is_paused = False
        self.browser = None
        self.page = None
        self.playwright = None
        self.time_begin = time_begin  # 开始时间戳(毫秒)
        self.time_end = time_end      # 结束时间戳(毫秒)
        
    def log(self, message: str):
        """日志输出"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        print(log_message)
        if self.log_callback:
            self.log_callback(log_message)

    async def init_browser(self):
        """初始化浏览器"""
        # 如果是打包环境，设置浏览器路径
        if getattr(sys, 'frozen', False):
            # 获取临时解压目录
            base_path = sys._MEIPASS
            # 设置 PLAYWRIGHT_BROWSERS_PATH 环境变量指向解压后的 browsers 目录
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "browsers")
            self.log(f"运行在打包模式，浏览器路径: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")
        
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
        )
        self.page = await self.browser.new_page()
    
    async def close_browser(self):
        """关闭浏览器"""
        if self.page:
            await self.page.close()
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()
    
    async def clear_cookies_and_storage(self):
        """清除 Cookies 和存储"""
        if self.page:
            await self.page.context.clear_cookies()
            await self.page.evaluate("localStorage.clear(); sessionStorage.clear();")
    
    async def login(self, account: dict) -> bool:
        """登录"""
        self.log(f"开始登录: {account['user']}")
        
        try:
            await self.page.goto('https://user.tungee.com/users/sign-in', wait_until='networkidle')
            await asyncio.sleep(2)
            
            # 点击账号登录标签
            self.log("查找账号登录标签...")
            await self.page.evaluate("""
                () => {
                    const dom = document.getElementsByClassName('_2me12')[1];
                    if (dom) dom.click();
                }
            """)
            await asyncio.sleep(1)
            
            # 输入用户名
            self.log("输入账号密码...")
            user_input = await self.page.query_selector('input[placeholder*="手机"], input[placeholder*="账号"], input[type="text"]')
            pass_input = await self.page.query_selector('input[type="password"]')
            
            if user_input and pass_input:
                await user_input.click(click_count=3)
                await user_input.type(account['user'], delay=100)
                
                await pass_input.click(click_count=3)
                await pass_input.type(account['pass'], delay=100)
                
                # 点击登录按钮
                self.log("点击登录按钮...")
                success = await self.page.evaluate("""
                    () => {
                        const elements = [...document.querySelectorAll('button'), ...document.querySelectorAll('div')];
                        const btn = elements.find(el => 
                            el.innerText.trim() === '登录' && 
                            (el.tagName === 'BUTTON' || el.classList.contains('submit') || el.className.includes('btn'))
                        );
                        if (btn) {
                            btn.click();
                            return true;
                        }
                        const fallbackBtn = elements.find(el => el.innerText.trim() === '登录');
                        if (fallbackBtn) {
                            fallbackBtn.click();
                            return true;
                        }
                        return false;
                    }
                """)
                
                if not success:
                    await self.page.keyboard.press('Enter')
                
                await asyncio.sleep(5)
                
                current_url = self.page.url
                if 'sign-in' in current_url:
                    self.log("警告: 仍在登录页面，登录可能失败或需要验证码")
                    return False
                else:
                    self.log("登录成功!")
                    # 点击 "探迹CRM" 进入CRM系统
                    self.log("点击 探迹CRM...")
                    await asyncio.sleep(2)
                    try:
                        await self.page.click('span.LGMxM:has-text("探迹CRM")')
                        self.log("已进入探迹CRM")
                        await asyncio.sleep(2)
                    except Exception as e:
                        self.log(f"点击探迹CRM失败: {e}")
                    return True
            else:
                self.log("错误: 未找到登录输入框")
                return False
                
        except Exception as e:
            self.log(f"登录错误: {e}")
            return False
    
    async def switch_account(self) -> bool:
        """切换账号"""
        if not self.config.accounts:
            self.log("错误: 没有可用账号")
            return False
            
        self.log("!!! 切换账号 !!!")
        self.current_account_index = (self.current_account_index + 1) % len(self.config.accounts)
        next_account = self.config.accounts[self.current_account_index]
        self.log(f"切换到账号 {self.current_account_index}: {next_account['user']}")
        
        await self.clear_cookies_and_storage()
        await self.page.reload(wait_until='networkidle')
        
        login_success = await self.login(next_account)
        if not login_success:
            self.log("切换账号时登录失败!")
            return False
        
        # 重新导航到客户列表
        self.log("重新导航到客户列表...")
        await self.page.evaluate("""
            () => {
                const dom = document.getElementsByClassName('LGMxM')[1];
                if (dom) dom.click();
            }
        """)
        await asyncio.sleep(2)
        await self.page.goto('https://crm.tungee.com/customers/following', wait_until='networkidle')
        
        return True
    
    async def fetch_enterprise_ids(self) -> list:
        """获取企业ID列表 - 使用 unlock-leads-call API"""
        self.log("开始获取企业ID列表...")
        self.log(f"时间范围: {datetime.fromtimestamp(self.time_begin/1000)} 到 {datetime.fromtimestamp(self.time_end/1000)}")
        
        all_ids = []
        begin = 0
        batch_size = 200  # 该API每页50条
        next_search_after = None
        prev_search_after = None
        max_pages = 1000
        page_count = 0
        
        while page_count < max_pages and self.is_running:
            if self.is_paused:
                await asyncio.sleep(1)
                continue
                
            end = begin + batch_size
            self.log(f"获取ID: begin={begin}, end={end}")
            
            request_data = {
                'begin': begin,
                'end': end,
                'time_begin': self.time_begin,
                'time_end': self.time_end,
                'next_search_after': next_search_after,
                'prev_search_after': prev_search_after
            }
            
            try:
                result = await self.page.evaluate("""
                    async (data) => {
                        const params = new URLSearchParams();
                        params.append('sort_field', 'create_time');
                        params.append('sort', '-1');
                        params.append('create_time_begin', data.time_begin);
                        params.append('create_time_end', data.time_end);
                        params.append('begin', data.begin);
                        params.append('end', data.end);
                        
                        if (data.prev_search_after) {
                            params.append('prev_search_after', JSON.stringify(data.prev_search_after));
                        }
                        if (data.next_search_after) {
                            params.append('next_search_after', JSON.stringify(data.next_search_after));
                        }
                        
                        const url = `https://sales.tungee.com/api/unlock-leads-call?${params.toString()}`;
                        
                        try {
                            const response = await fetch(url, {
                                method: 'GET',
                                headers: {
                                    'Accept': '*/*',
                                    'Content-Type': 'application/json'
                                }
                            });
                            
                            if (!response.ok) {
                                return { error: `HTTP error! status: ${response.status}` };
                            }
                            
                            return await response.json();
                        } catch (e) {
                            return { error: e.toString() };
                        }
                    }
                """, request_data)
                
                if result.get('error'):
                    self.log(f"获取错误: {result['error']}")
                    break
                
                # API返回的数据在 leads 字段中
                leads = result.get('leads', [])
                if not leads:
                    self.log("没有更多数据")
                    break
                
                # 从每个 lead 中提取 enterprise_id
                new_ids = [item['enterprise_id'] for item in leads if item.get('enterprise_id')]
                all_ids.extend(new_ids)
                self.log(f"  获取 {len(new_ids)} 个ID，总计: {len(all_ids)}")
                
                next_search_after = result.get('next_search_after')
                prev_search_after = result.get('prev_search_after')
                
                # 如果没有下一页游标，说明已经到最后一页
                if not next_search_after:
                    self.log("已获取所有数据")
                    break
                
            except Exception as e:
                self.log(f"获取企业ID错误: {e}")
                break
            
            begin += batch_size
            page_count += 1
            await asyncio.sleep(1 + 1 * (asyncio.get_event_loop().time() % 1))
        
        # 保存到文件
        Config.save_enterprise_ids(all_ids)
        self.log(f"已保存 {len(all_ids)} 个ID到 {Config.ENTERPRISE_IDS_FILE}")
        
        return all_ids
    
    def check_night_mode(self) -> bool:
        """检查是否是夜间模式 (23:00 - 09:00)"""
        hour = datetime.now().hour
        return hour >= 23 or hour < 9
    
    def get_sleep_duration_until_9am(self) -> float:
        """获取到早上9点的睡眠时间(秒)"""
        now = datetime.now()
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if now.hour >= 23:
            target += timedelta(days=1)
        return (target - now).total_seconds()
    
    async def crawl_contacts(self, ids_to_process: list):
        """爬取联系人信息"""
        self.log(f"开始爬取 {len(ids_to_process)} 个企业的联系人信息...")
        
        # 设置响应监听
        async def handle_response(response):
            url = response.url
            if '/api/lead/contact/top-n-contacts' in url and hasattr(self, 'current_enterprise_id'):
                try:
                    data = await response.json()
                    results = []
                    company_name = data.get('lead', {}).get('enterprise_name', 'Unknown')
                    
                    top_contacts = data.get('top_contacts', [])
                    for contact in top_contacts:
                        phone = contact.get('contact_label', '')
                        # 只保留标准中国手机号
                        if re.match(r'^1[3-9]\d{9}$', phone):
                            results.append({
                                'Company': company_name,
                                'EnterpriseID': self.current_enterprise_id,
                                'ContactName': contact.get('contactName', ''),
                                'Operator': contact.get('contactOperator', ''),
                                'Region': contact.get('contactRegion', ''),
                                'Phone': phone
                            })
                    
                    if results:
                        ExcelManager.save_to_excel(results)
                        self.log(f"  [{self.current_enterprise_id}] 保存 {len(results)} 个联系人")
                        Config.remove_enterprise_id(self.current_enterprise_id)
                    else:
                        self.log(f"  [{self.current_enterprise_id}] 未找到联系人")
                        Config.remove_enterprise_id(self.current_enterprise_id)
                        
                except Exception as e:
                    self.log(f"  [{self.current_enterprise_id}] 解析响应错误: {e}")
        
        self.page.on('response', handle_response)
        
        for i, enterprise_id in enumerate(ids_to_process):
            if not self.is_running:
                break
            
            while self.is_paused and self.is_running:
                await asyncio.sleep(1)
            
            # 每10个请求休息一下
            if i > 0 and i % 10 == 0:
                import random
                sleep_time = random.randint(20, 40)
                self.log(f"已处理 {i} 个请求，休息 {sleep_time} 秒...")
                await asyncio.sleep(sleep_time)
            
            # 每100个请求切换账号
            if i > 0 and i % 100 == 0:
                self.log(f"已处理 {i} 个请求，切换账号...")
                await self.switch_account()
            
            # 夜间模式检查
            if self.check_night_mode():
                sleep_duration = self.get_sleep_duration_until_9am()
                self.log(f"夜间模式，暂停到早上9点 ({sleep_duration / 60:.0f} 分钟后)")
                await asyncio.sleep(sleep_duration)
                self.log("夜间休息结束，继续爬取...")
            
            self.current_enterprise_id = enterprise_id
            detail_url = f"https://sales.tungee.com/enterprise-details/{enterprise_id}/enterprise-information/basic-information"
            
            try:
                # 重试逻辑
                max_attempts = 3
                for attempt in range(max_attempts):
                    try:
                        if attempt > 0:
                            self.log(f"  [{enterprise_id}] 重试 {attempt + 1}/{max_attempts}...")
                        await self.page.goto(detail_url, wait_until='domcontentloaded', timeout=30000)
                        break
                    except Exception as nav_err:
                        self.log(f"  [{enterprise_id}] 尝试 {attempt + 1} 失败: {nav_err}")
                        if attempt == max_attempts - 1:
                            raise nav_err
                        await asyncio.sleep(2)
                
                # 等待API响应
                await asyncio.sleep(3)
                
            except Exception as e:
                self.log(f"  [{enterprise_id}] 导航错误: {e}")
        
        self.page.remove_listener('response', handle_response)
        self.log("爬取完成!")
    
    async def run(self, fetch_new_ids: bool = False):
        """主运行函数"""
        self.is_running = True
        self.is_paused = False
        
        try:
            self.log("版本 1.0.0 (Python)")
            
            if not self.config.accounts:
                self.log("错误: 请先配置账号!")
                return
            
            await self.init_browser()
            
            # 登录
            login_success = await self.login(self.config.accounts[self.current_account_index])
            if not login_success:
                self.log("登录失败，请检查账号密码或手动处理验证码")
                return
            
            await asyncio.sleep(2)
            
            # 点击导航
            await self.page.evaluate("""
                () => {
                    const dom = document.getElementsByClassName('LGMxM')[1];
                    if (dom) dom.click();
                }
            """)
            
            await asyncio.sleep(2)
            
            # 跳转到解锁列表页面
            self.log("跳转到解锁列表页面...")
            await self.page.goto('https://sales.tungee.com/unlock-list/enterprise')
            await asyncio.sleep(10)
            
            # 获取企业ID
            if fetch_new_ids:
                self.log("正在获取新的企业ID列表...")
                ids_to_process = await self.fetch_enterprise_ids()
            else:
                self.log("使用现有企业ID列表...")
                ids_to_process = Config.load_enterprise_ids()
            
            if not ids_to_process:
                self.log("没有企业ID可处理! 浏览器保持打开，可手动检查...")
                # 保持浏览器打开，等待用户手动关闭
                while self.is_running:
                    await asyncio.sleep(1)
            
            self.log(f"共 {len(ids_to_process)} 个企业ID待处理")
            
            # 爬取联系人
            await self.crawl_contacts(ids_to_process)
            
        except Exception as e:
            self.log(f"运行错误: {e}")
        finally:
            self.is_running = False
            await self.close_browser()
    
    def stop(self):
        """停止爬虫"""
        self.is_running = False
        self.is_paused = False
    
    def pause(self):
        """暂停爬虫"""
        self.is_paused = True
    
    def resume(self):
        """恢复爬虫"""
        self.is_paused = False


class CrawlerGUI:
    """爬虫GUI界面"""
    
    def __init__(self):
        self.config = Config()
        self.crawler = None
        self.crawler_thread = None
        
        self.root = tk.Tk()
        self.root.title("Tungee CRM 爬虫 - Python版")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        self.setup_ui()
        self.load_accounts_to_table()
    
    def setup_ui(self):
        """设置UI"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 账号配置区域
        account_frame = ttk.LabelFrame(main_frame, text="账号配置", padding="10")
        account_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 账号表格
        columns = ('用户名', '密码')
        self.account_tree = ttk.Treeview(account_frame, columns=columns, show='headings', height=3)
        self.account_tree.heading('用户名', text='用户名')
        self.account_tree.heading('密码', text='密码')
        self.account_tree.column('用户名', width=200)
        self.account_tree.column('密码', width=200)
        self.account_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 账号操作按钮
        btn_frame = ttk.Frame(account_frame)
        btn_frame.pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(btn_frame, text="添加账号", command=self.add_account).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="删除账号", command=self.remove_account).pack(fill=tk.X, pady=2)
        
        # 时间配置区域
        time_frame = ttk.LabelFrame(main_frame, text="时间范围配置", padding="10")
        time_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 开始时间
        ttk.Label(time_frame, text="开始日期:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.start_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.start_date_entry = ttk.Entry(time_frame, textvariable=self.start_date_var, width=15)
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(time_frame, text="(格式: YYYY-MM-DD)").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 结束时间
        ttk.Label(time_frame, text="结束日期:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.end_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.end_date_entry = ttk.Entry(time_frame, textvariable=self.end_date_var, width=15)
        self.end_date_entry.grid(row=0, column=4, padx=5, pady=5)
        ttk.Label(time_frame, text="(格式: YYYY-MM-DD)").grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        
        # 控制区域
        control_frame = ttk.LabelFrame(main_frame, text="控制", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 选项
        self.fetch_new_ids_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(control_frame, text="获取新的企业ID列表", variable=self.fetch_new_ids_var).pack(side=tk.LEFT)
        
        # 控制按钮
        btn_control_frame = ttk.Frame(control_frame)
        btn_control_frame.pack(side=tk.RIGHT)
        
        self.start_btn = ttk.Button(btn_control_frame, text="开始", command=self.start_crawler)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        
        self.pause_btn = ttk.Button(btn_control_frame, text="暂停", command=self.pause_crawler, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=2)
        
        self.stop_btn = ttk.Button(btn_control_frame, text="停止", command=self.stop_crawler, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
    def load_accounts_to_table(self):
        """加载账号到表格"""
        for item in self.account_tree.get_children():
            self.account_tree.delete(item)
        
        for acc in self.config.accounts:
            # 密码显示为星号
            masked_pass = '*' * len(acc['pass'])
            self.account_tree.insert('', tk.END, values=(acc['user'], masked_pass))
    
    def add_account(self):
        """添加账号对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加账号")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="用户名:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        user_entry = ttk.Entry(dialog, width=25)
        user_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="密码:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        pass_entry = ttk.Entry(dialog, width=25, show='*')
        pass_entry.grid(row=1, column=1, padx=10, pady=10)
        
        def save():
            user = user_entry.get().strip()
            password = pass_entry.get().strip()
            if user and password:
                self.config.add_account(user, password)
                self.load_accounts_to_table()
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "请输入用户名和密码")
        
        ttk.Button(dialog, text="保存", command=save).grid(row=2, column=0, columnspan=2, pady=20)
    
    def remove_account(self):
        """删除选中的账号"""
        selected = self.account_tree.selection()
        if selected:
            index = self.account_tree.index(selected[0])
            self.config.remove_account(index)
            self.load_accounts_to_table()
    
    def log(self, message: str):
        """添加日志"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
    
    def start_crawler(self):
        """启动爬虫"""
        if not self.config.accounts:
            messagebox.showwarning("警告", "请先添加账号!")
            return
        
        # 解析时间范围
        try:
            start_date_str = self.start_date_var.get().strip()
            end_date_str = self.end_date_var.get().strip()
            
            # 解析日期并转换为时间戳(毫秒)
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
            
            # 开始时间为当天00:00:00，结束时间为当天23:59:59
            time_begin = int(start_date.timestamp() * 1000)
            time_end = int((end_date.replace(hour=23, minute=59, second=59)).timestamp() * 1000)
            
            self.log(f"时间范围: {start_date_str} 00:00:00 到 {end_date_str} 23:59:59")
        except ValueError as e:
            messagebox.showwarning("警告", f"日期格式错误，请使用 YYYY-MM-DD 格式\n错误: {e}")
            return
        
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set("运行中...")
        
        self.crawler = TungeeCrawler(
            self.config, 
            log_callback=self.log,
            time_begin=time_begin,
            time_end=time_end
        )
        
        def run_async():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(self.crawler.run(fetch_new_ids=self.fetch_new_ids_var.get()))
            finally:
                loop.close()
                self.root.after(0, self.on_crawler_finished)
        
        self.crawler_thread = threading.Thread(target=run_async, daemon=True)
        self.crawler_thread.start()
    
    def pause_crawler(self):
        """暂停/恢复爬虫"""
        if self.crawler:
            if self.crawler.is_paused:
                self.crawler.resume()
                self.pause_btn.config(text="暂停")
                self.status_var.set("运行中...")
            else:
                self.crawler.pause()
                self.pause_btn.config(text="恢复")
                self.status_var.set("已暂停")
    
    def stop_crawler(self):
        """停止爬虫"""
        if self.crawler:
            self.crawler.stop()
            self.status_var.set("正在停止...")
    
    def on_crawler_finished(self):
        """爬虫结束回调"""
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text="暂停")
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")
    
    def run(self):
        """运行GUI"""
        self.root.mainloop()


def main():
    """主函数"""
    app = CrawlerGUI()
    app.run()


if __name__ == "__main__":
    main()
