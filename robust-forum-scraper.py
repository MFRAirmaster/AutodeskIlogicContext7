#!/usr/bin/env python3
"""
Robust Forum Scraper for large scale scraping
Scrapes forum pages sequentially but posts in parallel for reliability
"""

import asyncio
import json
import re
import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
from urllib.parse import urljoin

from playwright.async_api import async_playwright, Browser, Page
import aiofiles


class RobustForumScraper:
    """Robust scraper for massive forum scraping operations"""

    def __init__(self, max_concurrent_posts: int = 3):
        self.base_url = "https://forums.autodesk.com"
        
        # Multiple forum sections
        self.forum_sections = [
            "/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en",
            "/t5/inventor-programming-forum/bd-p/inventor-programming-api-forum-en",
            "/t5/inventor-programming-forum/bd-p/inventor-programming-interface-forum-en"
        ]

        self.scraped_posts = set()
        self.results = []
        self.stats = {
            'total_pages': 0,
            'total_posts_found': 0,
            'total_scraped': 0,
            'with_code': 0,
            'errors': 0,
            'start_time': datetime.now()
        }

        # Limit concurrent post scraping
        self.max_concurrent_posts = max_concurrent_posts
        self.semaphore = asyncio.Semaphore(max_concurrent_posts)

        # Output directories
        self.output_dir = "docs/examples/parallel-forum-scraped"
        self.code_dir = "docs/examples/code-snippets-parallel"
        self.db_path = "docs/examples/ilogic-snippets.db"
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.code_dir, exist_ok=True)

        self.init_database()

    def init_database(self):
        """Initialize SQLite database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS posts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                post_id TEXT UNIQUE,
                title TEXT,
                author TEXT,
                date TEXT,
                content TEXT,
                url TEXT,
                category TEXT,
                scraped_at TEXT,
                code_blocks_count INTEGER
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS code_snippets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                post_id TEXT,
                snippet_index INTEGER,
                code TEXT,
                language TEXT,
                category TEXT,
                FOREIGN KEY (post_id) REFERENCES posts (post_id)
            )
        ''')

        conn.commit()
        conn.close()

    async def save_to_database(self, post_data: Dict, category: str, post_id: str):
        """Save post data to database"""
        def _save():
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            try:
                cursor.execute('''
                    INSERT OR REPLACE INTO posts
                    (post_id, title, author, date, content, url, category, scraped_at, code_blocks_count)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    post_id,
                    post_data.get('title', 'Unknown'),
                    post_data.get('author', 'Anonymous'),
                    post_data.get('date', ''),
                    post_data.get('content', ''),
                    post_data['url'],
                    category,
                    post_data.get('scraped_at', datetime.now().isoformat()),
                    len(post_data.get('codeBlocks', []))
                ))

                for i, code_block in enumerate(post_data.get('codeBlocks', [])):
                    cursor.execute('''
                        INSERT OR REPLACE INTO code_snippets
                        (post_id, snippet_index, code, language, category)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (
                        post_id,
                        i,
                        code_block['code'],
                        code_block.get('language', 'vb'),
                        category
                    ))

                conn.commit()
            except Exception as e:
                print(f"❌ Database error: {e}")
                conn.rollback()
            finally:
                conn.close()

        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, _save)

    async def scrape_forum_page(self, page: Page, forum_url: str, page_num: int) -> List[str]:
        """Scrape single forum page for post URLs"""
        urls = []
        
        try:
            forum_page_url = f"{self.base_url}{forum_url}?page={page_num}"
            print(f"📄 Page {page_num}: {forum_url.split('/')[-1]}")

            await page.goto(forum_page_url, wait_until='domcontentloaded', timeout=30000)
            await page.wait_for_timeout(2000)

            # Extract post links
            post_selectors = [
                'a[href*="/td-p/"]',
                'a[data-ntype="discussion-topic"]'
            ]

            for selector in post_selectors:
                try:
                    elements = await page.query_selector_all(selector)
                    for element in elements:
                        href = await element.get_attribute('href')
                        if href and '/td-p/' in href:
                            full_url = urljoin(self.base_url, href)
                            if full_url not in self.scraped_posts:
                                urls.append(full_url)
                except Exception:
                    continue

            urls = list(set(urls))
            print(f"   → Found {len(urls)} posts")
            return urls

        except Exception as e:
            print(f"   ❌ Error: {str(e)[:50]}")
            return []

    async def scrape_post(self, browser: Browser, url: str) -> Optional[Dict]:
        """Scrape individual post"""
        if url in self.scraped_posts:
            return None

        async with self.semaphore:
            page = None
            try:
                page = await browser.new_page()
                await page.goto(url, wait_until='domcontentloaded', timeout=30000)
                await page.wait_for_timeout(1500)

                post_data = await page.evaluate("""
                    () => {
                        const titleEl = document.querySelector('h1, .message-subject');
                        const authorEl = document.querySelector('.lia-user-name-link, .UserName');
                        const dateEl = document.querySelector('.DateTime, .post-date');
                        const contentEl = document.querySelector('.lia-message-body-content');

                        const codeBlocks = [];
                        const codeEls = document.querySelectorAll('pre, code, .code-block');

                        codeEls.forEach((el, idx) => {
                            const code = el.textContent.trim();
                            if (code && code.length > 10) {
                                codeBlocks.push({
                                    index: idx,
                                    code: code,
                                    language: 'vb'
                                });
                            }
                        });

                        return {
                            title: titleEl ? titleEl.textContent.trim() : 'Unknown',
                            author: authorEl ? authorEl.textContent.trim() : 'Anonymous',
                            date: dateEl ? dateEl.textContent.trim() : '',
                            content: contentEl ? contentEl.textContent.trim() : '',
                            codeBlocks: codeBlocks,
                            url: window.location.href
                        };
                    }
                """)

                if post_data and post_data.get('codeBlocks') and len(post_data['codeBlocks']) > 0:
                    self.scraped_posts.add(url)
                    post_data['scraped_at'] = datetime.now().isoformat()
                    return post_data

                return None

            except Exception as e:
                print(f"⚠️  Post error: {str(e)[:30]}")
                self.stats['errors'] += 1
                return None
            finally:
                if page:
                    try:
                        await page.close()
                    except Exception:
                        pass

    def categorize_post(self, post_data: Dict) -> str:
        """Categorize post"""
        title = post_data.get('title', '').lower()
        content = post_data.get('content', '').lower()
        code = ' '.join([cb.get('code', '') for cb in post_data.get('codeBlocks', [])]).lower()
        combined = title + content + code

        if any(kw in combined for kw in ['api', 'automation', 'assembly', 'geometry', 'feature']):
            return 'advanced'
        if any(kw in combined for kw in ['error', 'problem', 'issue', 'fix', 'debug']):
            return 'troubleshooting'
        return 'basic'

    async def save_post_data(self, post_data: Dict):
        """Save post data"""
        category = self.categorize_post(post_data)
        post_id = re.search(r'/td-p/(\d+)', post_data['url'])
        if not post_id:
            return

        post_id = post_id.group(1)

        await self.save_to_database(post_data, category, post_id)

        for i, code_block in enumerate(post_data.get('codeBlocks', [])):
            code_filename = f"{category}_{post_id}_{i}.vb"
            code_filepath = os.path.join(self.code_dir, code_filename)

            async with aiofiles.open(code_filepath, 'w', encoding='utf-8') as f:
                await f.write(f"' Title: {post_data.get('title', 'Unknown')}\n")
                await f.write(f"' URL: {post_data['url']}\n")
                await f.write(f"' Category: {category}\n\n")
                await f.write(code_block['code'])

        post_data['category'] = category
        post_data['post_id'] = post_id
        self.results.append(post_data)

    async def run(self, target_pages: int = 250):
        """Main scraping loop"""
        print(f"🚀 ROBUST FORUM SCRAPER")
        print(f"Target: {target_pages} pages per section")
        print(f"Concurrent posts: {self.max_concurrent_posts}")
        print(f"Sections: {len(self.forum_sections)}")
        print()

        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(headless=True)

        try:
            # Create a single page for forum browsing
            page = await browser.new_page()

            # Scrape each forum section
            for forum_url in self.forum_sections:
                print(f"\n📁 Section: {forum_url.split('/')[-1]}")
                
                all_post_urls = []
                
                # Scrape pages sequentially
                for page_num in range(1, target_pages + 1):
                    urls = await self.scrape_forum_page(page, forum_url, page_num)
                    all_post_urls.extend(urls)
                    self.stats['total_pages'] += 1

                    # Every 10 pages, scrape the collected posts
                    if page_num % 10 == 0 or page_num == target_pages:
                        print(f"\n🔄 Scraping {len(all_post_urls)} posts...")
                        self.stats['total_posts_found'] += len(all_post_urls)
                        
                        # Scrape posts in parallel batches
                        tasks = [self.scrape_post(browser, url) for url in all_post_urls]
                        results = await asyncio.gather(*tasks, return_exceptions=True)

                        for result in results:
                            if isinstance(result, dict):
                                await self.save_post_data(result)
                                self.stats['with_code'] += 1
                            self.stats['total_scraped'] += 1

                        all_post_urls = []  # Clear for next batch
                        
                        # Print progress
                        print(f"✅ Progress: {self.stats['with_code']} posts with code")
                        print()

            await page.close()

        finally:
            await browser.close()
            await playwright.stop()

        # Print final summary
        duration = datetime.now() - self.stats['start_time']
        print("\n" + "="*50)
        print("📊 FINAL SUMMARY")
        print("="*50)
        print(f"Pages scraped: {self.stats['total_pages']}")
        print(f"Posts found: {self.stats['total_posts_found']}")
        print(f"Posts with code: {self.stats['with_code']}")
        print(f"Errors: {self.stats['errors']}")
        print(f"Duration: {duration}")
        print(f"Database: {self.db_path}")
        print(f"Code files: {self.code_dir}")
        print("="*50)


async def main():
    scraper = RobustForumScraper(max_concurrent_posts=3)
    await scraper.run(target_pages=250)


if __name__ == "__main__":
    asyncio.run(main())
