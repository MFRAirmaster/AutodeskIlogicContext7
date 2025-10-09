#!/usr/bin/env python3
"""
Conservative Batch Forum Scraper
Scrapes in small controlled batches for maximum reliability
Designed for long-running operations to reach 200+ pages
"""

import asyncio
import json
import re
import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
from urllib.parse import urljoin

from playwright.async_api import async_playwright
import aiofiles


class BatchForumScraper:
    """Conservative batch scraper for maximum reliability"""

    def __init__(self):
        self.base_url = "https://forums.autodesk.com"
        
        self.forum_sections = [
            "/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en",
            "/t5/inventor-programming-forum/bd-p/inventor-programming-api-forum-en",
            "/t5/inventor-programming-forum/bd-p/inventor-programming-interface-forum-en"
        ]

        # Output directories - define before loading scraped posts
        self.output_dir = "docs/examples/parallel-forum-scraped"
        self.code_dir = "docs/examples/code-snippets-parallel"
        self.db_path = "docs/examples/ilogic-snippets.db"
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.code_dir, exist_ok=True)

        self.scraped_posts = set()
        self.init_database()
        self.load_scraped_posts()
        
        self.stats = {
            'pages_scraped': 0,
            'posts_found': 0,
            'posts_scraped': 0,
            'with_code': 0,
            'errors': 0,
            'start_time': datetime.now()
        }

    def load_scraped_posts(self):
        """Load already scraped post URLs from database"""
        if os.path.exists(self.db_path):
            try:
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT url FROM posts")
                self.scraped_posts = set(row[0] for row in cursor.fetchall())
                conn.close()
                print(f"📚 Loaded {len(self.scraped_posts)} already-scraped posts")
            except Exception:
                pass

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
                print(f"❌ DB error: {e}")
                conn.rollback()
            finally:
                conn.close()

        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, _save)

    async def scrape_forum_page(self, page, forum_url: str, page_num: int) -> List[str]:
        """Scrape single forum page"""
        try:
            url = f"{self.base_url}{forum_url}?page={page_num}"
            await page.goto(url, wait_until='networkidle', timeout=45000)
            await page.wait_for_timeout(3000)

            # Extract post URLs
            post_links = await page.evaluate("""
                () => {
                    const links = Array.from(document.querySelectorAll('a[href*="/td-p/"]'));
                    return links.map(a => a.href).filter(h => h.includes('/td-p/'));
                }
            """)

            # Filter out already scraped
            new_links = [url for url in post_links if url not in self.scraped_posts]
            
            return list(set(new_links))

        except Exception as e:
            print(f"⚠️  Page {page_num} error: {str(e)[:50]}")
            return []

    async def scrape_post(self, page, url: str) -> Optional[Dict]:
        """Scrape individual post"""
        if url in self.scraped_posts:
            return None

        try:
            await page.goto(url, wait_until='networkidle', timeout=45000)
            await page.wait_for_timeout(2000)

            post_data = await page.evaluate("""
                () => {
                    const titleEl = document.querySelector('h1, .message-subject, .thread-subject');
                    const authorEl = document.querySelector('.lia-user-name-link, .UserName');
                    const dateEl = document.querySelector('.DateTime, .post-date');
                    const contentEl = document.querySelector('.lia-message-body-content, .post-content');

                    const codeBlocks = [];
                    const codeEls = document.querySelectorAll('pre, code, .code-block, .lia-code-sample');

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

    def categorize_post(self, post_data: Dict) -> str:
        """Categorize post"""
        combined = (
            post_data.get('title', '') + ' ' + 
            post_data.get('content', '') + ' ' +
            ' '.join([cb.get('code', '') for cb in post_data.get('codeBlocks', [])])
        ).lower()

        if any(kw in combined for kw in ['api', 'automation', 'assembly', 'geometry']):
            return 'advanced'
        if any(kw in combined for kw in ['error', 'problem', 'issue', 'fix']):
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

        # Save code files
        for i, code_block in enumerate(post_data.get('codeBlocks', [])):
            code_filepath = os.path.join(self.code_dir, f"{category}_{post_id}_{i}.vb")

            async with aiofiles.open(code_filepath, 'w', encoding='utf-8') as f:
                await f.write(f"' Title: {post_data.get('title', 'Unknown')}\n")
                await f.write(f"' URL: {post_data['url']}\n")
                await f.write(f"' Category: {category}\n\n")
                await f.write(code_block['code'])

        self.stats['with_code'] += 1

    async def scrape_batch(self, start_page: int, pages_per_batch: int):
        """Scrape a batch of pages"""
        print(f"\n{'='*60}")
        print(f"🔄 BATCH: Pages {start_page} to {start_page + pages_per_batch - 1}")
        print(f"{'='*60}")

        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(headless=True)

        try:
            # Collect post URLs
            all_urls = []
            
            for forum_url in self.forum_sections:
                section_name = forum_url.split('/')[-1]
                print(f"\n📁 Section: {section_name}")
                
                page = await browser.new_page()
                
                for page_num in range(start_page, start_page + pages_per_batch):
                    urls = await self.scrape_forum_page(page, forum_url, page_num)
                    all_urls.extend(urls)
                    self.stats['pages_scraped'] += 1
                    print(f"  Page {page_num}: {len(urls)} new posts")
                    await asyncio.sleep(2)  # Polite delay
                
                await page.close()

            print(f"\n📋 Total new posts to scrape: {len(all_urls)}")
            self.stats['posts_found'] += len(all_urls)

            # Scrape posts
            if all_urls:
                page = await browser.new_page()
                
                for i, url in enumerate(all_urls, 1):
                    post_data = await self.scrape_post(page, url)
                    if post_data:
                        await self.save_post_data(post_data)
                        print(f"  ✅ {i}/{len(all_urls)}: {post_data.get('title', '')[:50]}")
                    else:
                        print(f"  ⏭️  {i}/{len(all_urls)}: Skipped (no code or error)")
                    
                    self.stats['posts_scraped'] += 1
                    
                    # Small delay between posts
                    await asyncio.sleep(1)
                
                await page.close()

        finally:
            await browser.close()
            await playwright.stop()

        # Print batch summary
        duration = datetime.now() - self.stats['start_time']
        print(f"\n{'='*60}")
        print(f"✅ BATCH COMPLETE")
        print(f"{'='*60}")
        print(f"Posts with code this batch: {self.stats['with_code']}")
        print(f"Total posts scraped: {self.stats['posts_scraped']}")
        print(f"Total duration: {duration}")
        print()

    async def run(self, total_pages: int = 250, pages_per_batch: int = 10):
        """Run scraper in batches"""
        print(f"🚀 BATCH FORUM SCRAPER")
        print(f"{'='*60}")
        print(f"Target: {total_pages} pages per section")
        print(f"Batch size: {pages_per_batch} pages")
        print(f"Sections: {len(self.forum_sections)}")
        print(f"Already scraped: {len(self.scraped_posts)} posts")
        print()

        for batch_start in range(1, total_pages + 1, pages_per_batch):
            await self.scrape_batch(batch_start, pages_per_batch)
            
            # Save progress checkpoint
            print(f"💾 Checkpoint: {self.stats['with_code']} posts with code total")
            print()
            
            # Small break between batches
            await asyncio.sleep(5)

        # Final summary
        duration = datetime.now() - self.stats['start_time']
        print(f"\n{'='*60}")
        print(f"🎉 SCRAPING COMPLETE")
        print(f"{'='*60}")
        print(f"Pages scraped: {self.stats['pages_scraped']}")
        print(f"New posts found: {self.stats['posts_found']}")
        print(f"Posts scraped: {self.stats['posts_scraped']}")
        print(f"Posts with code: {self.stats['with_code']}")
        print(f"Errors: {self.stats['errors']}")
        print(f"Duration: {duration}")
        print(f"Database: {self.db_path}")
        print(f"{'='*60}")


async def main():
    scraper = BatchForumScraper()
    # Scrape 250 pages, 10 pages per batch for stability
    await scraper.run(total_pages=250, pages_per_batch=10)


if __name__ == "__main__":
    asyncio.run(main())
