#!/usr/bin/env python3
"""
Parallel Advanced iLogic Forum Scraper
Uses multiple concurrent workers to scrape Autodesk Inventor Programming Forum
"""

import asyncio
import json
import re
import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor

from playwright.async_api import async_playwright, Browser, Page, BrowserContext
import aiofiles

class ParallelForumScraper:
    """Parallel scraper for Autodesk Inventor Programming Forum"""

    def __init__(self, max_workers: int = 5):  # Reduced default workers for better success rate
        self.base_url = "https://forums.autodesk.com"

        # Multiple forum sections to scrape
        self.forum_sections = [
            "/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en",  # iLogic
            "/t5/inventor-programming-forum/bd-p/inventor-programming-api-forum-en",     # API
            "/t5/inventor-programming-forum/bd-p/inventor-programming-interface-forum-en" # Interface
        ]

        self.scraped_posts = set()
        self.results = []
        self.stats = {
            'total_scraped': 0,
            'with_code': 0,
            'errors': 0,
            'start_time': datetime.now()
        }

        # Configuration
        self.max_workers = max_workers
        self.semaphore = asyncio.Semaphore(max_workers)  # Limit concurrent requests

        # Output directories
        self.output_dir = "docs/examples/parallel-forum-scraped"
        self.code_dir = "docs/examples/code-snippets-parallel"
        self.db_path = "docs/examples/ilogic-snippets.db"
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.code_dir, exist_ok=True)

        # Initialize database
        self.init_database()

    def init_database(self):
        """Initialize SQLite database for storing snippets"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Create posts table
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

        # Create code_snippets table
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

        # Create search index table for better querying
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS snippet_search (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                snippet_id INTEGER,
                search_text TEXT,
                FOREIGN KEY (snippet_id) REFERENCES code_snippets (id)
            )
        ''')

        conn.commit()
        conn.close()

    async def save_to_database(self, post_data: Dict, category: str, post_id: str):
        """Save post data to SQLite database"""
        def _save():
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            try:
                # Insert post data
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

                # Insert code snippets
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

                    # Get the snippet ID for search indexing
                    snippet_id = cursor.lastrowid

                    # Create search text for indexing
                    search_text = f"{post_data.get('title', '')} {post_data.get('content', '')} {code_block['code']}".lower()
                    cursor.execute('''
                        INSERT INTO snippet_search (snippet_id, search_text)
                        VALUES (?, ?)
                    ''', (snippet_id, search_text))

                conn.commit()

            except Exception as e:
                print(f"❌ Database error: {e}")
                conn.rollback()
            finally:
                conn.close()

        # Run database operations in thread pool to avoid blocking
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, _save)

    async def init_browser(self) -> Browser:
        """Initialize Playwright browser"""
        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-accelerated-2d-canvas',
                '--no-first-run',
                '--no-zygote',
                '--single-process',
                '--disable-gpu'
            ]
        )
        return browser

    async def create_context(self, browser: Browser) -> BrowserContext:
        """Create a browser context with proper headers"""
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            extra_http_headers={
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                'Accept-Encoding': 'gzip, deflate',
                'DNT': '1',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
            }
        )
        return context

    async def scrape_forum_page(self, context: BrowserContext, forum_url: str, page_num: int) -> List[str]:
        """Scrape a forum page for post URLs"""
        urls = []

        async with self.semaphore:  # Limit concurrent requests
            try:
                page = await context.new_page()

                # Navigate to forum page
                forum_page_url = f"{self.base_url}{forum_url}?page={page_num}"
                print(f"📄 Scraping forum page {page_num} ({forum_url}): {forum_page_url}")

                await page.goto(forum_page_url, wait_until='networkidle', timeout=30000)

                # Wait for content to load
                await page.wait_for_timeout(2000)

                # Look for post links
                post_selectors = [
                    'a[data-ntype="discussion-topic"]',
                    '.lia-link-navigation[data-ntype="discussion-topic"]',
                    'a[href*="/td-p/"]',
                    '.message-subject a',
                    '.thread-list .subject a'
                ]

                post_links = []
                for selector in post_selectors:
                    try:
                        elements = await page.query_selector_all(selector)
                        for element in elements:
                            href = await element.get_attribute('href')
                            if href and '/td-p/' in href:
                                full_url = urljoin(self.base_url, href)
                                if full_url not in self.scraped_posts:
                                    post_links.append(full_url)
                    except Exception as e:
                        print(f"⚠️  Selector {selector} failed: {e}")
                        continue

                # Remove duplicates
                urls = list(set(post_links))
                print(f"🔗 Found {len(urls)} post URLs on page {page_num} ({forum_url})")

            except Exception as e:
                print(f"❌ Error scraping forum page {page_num} ({forum_url}): {e}")
            finally:
                await page.close()

        return urls

    async def scrape_post_content(self, context: BrowserContext, url: str) -> Optional[Dict]:
        """Scrape individual post content with better rate limiting"""
        if url in self.scraped_posts:
            return None

        async with self.semaphore:  # Limit concurrent requests
            page = None
            try:
                page = await context.new_page()
                print(f"🔍 Scraping post: {url}")

                # Add random delay to avoid rate limiting (1-3 seconds)
                await page.wait_for_timeout(1000 + (hash(url) % 2000))

                await page.goto(url, wait_until='networkidle', timeout=45000)
                await page.wait_for_timeout(2000)  # Reduced wait time

                # Extract post data
                post_data = await page.evaluate("""
                    () => {
                        const titleElement = document.querySelector('h1, .message-subject, .thread-subject');
                        const title = titleElement ? titleElement.textContent.trim() : 'Unknown Title';

                        const authorElement = document.querySelector('.lia-user-name-link, .UserName, .user-name');
                        const author = authorElement ? authorElement.textContent.trim() : 'Anonymous';

                        const dateElement = document.querySelector('.DateTime, .post-date, .message-date');
                        const date = dateElement ? dateElement.textContent.trim() : '';

                        const contentElement = document.querySelector('.lia-message-body-content, .post-content, .message-body');
                        const content = contentElement ? contentElement.textContent.trim() : '';

                        const codeBlocks = [];
                        const codeElements = document.querySelectorAll('pre, code, .code-block, .lia-code-sample');

                        codeElements.forEach((element, index) => {
                            const code = element.textContent.trim();
                            if (code && code.length > 10) {
                                codeBlocks.push({
                                    index: index,
                                    code: code,
                                    language: 'vb'
                                });
                            }
                        });

                        return {
                            title: title,
                            author: author,
                            date: date,
                            content: content,
                            codeBlocks: codeBlocks,
                            url: window.location.href
                        };
                    }
                """)

                if post_data and post_data.get('codeBlocks') and len(post_data['codeBlocks']) > 0:
                    self.scraped_posts.add(url)
                    post_data['scraped_at'] = datetime.now().isoformat()
                    return post_data
                else:
                    print(f"⚠️  No code found in post: {url}")
                    return None

            except Exception as e:
                print(f"❌ Error scraping post {url}: {e}")
                self.stats['errors'] += 1
                return None
            finally:
                if page:
                    try:
                        await page.close()
                    except Exception:
                        pass  # Ignore cleanup errors

    async def categorize_post(self, post_data: Dict) -> str:
        """Categorize post based on content and code"""
        title = post_data.get('title', '').lower()
        content = post_data.get('content', '').lower()
        code = ' '.join([cb.get('code', '') for cb in post_data.get('codeBlocks', [])]).lower()

        # Advanced patterns
        if any(keyword in (title + content + code) for keyword in [
            'event trigger', 'propertyset', 'api', 'automation', 'advanced',
            'constraint', 'assembly', 'geometry', 'feature', 'pattern'
        ]):
            return 'advanced'

        # API specific
        if any(keyword in (title + content + code) for keyword in [
            'control definition', 'execute2', 'commandmanager', 'application',
            'document', 'component', 'feature', 'sketch'
        ]):
            return 'api'

        # Troubleshooting
        if any(keyword in (title + content + code) for keyword in [
            'error', 'problem', 'issue', 'fix', 'debug', 'troubleshoot'
        ]):
            return 'troubleshooting'

        return 'basic'

    async def save_post_data(self, post_data: Dict):
        """Save post data to files and database"""
        category = await self.categorize_post(post_data)
        post_id = re.search(r'/td-p/(\d+)', post_data['url'])
        if not post_id:
            return

        post_id = post_id.group(1)

        # Save to database
        await self.save_to_database(post_data, category, post_id)

        # Save individual code snippets
        for i, code_block in enumerate(post_data.get('codeBlocks', [])):
            code_filename = f"{category}_{post_id}_{i}.vb"
            code_filepath = os.path.join(self.code_dir, code_filename)

            async with aiofiles.open(code_filepath, 'w', encoding='utf-8') as f:
                await f.write(f"' Title: {post_data.get('title', 'Unknown')}\n")
                await f.write(f"' URL: {post_data['url']}\n")
                await f.write(f"' Category: {category}\n")
                await f.write(f"' Scraped: {post_data.get('scraped_at', datetime.now().isoformat())}\n\n")
                await f.write(code_block['code'])

        # Add to results
        post_data['category'] = category
        post_data['post_id'] = post_id
        self.results.append(post_data)

    async def generate_documentation(self):
        """Generate documentation from scraped data"""
        print("📝 Generating documentation...")

        # Group by category
        categories = {}
        for post in self.results:
            cat = post.get('category', 'uncategorized')
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(post)

        # Generate category files
        for category, posts in categories.items():
            filename = f"{category}-examples-parallel.md"
            filepath = os.path.join(self.output_dir, filename)

            content = f"# Parallel {category.title()} iLogic Examples\n\n"
            content += f"Generated from {len(posts)} forum posts with advanced parallel scraping.\n\n"
            content += f"**Generated:** {datetime.now().isoformat()}\n\n"
            content += "---\n\n"

            for post in posts:
                content += f"## {post.get('title', 'Unknown Title')}\n\n"
                content += f"**Source:** [{post['url']}]({post['url']})\n\n"
                content += f"**Author:** {post.get('author', 'Anonymous')}\n\n"
                content += f"**Date:** {post.get('date', 'Unknown')}\n\n"

                if post.get('content'):
                    preview = post['content'][:500]
                    if len(post['content']) > 500:
                        preview += "..."
                    content += f"**Description:** {preview}\n\n"

                content += "**Code:**\n\n"
                for i, code_block in enumerate(post.get('codeBlocks', [])):
                    content += f"```vb\n{code_block['code']}\n```\n\n"

                content += "---\n\n"

            async with aiofiles.open(filepath, 'w', encoding='utf-8') as f:
                await f.write(content)

            print(f"✅ Generated {filename} with {len(posts)} examples")

    async def scrape_forum_pages_parallel(self, context: BrowserContext, max_pages: int) -> List[str]:
        """Scrape multiple forum pages in parallel across all forum sections"""
        print(f"🚀 Starting parallel forum scraping with {self.max_workers} workers...")
        print(f"Forum sections: {len(self.forum_sections)}")

        # Create tasks for all forum sections and pages
        tasks = []
        for forum_url in self.forum_sections:
            for page_num in range(1, max_pages + 1):
                task = asyncio.create_task(self.scrape_forum_page(context, forum_url, page_num))
                tasks.append(task)

        # Wait for all tasks to complete
        results = await asyncio.gather(*tasks, return_exceptions=True)

        # Collect all URLs
        all_post_urls = []
        for result in results:
            if isinstance(result, list):
                all_post_urls.extend(result)

        # Remove duplicates
        all_post_urls = list(set(all_post_urls))
        print(f"📋 Total unique post URLs found across all sections: {len(all_post_urls)}")

        return all_post_urls

    async def scrape_posts_parallel(self, context: BrowserContext, post_urls: List[str]) -> None:
        """Scrape posts in parallel"""
        print(f"🔄 Scraping {len(post_urls)} posts with {self.max_workers} parallel workers...")

        # Create semaphore to limit concurrent requests
        semaphore = asyncio.Semaphore(self.max_workers)

        async def scrape_with_semaphore(url: str):
            async with semaphore:
                post_data = await self.scrape_post_content(context, url)
                if post_data:
                    await self.save_post_data(post_data)
                    self.stats['with_code'] += 1
                self.stats['total_scraped'] += 1
                return post_data

        # Create tasks for all posts
        tasks = []
        for url in post_urls:
            task = asyncio.create_task(scrape_with_semaphore(url))
            tasks.append(task)

        # Wait for all tasks to complete
        await asyncio.gather(*tasks, return_exceptions=True)

    async def run(self, max_pages: int = 10):
        """Main parallel scraping workflow"""
        print(f"🚀 Starting parallel forum scraping with {self.max_workers} workers...")
        print(f"Target: {max_pages} pages")
        print()

        browser = await self.init_browser()
        context = await self.create_context(browser)

        try:
            # Step 1: Scrape forum pages in parallel
            all_post_urls = await self.scrape_forum_pages_parallel(context, max_pages)

            # Step 2: Scrape individual posts in parallel
            await self.scrape_posts_parallel(context, all_post_urls)

            # Step 3: Generate documentation
            await self.generate_documentation()

            # Print summary
            await self.print_summary()

        finally:
            await context.close()
            await browser.close()

    async def print_summary(self):
        """Print scraping summary"""
        duration = datetime.now() - self.stats['start_time']

        print("\n📊 Parallel Scraping Summary")
        print("=" * 40)
        print(f"Total posts scraped: {self.stats['total_scraped']}")
        print(f"Posts with code: {self.stats['with_code']}")
        print(f"Errors: {self.stats['errors']}")
        print(f"Duration: {duration}")
        print(f"Rate: {self.stats['total_scraped'] / duration.total_seconds():.1f} posts/sec")
        print(f"Success rate: {(self.stats['with_code'] / max(1, self.stats['total_scraped'])) * 100:.1f}%")

        # Category breakdown
        categories = {}
        for post in self.results:
            cat = post.get('category', 'uncategorized')
            categories[cat] = categories.get(cat, 0) + 1

        print("\n📁 Categories:")
        for cat, count in categories.items():
            print(f"  {cat}: {count} examples")

async def main():
    """Main entry point"""
    # Create scraper with 5 parallel workers for better success rate
    scraper = ParallelForumScraper(max_workers=5)
    await scraper.run(max_pages=10)

if __name__ == "__main__":
    asyncio.run(main())
