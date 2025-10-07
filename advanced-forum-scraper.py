#!/usr/bin/env python3
"""
Advanced iLogic Forum Scraper
Automatically scrapes Autodesk Inventor Programming Forum for iLogic code examples
Uses Playwright for full JavaScript rendering and advanced scraping capabilities
"""

import asyncio
import json
import re
import os
from datetime import datetime
from typing import List, Dict, Optional
from urllib.parse import urljoin, urlparse

from playwright.async_api import async_playwright, Browser, Page, ElementHandle
import aiofiles


class AdvancedForumScraper:
    """Advanced scraper for Autodesk Inventor Programming Forum"""

    def __init__(self):
        self.base_url = "https://forums.autodesk.com"
        self.forum_url = "/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en"
        self.scraped_posts = set()
        self.results = []
        self.stats = {
            'total_scraped': 0,
            'with_code': 0,
            'errors': 0,
            'start_time': datetime.now()
        }

        # Output directories
        self.output_dir = "docs/examples/advanced-forum-scraped"
        self.code_dir = "docs/examples/code-snippets"
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.code_dir, exist_ok=True)

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

    async def create_page(self, browser: Browser):
        """Create a page with proper headers"""
        page = await browser.new_page(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        )

        # Set additional headers to avoid blocking
        await page.set_extra_http_headers({
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })

        return page

    async def scrape_forum_page(self, page: Page, page_num: int) -> List[str]:
        """Scrape a forum page for post URLs"""
        urls = []

        try:
            # Navigate to forum page
            forum_page_url = f"{self.base_url}{self.forum_url}?page={page_num}"
            print(f"📄 Scraping forum page {page_num}: {forum_page_url}")

            await page.goto(forum_page_url, wait_until='networkidle', timeout=30000)

            # Wait for content to load
            await page.wait_for_timeout(2000)

            # Look for post links using multiple selectors
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
            print(f"🔗 Found {len(urls)} post URLs on page {page_num}")

        except Exception as e:
            print(f"❌ Error scraping forum page {page_num}: {e}")

        return urls

    async def scrape_post_content(self, page: Page, url: str) -> Optional[Dict]:
        """Scrape individual post content with full JavaScript rendering"""
        if url in self.scraped_posts:
            return None

        try:
            print(f"🔍 Scraping post: {url}")
            await page.goto(url, wait_until='networkidle', timeout=30000)

            # Wait for dynamic content
            await page.wait_for_timeout(3000)

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
        """Save post data to files"""
        category = await self.categorize_post(post_data)
        post_id = re.search(r'/td-p/(\d+)', post_data['url'])
        if not post_id:
            return

        post_id = post_id.group(1)

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
            filename = f"{category}-examples-advanced.md"
            filepath = os.path.join(self.output_dir, filename)

            content = f"# Advanced {category.title()} iLogic Examples\n\n"
            content += f"Generated from {len(posts)} forum posts with advanced code examples.\n\n"
            content += f"**Generated:** {datetime.now().isoformat()}\n\n"
            content += "---\n\n"

            for post in posts:
                content += f"## {post.get('title', 'Unknown Title')}\n\n"
                content += f"**Source:** [{post['url']}]({post['url']})\n\n"
                content += f"**Author:** {post.get('author', 'Anonymous')}\n\n"
                content += f"**Date:** {post.get('date', 'Unknown')}\n\n"

                if post.get('content'):
                    # Truncate content if too long
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

    async def run(self, max_pages: int = 10, delay: float = 2.0):
        """Main scraping workflow"""
        print("🚀 Starting advanced forum scraping...")
        print(f"Target: {max_pages} pages")
        print(f"Delay between requests: {delay}s")
        print()

        browser = await self.init_browser()

        try:
            # Create new page with proper headers
            page = await self.create_page(browser)
            page.set_default_timeout(30000)

            # Scrape forum pages
            all_post_urls = []
            for page_num in range(1, max_pages + 1):
                page_urls = await self.scrape_forum_page(page, page_num)
                all_post_urls.extend(page_urls)

                if page_num < max_pages:
                    print(f"⏳ Waiting {delay}s before next page...")
                    await asyncio.sleep(delay)

            # Remove duplicates
            all_post_urls = list(set(all_post_urls))
            print(f"📋 Total unique post URLs found: {len(all_post_urls)}")

            # Scrape individual posts
            for i, url in enumerate(all_post_urls):
                print(f"📊 Progress: {i+1}/{len(all_post_urls)}")

                post_data = await self.scrape_post_content(page, url)
                if post_data:
                    await self.save_post_data(post_data)
                    self.stats['with_code'] += 1

                self.stats['total_scraped'] += 1

                # Delay between posts
                if i < len(all_post_urls) - 1:
                    await asyncio.sleep(delay)

            # Generate documentation
            await self.generate_documentation()

            # Print summary
            await self.print_summary()

        finally:
            await browser.close()

    async def print_summary(self):
        """Print scraping summary"""
        duration = datetime.now() - self.stats['start_time']

        print("\n📊 Advanced Scraping Summary")
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
    scraper = AdvancedForumScraper()

    # Run with default settings (10 pages, 2 second delay)
    await scraper.run(max_pages=10, delay=2.0)


if __name__ == "__main__":
    asyncio.run(main())
