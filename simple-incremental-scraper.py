#!/usr/bin/env python3
"""
Simple Incremental Scraper
Starts from a specific page to continue scraping where we left off
"""

import asyncio
import os
import sqlite3
import re
from datetime import datetime
from playwright.async_api import async_playwright
import aiofiles


async def scrape_and_save():
    """Scrape forum posts and save to database"""
    
    base_url = "https://forums.autodesk.com"
    forum_url = "/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en"
    db_path = "docs/examples/ilogic-snippets.db"
    code_dir = "docs/examples/code-snippets-parallel"
    
    # Load existing posts
    existing_urls = set()
    if os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute("SELECT url FROM posts")
        existing_urls = set(row[0] for row in c.fetchall())
        conn.close()
    
    print(f"🚀 Starting scraper")
    print(f"   Existing posts: {len(existing_urls)}")
    print()
    
    stats = {'new_posts': 0, 'with_code': 0}
    
    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch(headless=True)
    context = await browser.new_context()
    
    try:
        # Scrape pages 20-100 to get new content
        for page_num in range(20, 101):
            page_url = f"{base_url}{forum_url}?page={page_num}"
            print(f"📄 Scraping page {page_num}...")
            
            page = await context.new_page()
            
            try:
                await page.goto(page_url, wait_until='domcontentloaded', timeout=30000)
                await page.wait_for_timeout(2000)
                
                # Get post URLs
                post_urls = await page.evaluate("""
                    () => {
                        const links = Array.from(document.querySelectorAll('a[href*="/td-p/"]'));
                        return [...new Set(links.map(a => a.href))].filter(h => h.includes('/td-p/'));
                    }
                """)
                
                new_urls = [url for url in post_urls if url not in existing_urls]
                print(f"   Found {len(new_urls)} new posts")
                
                # Scrape each new post
                for post_url in new_urls:
                    try:
                        await page.goto(post_url, wait_until='domcontentloaded', timeout=30000)
                        await page.wait_for_timeout(1500)
                        
                        # Extract post data
                        post_data = await page.evaluate("""
                            () => {
                                const title = document.querySelector('h1, .message-subject')?.textContent.trim() || 'Unknown';
                                const author = document.querySelector('.lia-user-name-link')?.textContent.trim() || 'Anonymous';
                                const date = document.querySelector('.DateTime')?.textContent.trim() || '';
                                const content = document.querySelector('.lia-message-body-content')?.textContent.trim() || '';
                                
                                const codeBlocks = [];
                                document.querySelectorAll('pre, code, .lia-code-sample').forEach((el, idx) => {
                                    const code = el.textContent.trim();
                                    if (code && code.length > 10) {
                                        codeBlocks.push({ index: idx, code: code });
                                    }
                                });
                                
                                return { title, author, date, content, codeBlocks, url: window.location.href };
                            }
                        """)
                        
                        if post_data['codeBlocks'] and len(post_data['codeBlocks']) > 0:
                            # Save to database
                            post_id = re.search(r'/td-p/(\d+)', post_url)
                            if post_id:
                                post_id = post_id.group(1)
                                
                                # Determine category
                                combined = (post_data['title'] + ' ' + post_data['content']).lower()
                                if any(kw in combined for kw in ['api', 'automation', 'assembly']):
                                    category = 'advanced'
                                elif any(kw in combined for kw in ['error', 'problem', 'fix']):
                                    category = 'troubleshooting'
                                else:
                                    category = 'basic'
                                
                                # Save to DB
                                conn = sqlite3.connect(db_path)
                                c = conn.cursor()
                                c.execute('''
                                    INSERT OR IGNORE INTO posts
                                    (post_id, title, author, date, content, url, category, scraped_at, code_blocks_count)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (
                                    post_id, post_data['title'], post_data['author'],
                                    post_data['date'], post_data['content'], post_url,
                                    category, datetime.now().isoformat(),
                                    len(post_data['codeBlocks'])
                                ))
                                
                                for i, cb in enumerate(post_data['codeBlocks']):
                                    c.execute('''
                                        INSERT OR IGNORE INTO code_snippets
                                        (post_id, snippet_index, code, language, category)
                                        VALUES (?, ?, ?, ?, ?)
                                    ''', (post_id, i, cb['code'], 'vb', category))
                                    
                                    # Save code file
                                    filepath = os.path.join(code_dir, f"{category}_{post_id}_{i}.vb")
                                    async with aiofiles.open(filepath, 'w', encoding='utf-8') as f:
                                        await f.write(f"' {post_data['title']}\n")
                                        await f.write(f"' {post_url}\n\n")
                                        await f.write(cb['code'])
                                
                                conn.commit()
                                conn.close()
                                
                                existing_urls.add(post_url)
                                stats['with_code'] += 1
                                print(f"      ✅ {post_data['title'][:60]}")
                        
                        stats['new_posts'] += 1
                        await asyncio.sleep(0.5)
                        
                    except Exception as e:
                        print(f"      ⚠️  Error: {str(e)[:40]}")
                
            except Exception as e:
                print(f"   ❌ Page error: {str(e)[:50]}")
            finally:
                await page.close()
            
            # Status update every 10 pages
            if page_num % 10 == 0:
                print(f"\n📊 Progress: {stats['new_posts']} new posts, {stats['with_code']} with code\n")
            
            await asyncio.sleep(2)  # Polite delay between pages
    
    finally:
        await context.close()
        await browser.close()
        await playwright.stop()
    
    print(f"\n{'='*60}")
    print(f"✅ SCRAPING COMPLETE")
    print(f"{'='*60}")
    print(f"New posts found: {stats['new_posts']}")
    print(f"Posts with code: {stats['with_code']}")
    print(f"{'='*60}")


if __name__ == "__main__":
    asyncio.run(scrape_and_save())
