#!/usr/bin/env python3
"""
Test script to debug forum scraping
"""

import asyncio
from playwright.async_api import async_playwright

async def test_forum_scraping():
    """Test forum page scraping"""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)  # Use headless mode
        page = await browser.new_page(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        )

        # Set additional headers
        await page.set_extra_http_headers({
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })

        try:
            # Navigate to main forum page first
            main_forum_url = "https://forums.autodesk.com/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en"
            print(f"Navigating to main forum: {main_forum_url}")

            await page.goto(main_forum_url, wait_until='networkidle', timeout=30000)
            await page.wait_for_timeout(3000)

            # Then try page 1
            forum_url = "https://forums.autodesk.com/t5/inventor-programming-forum/bd-p/inventor-programming-ilogic-forum-en?page=1"
            print(f"Navigating to page 1: {forum_url}")

            await page.goto(forum_url, wait_until='networkidle', timeout=30000)

            # Wait for content
            await page.wait_for_timeout(5000)

            # Get page title
            title = await page.title()
            print(f"Page title: {title}")

            # Get page content length
            content = await page.content()
            print(f"Page content length: {len(content)}")

            # Try different selectors
            selectors_to_try = [
                'a[href*="/td-p/"]',
                '.message-subject a',
                '.thread-list a',
                'a[data-ntype="discussion-topic"]',
                '.lia-link-navigation',
                'h2 a',
                '.topic-list a'
            ]

            for selector in selectors_to_try:
                try:
                    elements = await page.query_selector_all(selector)
                    print(f"Selector '{selector}': found {len(elements)} elements")

                    # Show first few hrefs
                    for i, element in enumerate(elements[:3]):
                        try:
                            href = await element.get_attribute('href')
                            text = await element.inner_text()
                            print(f"  {i+1}. {text[:50]}... -> {href}")
                        except Exception as e:
                            print(f"  {i+1}. Error getting element data: {e}")

                except Exception as e:
                    print(f"Error with selector '{selector}': {e}")

            # Try to find any links with td-p in them
            all_links = await page.query_selector_all('a')
            td_p_links = []

            for link in all_links:
                try:
                    href = await link.get_attribute('href')
                    if href and '/td-p/' in href:
                        text = await link.inner_text()
                        td_p_links.append((text[:50], href))
                except:
                    continue

            print(f"\nFound {len(td_p_links)} links containing '/td-p/':")
            for i, (text, href) in enumerate(td_p_links[:10]):
                print(f"  {i+1}. {text}... -> {href}")

            # Take screenshot for debugging
            await page.screenshot(path="forum_debug.png")
            print("Screenshot saved as forum_debug.png")

        except Exception as e:
            print(f"Error during test: {e}")

        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(test_forum_scraping())
