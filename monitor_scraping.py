#!/usr/bin/env python3
"""
Monitor the progress of forum scraping
Shows real-time statistics from the database
"""

import sqlite3
import os
import time
from datetime import datetime

def get_db_stats():
    """Get statistics from the database"""
    db_path = "docs/examples/ilogic-snippets.db"
    
    if not os.path.exists(db_path):
        return None
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get total posts
    cursor.execute("SELECT COUNT(*) FROM posts")
    total_posts = cursor.fetchone()[0]
    
    # Get posts with code
    cursor.execute("SELECT COUNT(*) FROM posts WHERE code_blocks_count > 0")
    posts_with_code = cursor.fetchone()[0]
    
    # Get total code snippets
    cursor.execute("SELECT COUNT(*) FROM code_snippets")
    total_snippets = cursor.fetchone()[0]
    
    # Get categories
    cursor.execute("SELECT category, COUNT(*) FROM posts GROUP BY category")
    categories = dict(cursor.fetchall())
    
    # Get recent posts
    cursor.execute("""
        SELECT title, scraped_at, code_blocks_count 
        FROM posts 
        ORDER BY scraped_at DESC 
        LIMIT 5
    """)
    recent_posts = cursor.fetchall()
    
    conn.close()
    
    return {
        'total_posts': total_posts,
        'posts_with_code': posts_with_code,
        'total_snippets': total_snippets,
        'categories': categories,
        'recent_posts': recent_posts
    }

def count_code_files():
    """Count code snippet files"""
    code_dir = "docs/examples/code-snippets-parallel"
    if not os.path.exists(code_dir):
        return 0
    return len([f for f in os.listdir(code_dir) if f.endswith('.vb')])

def print_progress():
    """Print current progress"""
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print("=" * 60)
    print("🔍 FORUM SCRAPING PROGRESS MONITOR")
    print("=" * 60)
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    stats = get_db_stats()
    
    if not stats:
        print("⏳ Waiting for database to be created...")
        print("   Scraping is starting up...")
        return
    
    # Overall stats
    print("📊 OVERALL STATISTICS")
    print("-" * 60)
    print(f"  Total Posts Scraped:    {stats['total_posts']:,}")
    print(f"  Posts with Code:        {stats['posts_with_code']:,}")
    print(f"  Total Code Snippets:    {stats['total_snippets']:,}")
    
    if stats['total_posts'] > 0:
        success_rate = (stats['posts_with_code'] / stats['total_posts']) * 100
        print(f"  Success Rate:           {success_rate:.1f}%")
    
    code_files = count_code_files()
    print(f"  Code Files Saved:       {code_files:,}")
    print()
    
    # Categories
    if stats['categories']:
        print("📁 CATEGORIES")
        print("-" * 60)
        for category, count in sorted(stats['categories'].items(), key=lambda x: x[1], reverse=True):
            print(f"  {category.ljust(20)}: {count:,} posts")
        print()
    
    # Recent posts
    if stats['recent_posts']:
        print("📝 RECENT POSTS (Last 5)")
        print("-" * 60)
        for title, scraped_at, code_count in stats['recent_posts']:
            # Truncate title if too long
            display_title = title[:50] + "..." if len(title) > 50 else title
            time_str = scraped_at.split('T')[1].split('.')[0] if 'T' in scraped_at else scraped_at
            print(f"  [{time_str}] {display_title}")
            print(f"              → {code_count} code block(s)")
        print()
    
    # Estimate remaining
    target = 250 * 57  # 250 pages * ~57 posts per page
    remaining = max(0, target - stats['total_posts'])
    
    print("🎯 PROGRESS TO TARGET")
    print("-" * 60)
    print(f"  Target Posts:           {target:,}")
    print(f"  Remaining:              {remaining:,}")
    
    if stats['total_posts'] > 0:
        progress = (stats['total_posts'] / target) * 100
        bar_length = 40
        filled = int((progress / 100) * bar_length)
        bar = "█" * filled + "░" * (bar_length - filled)
        print(f"  Progress:               [{bar}] {progress:.1f}%")
    
    print()
    print("Press Ctrl+C to stop monitoring (scraping will continue)")
    print("=" * 60)

def main():
    """Main monitoring loop"""
    try:
        while True:
            print_progress()
            time.sleep(5)  # Update every 5 seconds
    except KeyboardInterrupt:
        print("\n\n✅ Monitoring stopped. Scraping continues in background.")
        print("   Run this script again to resume monitoring.")

if __name__ == "__main__":
    main()
