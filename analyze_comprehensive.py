#!/usr/bin/env python3
"""
Comprehensive Database Analysis
Provides detailed insights into scraped iLogic code examples
"""

import sqlite3
import os
from collections import Counter
from datetime import datetime


def analyze_database():
    """Analyze the iLogic snippets database"""
    db_path = "docs/examples/ilogic-snippets.db"
    
    if not os.path.exists(db_path):
        print("❌ Database not found!")
        return
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    print("=" * 70)
    print("🔍 COMPREHENSIVE DATABASE ANALYSIS")
    print("=" * 70)
    print()
    
    # Basic stats
    cursor.execute("SELECT COUNT(*) FROM posts")
    total_posts = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM posts WHERE code_blocks_count > 0")
    posts_with_code = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM code_snippets")
    total_snippets = cursor.fetchone()[0]
    
    cursor.execute("SELECT SUM(code_blocks_count) FROM posts")
    total_code_blocks = cursor.fetchone()[0] or 0
    
    print("📊 OVERALL STATISTICS")
    print("-" * 70)
    print(f"  Total Posts in Database:        {total_posts:,}")
    print(f"  Posts with Code:                {posts_with_code:,}")
    print(f"  Total Code Snippets:            {total_snippets:,}")
    print(f"  Total Code Blocks (from posts): {total_code_blocks:,}")
    
    if total_posts > 0:
        print(f"  Average Code Blocks per Post:   {total_code_blocks/total_posts:.1f}")
        print(f"  Success Rate:                   {(posts_with_code/total_posts)*100:.1f}%")
    
    print()
    
    # Category breakdown
    cursor.execute("""
        SELECT category, COUNT(*), SUM(code_blocks_count) 
        FROM posts 
        GROUP BY category 
        ORDER BY COUNT(*) DESC
    """)
    categories = cursor.fetchall()
    
    print("📁 CATEGORY BREAKDOWN")
    print("-" * 70)
    for cat, count, blocks in categories:
        cat_name = cat if cat else "uncategorized"
        blocks = blocks or 0
        print(f"  {cat_name.ljust(20)}: {count:>4} posts | {blocks:>5} code blocks")
    
    print()
    
    # Code analysis
    cursor.execute("""
        SELECT code, category 
        FROM code_snippets 
        LIMIT 1000
    """)
    code_samples = cursor.fetchall()
    
    # Analyze code patterns
    keywords = Counter()
    for code, category in code_samples:
        code_lower = code.lower() if code else ""
        
        # Key iLogic/VB.NET keywords
        patterns = [
            'parameter', 'iproperties', 'messagebox', 'inputbox',
            'transaction', 'component', 'assembly', 'part',
            'sketch', 'feature', 'constraint', 'dimension',
            'workplane', 'drawing', 'view', 'sheet',
            'export', 'import', 'file.', 'directory',
            'suppressed', 'visible', 'occurrence', 'bom',
            'material', 'custom', 'iproperty', 'rule'
        ]
        
        for pattern in patterns:
            if pattern in code_lower:
                keywords[pattern] += 1
    
    print("🔑 TOP CODE PATTERNS (from 1000 samples)")
    print("-" * 70)
    for keyword, count in keywords.most_common(20):
        print(f"  {keyword.ljust(20)}: {count:>4} occurrences")
    
    print()
    
    # Size analysis
    cursor.execute("""
        SELECT 
            LENGTH(code) as code_length,
            category
        FROM code_snippets
        WHERE LENGTH(code) > 0
    """)
    
    code_lengths = [row[0] for row in cursor.fetchall()]
    
    if code_lengths:
        avg_length = sum(code_lengths) / len(code_lengths)
        min_length = min(code_lengths)
        max_length = max(code_lengths)
        
        print("📏 CODE SIZE ANALYSIS")
        print("-" * 70)
        print(f"  Average Code Length:     {avg_length:.0f} characters")
        print(f"  Shortest Code Snippet:   {min_length} characters")
        print(f"  Longest Code Snippet:    {max_length} characters")
        print()
        
        # Size distribution
        small = sum(1 for l in code_lengths if l < 200)
        medium = sum(1 for l in code_lengths if 200 <= l < 1000)
        large = sum(1 for l in code_lengths if 1000 <= l < 5000)
        xlarge = sum(1 for l in code_lengths if l >= 5000)
        
        print("  Size Distribution:")
        print(f"    Small   (< 200 chars):   {small:>4} snippets")
        print(f"    Medium  (200-1K chars):  {medium:>4} snippets")
        print(f"    Large   (1K-5K chars):   {large:>4} snippets")
        print(f"    X-Large (> 5K chars):    {xlarge:>4} snippets")
    
    print()
    
    # Language analysis
    cursor.execute("""
        SELECT language, COUNT(*) 
        FROM code_snippets 
        GROUP BY language
    """)
    languages = cursor.fetchall()
    
    print("💻 LANGUAGE BREAKDOWN")
    print("-" * 70)
    for lang, count in languages:
        lang_name = lang if lang else "unknown"
        print(f"  {lang_name.ljust(20)}: {count:>4} snippets")
    
    print()
    
    # Recent activity
    cursor.execute("""
        SELECT title, category, code_blocks_count, scraped_at
        FROM posts
        ORDER BY scraped_at DESC
        LIMIT 10
    """)
    recent = cursor.fetchall()
    
    print("🕐 MOST RECENTLY SCRAPED POSTS")
    print("-" * 70)
    for title, cat, blocks, scraped in recent:
        title_short = title[:50] + "..." if len(title) > 50 else title
        cat_name = cat if cat else "uncategorized"
        print(f"  [{cat_name}] {title_short}")
        print(f"          → {blocks} code blocks | {scraped}")
    
    print()
    
    # File system check
    code_dir = "docs/examples/code-snippets-parallel"
    if os.path.exists(code_dir):
        vb_files = [f for f in os.listdir(code_dir) if f.endswith('.vb')]
        print("💾 FILE SYSTEM")
        print("-" * 70)
        print(f"  Code Files Saved:        {len(vb_files):,}")
        print(f"  Database Snippets:       {total_snippets:,}")
        
        diff = abs(len(vb_files) - total_snippets)
        if diff > 0:
            print(f"  Difference:              {diff} snippets")
            if len(vb_files) < total_snippets:
                print(f"  Status:                  Some snippets not saved to files")
            else:
                print(f"  Status:                  More files than DB entries (duplicates?)")
        else:
            print(f"  Status:                  ✅ Perfect sync!")
    
    print()
    print("=" * 70)
    print("Analysis Complete!")
    print("=" * 70)
    
    conn.close()


if __name__ == "__main__":
    analyze_database()
