#!/usr/bin/env python3
"""
Export all code snippets from database to files
Ensures all code from database is saved to individual files
"""

import sqlite3
import os


def export_snippets_to_files():
    """Export all code snippets from database to files"""
    db_path = "docs/examples/ilogic-snippets.db"
    code_dir = "docs/examples/code-snippets-parallel"
    
    if not os.path.exists(db_path):
        print("❌ Database not found!")
        return
    
    os.makedirs(code_dir, exist_ok=True)
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get all snippets with post info
    cursor.execute("""
        SELECT 
            cs.post_id,
            cs.snippet_index,
            cs.code,
            cs.category,
            p.title,
            p.url
        FROM code_snippets cs
        JOIN posts p ON cs.post_id = p.post_id
        ORDER BY cs.post_id, cs.snippet_index
    """)
    
    snippets = cursor.fetchall()
    
    print(f"🔄 Exporting {len(snippets)} code snippets to files...")
    print()
    
    exported = 0
    skipped = 0
    
    for post_id, snippet_index, code, category, title, url in snippets:
        # Create filename
        filename = f"{category}_{post_id}_{snippet_index}.vb"
        filepath = os.path.join(code_dir, filename)
        
        # Check if file already exists
        if os.path.exists(filepath):
            skipped += 1
            continue
        
        # Write code to file
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(f"' Title: {title}\n")
                f.write(f"' URL: {url}\n")
                f.write(f"' Category: {category}\n")
                f.write(f"' Post ID: {post_id}\n")
                f.write(f"' Snippet: {snippet_index}\n")
                f.write(f"\n")
                f.write(code)
            
            exported += 1
            
            if exported % 100 == 0:
                print(f"  ✅ Exported {exported} files...")
                
        except Exception as e:
            print(f"  ❌ Error exporting {filename}: {e}")
    
    print()
    print("=" * 60)
    print("📊 EXPORT SUMMARY")
    print("=" * 60)
    print(f"  Total Snippets in DB:    {len(snippets):,}")
    print(f"  Files Exported:          {exported:,}")
    print(f"  Files Skipped (existed): {skipped:,}")
    print(f"  Output Directory:        {code_dir}")
    print("=" * 60)
    
    conn.close()


if __name__ == "__main__":
    export_snippets_to_files()
