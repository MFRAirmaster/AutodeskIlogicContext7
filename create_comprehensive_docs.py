import sqlite3
import os
from collections import defaultdict

def create_comprehensive_documentation():
    """Create comprehensive documentation from database"""
    
    conn = sqlite3.connect('docs/examples/ilogic-snippets.db')
    cursor = conn.cursor()
    
    # Get all posts with code
    cursor.execute('''
        SELECT p.post_id, p.title, p.author, p.date, p.content, p.url, p.category, p.code_blocks_count,
               GROUP_CONCAT(cs.code, '|||') as codes
        FROM posts p
        LEFT JOIN code_snippets cs ON p.post_id = cs.post_id
        WHERE p.code_blocks_count > 0
        GROUP BY p.post_id
        ORDER BY p.category, p.title
    ''')
    
    posts = cursor.fetchall()
    
    # Group by category
    categories = defaultdict(list)
    for post in posts:
        post_id, title, author, date, content, url, category, code_count, codes = post
        categories[category].append({
            'post_id': post_id,
            'title': title,
            'author': author,
            'date': date,
            'content': content,
            'url': url,
            'code_count': code_count,
            'codes': codes.split('|||') if codes else []
        })
    
    # Create master documentation file
    output_dir = 'docs/examples'
    os.makedirs(output_dir, exist_ok=True)
    
    # Create comprehensive guide
    with open(os.path.join(output_dir, 'comprehensive-ilogic-database.md'), 'w', encoding='utf-8') as f:
        f.write('# Comprehensive iLogic Code Database\n\n')
        f.write(f'**Total Posts:** {len(posts)}\n')
        f.write(f'**Total Code Snippets:** {sum(len(p["codes"]) for cat in categories.values() for p in cat)}\n\n')
        
        f.write('## Statistics by Category\n\n')
        for cat in sorted(categories.keys()):
            f.write(f'- **{cat.title()}**: {len(categories[cat])} posts\n')
        
        f.write('\n---\n\n')
        
        # Table of contents
        f.write('## Table of Contents\n\n')
        for cat in sorted(categories.keys()):
            f.write(f'- [{cat.title()} Examples (#{len(categories[cat])})]' + f'(#{cat}-examples)\n')
        
        f.write('\n---\n\n')
        
        # Write each category
        for cat in sorted(categories.keys()):
            f.write(f'## {cat.title()} Examples\n\n')
            f.write(f'Total: {len(categories[cat])} posts\n\n')
            
            for i, post in enumerate(categories[cat], 1):
                f.write(f'### {i}. {post["title"]}\n\n')
                f.write(f'**URL:** [{post["url"]}]({post["url"]})\n\n')
                
                if post['author']:
                    f.write(f'**Author:** {post["author"]}\n\n')
                
                if post['date']:
                    f.write(f'**Date:** {post["date"]}\n\n')
                
                if post['content']:
                    # Limit content preview
                    content_preview = post['content'][:300].replace('\n', ' ')
                    if len(post['content']) > 300:
                        content_preview += '...'
                    f.write(f'**Description:** {content_preview}\n\n')
                
                f.write(f'**Code Snippets ({post["code_count"]}):**\n\n')
                
                for j, code in enumerate(post['codes'], 1):
                    if code and code.strip():
                        f.write(f'#### Code Block {j}\n\n')
                        f.write('```vb\n')
                        f.write(code)
                        f.write('\n```\n\n')
                
                f.write('---\n\n')
    
    # Create category-specific files
    for cat in categories.keys():
        with open(os.path.join(output_dir, f'{cat}-examples-database.md'), 'w', encoding='utf-8') as f:
            f.write(f'# {cat.title()} iLogic Examples from Database\n\n')
            f.write(f'**Total Examples:** {len(categories[cat])}\n\n')
            f.write('---\n\n')
            
            for i, post in enumerate(categories[cat], 1):
                f.write(f'## {i}. {post["title"]}\n\n')
                f.write(f'**URL:** [{post["url"]}]({post["url"]})\n\n')
                
                if post['author']:
                    f.write(f'**Author:** {post["author"]}\n\n')
                
                if post['date']:
                    f.write(f'**Date:** {post["date"]}\n\n')
                
                if post['content']:
                    content_preview = post['content'][:400].replace('\n', ' ')
                    if len(post['content']) > 400:
                        content_preview += '...'
                    f.write(f'**Description:** {content_preview}\n\n')
                
                f.write(f'**Code Snippets ({post["code_count"]}):**\n\n')
                
                for j, code in enumerate(post['codes'], 1):
                    if code and code.strip():
                        f.write('```vb\n')
                        f.write(code)
                        f.write('\n```\n\n')
                
                f.write('---\n\n')
    
    # Create topic index
    create_topic_index(cursor, output_dir)
    
    conn.close()
    
    print(f'✅ Created comprehensive documentation')
    print(f'   - Master file: comprehensive-ilogic-database.md')
    print(f'   - Category files: {len(categories)} files')
    print(f'   - Topic index: topic-index.md')

def create_topic_index(cursor, output_dir):
    """Create index of common topics"""
    
    topics = {
        'Parameters': ['parameter', 'iproperties', 'custom property'],
        'Assembly': ['assembly', 'occurrence', 'component', 'constraint'],
        'Parts': ['part', 'feature', 'sketch', 'extrude', 'revolve'],
        'Drawing': ['drawing', 'view', 'sheet', 'dimension', 'annotation'],
        'Sheet Metal': ['sheet metal', 'flat pattern', 'bend', 'unfold'],
        'BOM': ['bom', 'bill of material', 'parts list'],
        'Export': ['export', 'dxf', 'pdf', 'dwg', 'savecopy'],
        'Forms': ['form', 'input', 'multi-value', 'user interface'],
        'Events': ['event', 'trigger', 'before save', 'after open'],
        'File Operations': ['file', 'directory', 'path', 'copy', 'move', 'delete']
    }
    
    with open(os.path.join(output_dir, 'topic-index.md'), 'w', encoding='utf-8') as f:
        f.write('# iLogic Topic Index\n\n')
        f.write('Find examples by topic or keyword.\n\n')
        f.write('---\n\n')
        
        for topic, keywords in topics.items():
            f.write(f'## {topic}\n\n')
            
            # Search for posts matching keywords
            keyword_query = ' OR '.join([f"LOWER(p.title || ' ' || p.content) LIKE '%{kw}%'" for kw in keywords])
            
            cursor.execute(f'''
                SELECT p.post_id, p.title, p.url, p.category, p.code_blocks_count
                FROM posts p
                WHERE ({keyword_query})
                AND p.code_blocks_count > 0
                ORDER BY p.category, p.title
                LIMIT 20
            ''')
            
            results = cursor.fetchall()
            
            if results:
                f.write(f'Found {len(results)} examples:\n\n')
                for post_id, title, url, category, code_count in results:
                    f.write(f'- [{title}]({url}) - *{category}* ({code_count} code snippets)\n')
                f.write('\n')
            else:
                f.write('No examples found.\n\n')
            
            f.write('---\n\n')

if __name__ == '__main__':
    create_comprehensive_documentation()
