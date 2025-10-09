import sqlite3

conn = sqlite3.connect('docs/examples/ilogic-snippets.db')
cursor = conn.cursor()

cursor.execute('SELECT COUNT(*) FROM posts')
print(f'Total posts: {cursor.fetchone()[0]}')

cursor.execute('SELECT COUNT(*) FROM code_snippets')
print(f'Total code snippets: {cursor.fetchone()[0]}')

cursor.execute('SELECT category, COUNT(*) FROM posts GROUP BY category')
print('\nPosts by category:')
for row in cursor.fetchall():
    print(f'  {row[0]}: {row[1]}')

# Get sample posts
cursor.execute('SELECT title, category, code_blocks_count FROM posts LIMIT 10')
print('\nSample posts:')
for row in cursor.fetchall():
    print(f'  [{row[1]}] {row[0]} - {row[2]} code blocks')

conn.close()
