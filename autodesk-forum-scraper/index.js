#!/usr/bin/env node

const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require('@modelcontextprotocol/sdk/types.js');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();

class AutodeskForumScraperServer {
  constructor() {
    this.server = new Server(
      {
        name: 'autodesk-forum-scraper',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupToolHandlers();
    
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'scrape_forum_post',
          description: 'Scrape an Autodesk forum post with full JavaScript rendering using Puppeteer',
          inputSchema: {
            type: 'object',
            properties: {
              url: {
                type: 'string',
                description: 'The URL of the forum post to scrape',
              },
              waitForSelector: {
                type: 'string',
                description: 'Optional CSS selector to wait for before extracting content',
                default: 'body',
              },
              extractCode: {
                type: 'boolean',
                description: 'Whether to specifically extract code blocks',
                default: true,
              },
            },
            required: ['url'],
          },
        },
        {
          name: 'scrape_forum_list',
          description: 'Scrape multiple forum post links from a list/search page',
          inputSchema: {
            type: 'object',
            properties: {
              url: {
                type: 'string',
                description: 'The URL of the forum list page',
              },
              maxPosts: {
                type: 'number',
                description: 'Maximum number of post links to extract',
                default: 10,
              },
              filterTag: {
                type: 'string',
                description: 'Optional tag to filter posts (e.g., "ilogic", "api")',
              },
            },
            required: ['url'],
          },
        },
        {
          name: 'search_local_snippets',
          description: 'Search for iLogic code snippets in the local database using keywords',
          inputSchema: {
            type: 'object',
            properties: {
              query: {
                type: 'string',
                description: 'Search query (keywords like "parameter", "feature", "assembly")',
              },
              category: {
                type: 'string',
                description: 'Optional category filter (basic, advanced, api, troubleshooting)',
              },
              limit: {
                type: 'number',
                description: 'Maximum number of results to return',
                default: 10,
              },
            },
            required: ['query'],
          },
        },
        {
          name: 'get_code_examples',
          description: 'Get code examples by category from the local database',
          inputSchema: {
            type: 'object',
            properties: {
              category: {
                type: 'string',
                description: 'Category to get examples for (basic, advanced, api, troubleshooting)',
              },
              limit: {
                type: 'number',
                description: 'Maximum number of examples to return',
                default: 20,
              },
            },
            required: ['category'],
          },
        },
        {
          name: 'get_related_examples',
          description: 'Find related code examples based on a code snippet or description',
          inputSchema: {
            type: 'object',
            properties: {
              code_snippet: {
                type: 'string',
                description: 'Code snippet to find related examples for',
              },
              limit: {
                type: 'number',
                description: 'Maximum number of related examples to return',
                default: 5,
              },
            },
            required: ['code_snippet'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        if (request.params.name === 'scrape_forum_post') {
          return await this.scrapeForumPost(request.params.arguments);
        } else if (request.params.name === 'scrape_forum_list') {
          return await this.scrapeForumList(request.params.arguments);
        } else if (request.params.name === 'search_local_snippets') {
          return await this.searchLocalSnippets(request.params.arguments);
        } else if (request.params.name === 'get_code_examples') {
          return await this.getCodeExamples(request.params.arguments);
        } else if (request.params.name === 'get_related_examples') {
          return await this.getRelatedExamples(request.params.arguments);
        } else {
          throw new Error(`Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `Error: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  async scrapeForumPost(args) {
    const { url, waitForSelector = 'body', extractCode = true } = args;
    
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    try {
      const page = await browser.newPage();
      
      // Set user agent to avoid bot detection
      await page.setUserAgent(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      );

      await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
      await page.waitForSelector(waitForSelector, { timeout: 10000 });

      // Extract post content
      const content = await page.evaluate((shouldExtractCode) => {
        const result = {
          title: '',
          author: '',
          date: '',
          content: '',
          codeBlocks: [],
          replies: [],
        };

        // Get title
        const titleEl = document.querySelector('h1, .lia-message-subject');
        if (titleEl) result.title = titleEl.textContent.trim();

        // Get main post content
        const postBody = document.querySelector('.lia-message-body-content, .message-body');
        if (postBody) {
          result.content = postBody.innerText.trim();
          
          // Extract code blocks if requested
          if (shouldExtractCode) {
            const codeElements = postBody.querySelectorAll('pre, code, .code-block');
            codeElements.forEach((el, idx) => {
              const code = el.textContent.trim();
              if (code.length > 10) { // Only include substantial code blocks
                result.codeBlocks.push({
                  index: idx,
                  code: code,
                  language: el.className.includes('vb') ? 'vb' : 
                           el.className.includes('csharp') ? 'csharp' : 'unknown',
                });
              }
            });
          }
        }

        // Get author
        const authorEl = document.querySelector('.lia-message-author-username, .author-name');
        if (authorEl) result.author = authorEl.textContent.trim();

        // Get date
        const dateEl = document.querySelector('.lia-message-post-date, .post-date');
        if (dateEl) result.date = dateEl.textContent.trim();

        // Get replies
        const replyElements = document.querySelectorAll('.lia-message-reply, .reply-message');
        replyElements.forEach((reply, idx) => {
          const replyBody = reply.querySelector('.lia-message-body-content, .message-body');
          const replyAuthor = reply.querySelector('.lia-message-author-username, .author-name');
          
          if (replyBody) {
            const replyObj = {
              index: idx,
              content: replyBody.innerText.trim().substring(0, 500), // Limit reply length
              author: replyAuthor ? replyAuthor.textContent.trim() : 'Unknown',
              codeBlocks: [],
            };

            if (shouldExtractCode) {
              const replyCodes = replyBody.querySelectorAll('pre, code');
              replyCodes.forEach((el, codeIdx) => {
                const code = el.textContent.trim();
                if (code.length > 10) {
                  replyObj.codeBlocks.push({
                    index: codeIdx,
                    code: code,
                  });
                }
              });
            }

            result.replies.push(replyObj);
          }
        });

        return result;
      }, extractCode);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(content, null, 2),
          },
        ],
      };
    } finally {
      await browser.close();
    }
  }

  async scrapeForumList(args) {
    const { url, maxPosts = 10, filterTag } = args;
    
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    try {
      const page = await browser.newPage();
      
      await page.setUserAgent(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      );

      await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
      await page.waitForSelector('body', { timeout: 10000 });

      // Extract post links
      const posts = await page.evaluate((max, tag) => {
        const result = [];
        const postLinks = document.querySelectorAll('a[href*="/td-p/"], a[href*="/m-p/"], .topic-link, .message-subject a');
        
        for (let i = 0; i < Math.min(postLinks.length, max); i++) {
          const link = postLinks[i];
          const href = link.href;
          const title = link.textContent.trim();
          
          // Check for tag filter if provided
          if (tag) {
            const parentRow = link.closest('.row, .topic-row, .message-row');
            if (parentRow) {
              const tagsText = parentRow.textContent.toLowerCase();
              if (!tagsText.includes(tag.toLowerCase())) {
                continue;
              }
            }
          }
          
          if (href && title) {
            result.push({
              title: title,
              url: href,
            });
          }
          
          if (result.length >= max) break;
        }
        
        return result;
      }, maxPosts, filterTag);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({ posts, total: posts.length }, null, 2),
          },
        ],
      };
    } finally {
      await browser.close();
    }
  }

  async searchLocalSnippets(args) {
    const { query, category, limit = 10 } = args;

    return new Promise((resolve, reject) => {
      const dbPath = path.join(__dirname, '..', 'docs', 'examples', 'ilogic-snippets.db');

      if (!fs.existsSync(dbPath)) {
        resolve({
          content: [
            {
              type: 'text',
              text: 'Database not found. Please run the scraper first to create the database.',
            },
          ],
        });
        return;
      }

      const db = new sqlite3.Database(dbPath);
      let sql = `
        SELECT p.post_id, p.title, p.author, p.content, p.category, p.url,
               cs.code, cs.language
        FROM posts p
        JOIN code_snippets cs ON p.post_id = cs.post_id
        WHERE (p.title LIKE ? OR p.content LIKE ? OR cs.code LIKE ?)
      `;
      const params = [`%${query}%`, `%${query}%`, `%${query}%`];

      if (category) {
        sql += ' AND p.category = ?';
        params.push(category);
      }

      sql += ' LIMIT ?';
      params.push(limit);

      db.all(sql, params, (err, rows) => {
        db.close();

        if (err) {
          reject(err);
          return;
        }

        const results = rows.map(row => ({
          post_id: row.post_id,
          title: row.title,
          author: row.author,
          category: row.category,
          url: row.url,
          code_preview: row.code.substring(0, 200) + (row.code.length > 200 ? '...' : ''),
          language: row.language,
        }));

        resolve({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                query: query,
                category: category || 'all',
                total_results: results.length,
                results: results,
              }, null, 2),
            },
          ],
        });
      });
    });
  }

  async getCodeExamples(args) {
    const { category, limit = 20 } = args;

    return new Promise((resolve, reject) => {
      const dbPath = path.join(__dirname, '..', 'docs', 'examples', 'ilogic-snippets.db');

      if (!fs.existsSync(dbPath)) {
        resolve({
          content: [
            {
              type: 'text',
              text: 'Database not found. Please run the scraper first to create the database.',
            },
          ],
        });
        return;
      }

      const db = new sqlite3.Database(dbPath);
      const sql = `
        SELECT p.post_id, p.title, p.author, p.content, p.category, p.url,
               cs.code, cs.language
        FROM posts p
        JOIN code_snippets cs ON p.post_id = cs.post_id
        WHERE p.category = ?
        ORDER BY p.scraped_at DESC
        LIMIT ?
      `;

      db.all(sql, [category, limit], (err, rows) => {
        db.close();

        if (err) {
          reject(err);
          return;
        }

        const results = rows.map(row => ({
          post_id: row.post_id,
          title: row.title,
          author: row.author,
          category: row.category,
          url: row.url,
          code: row.code,
          language: row.language,
        }));

        resolve({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                category: category,
                total_examples: results.length,
                examples: results,
              }, null, 2),
            },
          ],
        });
      });
    });
  }

  async getRelatedExamples(args) {
    const { code_snippet, limit = 5 } = args;

    return new Promise((resolve, reject) => {
      const dbPath = path.join(__dirname, '..', 'docs', 'examples', 'ilogic-snippets.db');

      if (!fs.existsSync(dbPath)) {
        resolve({
          content: [
            {
              type: 'text',
              text: 'Database not found. Please run the scraper first to create the database.',
            },
          ],
        });
        return;
      }

      const db = new sqlite3.Database(dbPath);

      // Extract keywords from the code snippet
      const keywords = code_snippet.toLowerCase()
        .split(/[^a-zA-Z0-9_]/)
        .filter(word => word.length > 3)
        .slice(0, 5); // Take first 5 keywords

      if (keywords.length === 0) {
        resolve({
          content: [
            {
              type: 'text',
              text: 'No meaningful keywords found in the code snippet.',
            },
          ],
        });
        return;
      }

      // Build search query
      const whereClause = keywords.map(() => 'cs.code LIKE ?').join(' OR ');
      const sql = `
        SELECT p.post_id, p.title, p.author, p.category, p.url,
               cs.code, cs.language
        FROM posts p
        JOIN code_snippets cs ON p.post_id = cs.post_id
        WHERE ${whereClause}
        ORDER BY p.scraped_at DESC
        LIMIT ?
      `;

      const params = [...keywords.map(k => `%${k}%`), limit];

      db.all(sql, params, (err, rows) => {
        db.close();

        if (err) {
          reject(err);
          return;
        }

        const results = rows.map(row => ({
          post_id: row.post_id,
          title: row.title,
          author: row.author,
          category: row.category,
          url: row.url,
          code_preview: row.code.substring(0, 300) + (row.code.length > 300 ? '...' : ''),
          language: row.language,
        }));

        resolve({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                original_keywords: keywords,
                total_related: results.length,
                related_examples: results,
            }, null, 2),
            },
          ],
        });
      });
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Autodesk Forum Scraper MCP server running on stdio');
  }
}

const server = new AutodeskForumScraperServer();
server.run().catch(console.error);
