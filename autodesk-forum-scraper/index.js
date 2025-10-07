#!/usr/bin/env node

const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require('@modelcontextprotocol/sdk/types.js');
const puppeteer = require('puppeteer');

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
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        if (request.params.name === 'scrape_forum_post') {
          return await this.scrapeForumPost(request.params.arguments);
        } else if (request.params.name === 'scrape_forum_list') {
          return await this.scrapeForumList(request.params.arguments);
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

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Autodesk Forum Scraper MCP server running on stdio');
  }
}

const server = new AutodeskForumScraperServer();
server.run().catch(console.error);
