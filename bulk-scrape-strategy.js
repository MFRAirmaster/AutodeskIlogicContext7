#!/usr/bin/env node

/**
 * Bulk Forum Scraping Strategy for iLogic Documentation
 *
 * This script outlines the approach for scraping 500+ iLogic forum posts
 * to build comprehensive documentation.
 */

const fs = require('fs');
const path = require('path');

// Configuration
const CONFIG = {
  targetPosts: 500,
  batchSize: 10,
  delayBetweenBatches: 2000, // ms
  delayBetweenRequests: 1000, // ms
  outputDir: './docs/examples/forum-bulk/',
  categories: {
    basic: 'basic-examples.md',
    advanced: 'advanced-examples.md',
    api: 'api-examples.md',
    troubleshooting: 'troubleshooting-examples.md'
  }
};

// Known forum URLs to start with
const INITIAL_URLS = [
  // Recent posts (already scraped)
  "https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-measure-area-of-two-faces/td-p/13831748",
  "https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-needed-finish-feature-is-it-used-if-so-set-the-first/td-p/13808368",
  "https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725",
  "https://forums.autodesk.com/t5/inventor-programming-forum/controldefinition-execute2-does-not-run-unless-it-is-the-last/td-p/13838361",
  "https://forums.autodesk.com/t5/inventor-programming-forum/3d-sketch-line-between-two-ucs-center-points/td-p/13839610",
  "https://forums.autodesk.com/t5/inventor-programming-forum/selecting-sketch-profile-on-the-assembly/td-p/13841032",
  "https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-constraint-with-a-part-and-the-assembly-s-origin/td-p/13840352",
  "https://forums.autodesk.com/t5/inventor-programming-forum/cut-a-solid-by-revolvefeature/td-p/13837083",

  // Add more URLs from forum list scraping
  // This would be populated by scraping forum pages
];

class BulkForumScraper {
  constructor() {
    this.scrapedPosts = new Set();
    this.postsByCategory = {
      basic: [],
      advanced: [],
      api: [],
      troubleshooting: []
    };
    this.stats = {
      totalScraped: 0,
      withCode: 0,
      errors: 0,
      startTime: Date.now()
    };
  }

  /**
   * Main scraping workflow
   */
  async run() {
    console.log('🚀 Starting bulk forum scraping...');
    console.log(`Target: ${CONFIG.targetPosts} posts`);
    console.log(`Batch size: ${CONFIG.batchSize}`);
    console.log('');

    // Ensure output directory exists
    this.ensureOutputDir();

    // Start with known URLs
    const urlsToScrape = [...INITIAL_URLS];

    // Process in batches
    while (urlsToScrape.length > 0 && this.stats.totalScraped < CONFIG.targetPosts) {
      const batch = urlsToScrape.splice(0, CONFIG.batchSize);

      console.log(`📦 Processing batch ${Math.floor(this.stats.totalScraped / CONFIG.batchSize) + 1}`);
      console.log(`Queue: ${urlsToScrape.length} remaining`);

      await this.processBatch(batch);

      // Delay between batches
      if (urlsToScrape.length > 0) {
        console.log(`⏳ Waiting ${CONFIG.delayBetweenBatches}ms before next batch...`);
        await this.delay(CONFIG.delayBetweenBatches);
      }
    }

    // Generate final documentation
    await this.generateDocumentation();

    // Print summary
    this.printSummary();
  }

  /**
   * Process a batch of URLs
   */
  async processBatch(urls) {
    const promises = urls.map(url => this.scrapePost(url));
    const results = await Promise.allSettled(promises);

    for (const result of results) {
      if (result.status === 'fulfilled') {
        const postData = result.value;
        if (postData && postData.codeBlocks && postData.codeBlocks.length > 0) {
          this.categorizePost(postData);
          this.stats.withCode++;
        }
        this.stats.totalScraped++;
      } else {
        console.error('❌ Scraping error:', result.reason);
        this.stats.errors++;
      }
    }
  }

  /**
   * Scrape individual post (placeholder - would use MCP tool)
   */
  async scrapePost(url) {
    if (this.scrapedPosts.has(url)) {
      return null; // Already scraped
    }

    this.scrapedPosts.add(url);

    // Delay between requests
    await this.delay(CONFIG.delayBetweenRequests);

    // This would call the MCP tool
    // For now, return mock data structure
    return {
      url,
      title: 'Mock Post Title',
      content: 'Mock content',
      codeBlocks: [
        {
          code: 'Mock iLogic code',
          language: 'vb'
        }
      ],
      category: this.determineCategory(url)
    };
  }

  /**
   * Categorize post based on content
   */
  categorizePost(postData) {
    const category = this.determineCategory(postData.url);
    this.postsByCategory[category].push(postData);
  }

  /**
   * Determine category based on URL or content
   */
  determineCategory(url) {
    // Simple categorization logic
    if (url.includes('api') || url.includes('constraint') || url.includes('geometry')) {
      return 'api';
    } else if (url.includes('error') || url.includes('problem') || url.includes('issue')) {
      return 'troubleshooting';
    } else if (url.includes('advanced') || url.includes('complex') || url.includes('assembly')) {
      return 'advanced';
    } else {
      return 'basic';
    }
  }

  /**
   * Generate documentation files
   */
  async generateDocumentation() {
    console.log('📝 Generating documentation...');

    for (const [category, posts] of Object.entries(this.postsByCategory)) {
      const filename = CONFIG.categories[category];
      const filepath = path.join(CONFIG.outputDir, filename);

      const content = this.generateCategoryContent(category, posts);
      fs.writeFileSync(filepath, content, 'utf8');

      console.log(`✅ Generated ${filename} with ${posts.length} examples`);
    }
  }

  /**
   * Generate content for a category
   */
  generateCategoryContent(category, posts) {
    const title = category.charAt(0).toUpperCase() + category.slice(1) + ' iLogic Examples';

    let content = `# ${title}

This document contains ${category} iLogic examples scraped from Autodesk forums.

## Examples

`;

    posts.forEach((post, index) => {
      content += `### Example ${index + 1}: ${post.title}

**Source:** [Forum Post](${post.url})

**Code:**
\`\`\`vb
${post.codeBlocks[0]?.code || 'No code available'}
\`\`\`

---

`;
    });

    content += `
*Generated on: ${new Date().toISOString()}*
*Total examples: ${posts.length}*
`;

    return content;
  }

  /**
   * Ensure output directory exists
   */
  ensureOutputDir() {
    if (!fs.existsSync(CONFIG.outputDir)) {
      fs.mkdirSync(CONFIG.outputDir, { recursive: true });
    }
  }

  /**
   * Print scraping summary
   */
  printSummary() {
    const duration = Date.now() - this.stats.startTime;
    console.log('\n📊 Scraping Summary');
    console.log('==================');
    console.log(`Total posts scraped: ${this.stats.totalScraped}`);
    console.log(`Posts with code: ${this.stats.withCode}`);
    console.log(`Errors: ${this.stats.errors}`);
    console.log(`Duration: ${Math.round(duration / 1000)}s`);
    console.log(`Rate: ${Math.round(this.stats.totalScraped / (duration / 1000))} posts/sec`);

    console.log('\n📁 Categories:');
    Object.entries(this.postsByCategory).forEach(([cat, posts]) => {
      console.log(`  ${cat}: ${posts.length} examples`);
    });
  }

  /**
   * Utility delay function
   */
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Export for use as module
module.exports = BulkForumScraper;

// Run if called directly
if (require.main === module) {
  const scraper = new BulkForumScraper();
  scraper.run().catch(console.error);
}

console.log('Bulk Forum Scraper Strategy');
console.log('===========================');
console.log('');
console.log('This script provides a framework for bulk scraping iLogic forum posts.');
console.log('To actually scrape posts, use the MCP tool with individual URLs.');
console.log('');
console.log('Current status:');
console.log('- 8 posts manually scraped and documented');
console.log('- Documentation generated in docs/examples/');
console.log('- Ready for expansion to 500+ posts');
console.log('');
console.log('Next steps:');
console.log('1. Identify more forum post URLs');
console.log('2. Scrape in batches using MCP tool');
console.log('3. Categorize and document examples');
console.log('4. Update main documentation index');
