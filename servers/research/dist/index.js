import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import * as cheerio from "cheerio";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const LOGINOM_CACHE_FILE = path.join(__dirname, "..", "company_cache.json");
async function getLoginomCache() {
    try {
        const raw = await fs.readFile(LOGINOM_CACHE_FILE, "utf-8");
        return JSON.parse(raw);
    }
    catch {
        return null;
    }
}
async function saveLoginomCache(data) {
    const cache = { fetched_at: new Date().toISOString(), data };
    await fs.writeFile(LOGINOM_CACHE_FILE, JSON.stringify(cache, null, 2));
}
// --- HTTP fetch helper ---
async function fetchHtml(url) {
    const res = await fetch(url, {
        headers: {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept-Language": "ru-RU,ru;q=0.9",
        },
        signal: AbortSignal.timeout(10_000),
    });
    if (!res.ok)
        throw new Error(`HTTP ${res.status} for ${url}`);
    return res.text();
}
function extractStructured(html) {
    const $ = cheerio.load(html);
    $("script, style, noscript, iframe, svg").remove();
    const seen = new Set();
    const lines = [];
    const add = (prefix, text) => {
        const clean = text.replaceAll(/\s+/g, " ").trim();
        if (clean.length < 8 || seen.has(clean))
            return;
        seen.add(clean);
        lines.push(`${prefix}${clean}`);
    };
    $("h1").each((_, el) => add("# ", $(el).text()));
    $("h2").each((_, el) => add("## ", $(el).text()));
    $("h3").each((_, el) => add("### ", $(el).text()));
    $("p").each((_, el) => add("", $(el).text()));
    $("li").each((_, el) => add("• ", $(el).text()));
    return lines.join("\n").slice(0, 6000);
}
// --- Server ---
const server = new McpServer({
    name: "research",
    version: "1.0.0",
});
// Tool 1: get links from a page
server.registerTool("get_page_links", {
    description: "Fetch a web page and return all internal links found on it. Use this to discover which pages of a site are worth reading.",
    inputSchema: {
        url: z.string().describe("URL of the page to fetch"),
    },
}, async ({ url }) => {
    const html = await fetchHtml(url);
    const $ = cheerio.load(html);
    const base = new URL(url);
    const links = [];
    $("a[href]").each((_, el) => {
        const href = $(el).attr("href") ?? "";
        try {
            const resolved = new URL(href, base);
            if (resolved.hostname === base.hostname) {
                links.push(resolved.href);
            }
        }
        catch {
            // ignore malformed hrefs
        }
    });
    const unique = [...new Set(links)].slice(0, 50);
    return {
        content: [{ type: "text", text: JSON.stringify(unique, null, 2) }],
    };
});
// Tool 2: fetch page text content
server.registerTool("fetch_page", {
    description: "Fetch a web page and return its visible text content. Use this to read about pages, services, and contacts.",
    inputSchema: {
        url: z.string().describe("URL of the page to fetch"),
    },
}, async ({ url }) => {
    const html = await fetchHtml(url);
    const $ = cheerio.load(html);
    $("script, style, noscript, nav, footer, header").remove();
    const text = extractStructured(html);
    return {
        content: [{ type: "text", text }],
    };
});
// Tool 3: get financial data from checko.ru by INN
server.registerTool("get_company_finances", {
    description: "Fetch financial data for a company from checko.ru using its INN (tax ID). " +
        "Find the INN first by reading the company's website with fetch_page, then call this tool.",
    inputSchema: {
        inn: z.string().describe("Company INN (10 or 12 digits)"),
    },
}, async ({ inn }) => {
    const companyUrl = `https://checko.ru/company/${inn}`;
    const companyHtml = await fetchHtml(companyUrl);
    const text = extractStructured(companyHtml);
    return {
        content: [
            {
                type: "text",
                text: `Source: ${companyUrl}\n\n${text}`,
            },
        ],
    };
});
// Tool 4: get Loginom company info (cached)
server.registerTool("get_loginom_info", {
    description: "Return information about Loginom Company (our company) from cache or by fetching loginom.ru. " +
        "Use this to understand what services we offer before writing a pitch.",
    inputSchema: {
        force_refresh: z
            .boolean()
            .optional()
            .describe("Set true to bypass cache and re-fetch loginom.ru"),
    },
}, async ({ force_refresh }) => {
    if (!force_refresh) {
        const cache = await getLoginomCache();
        if (cache) {
            return {
                content: [
                    {
                        type: "text",
                        text: `[Cached at ${cache.fetched_at}]\n\n${JSON.stringify(cache.data, null, 2)}`,
                    },
                ],
            };
        }
    }
    const BASE = "https://loginom.ru";
    // URL segments worth crawling — covers company info, product, cases, solutions, services
    const USEFUL_PATTERNS = [
        /^\/$/, // main page
        /^\/(platform|product)/,
        /^\/(solutions|blog|cases|client)/,
        /^\/(services|about|company)/,
        /^\/(partners|integrations|marketplace)/,
        /^\/(bank|finance|retail|logistic|medical)/,
        // pricing/editions excluded — causes confusion with custom dev project costs
    ];
    // Step 1: collect all internal links from the main page
    const mainHtml = await fetchHtml(BASE);
    const $main = cheerio.load(mainHtml);
    const discovered = new Set();
    $main("a[href]").each((_, el) => {
        const href = $main(el).attr("href") ?? "";
        try {
            const url = new URL(href, BASE);
            if (url.hostname === "loginom.ru") {
                discovered.add(url.pathname);
            }
        }
        catch {
            // ignore
        }
    });
    // Step 2: filter by useful patterns
    const toFetch = [
        BASE,
        ...[...discovered]
            .filter((p) => USEFUL_PATTERNS.some((re) => re.test(p)))
            .map((p) => `${BASE}${p}`),
    ].slice(0, 20); // cap at 20 pages
    // Step 3: fetch each page
    const results = {};
    for (const pageUrl of toFetch) {
        try {
            const html = await fetchHtml(pageUrl);
            results[pageUrl] = extractStructured(html);
        }
        catch (e) {
            results[pageUrl] = `Error: ${e}`;
        }
    }
    await saveLoginomCache(results);
    return {
        content: [
            {
                type: "text",
                text: `Fetched ${Object.keys(results).length} pages from loginom.ru\n\n${JSON.stringify(results, null, 2)}`,
            },
        ],
    };
});
server.registerPrompt("process_order_email", {
    description: "Full pipeline: parse client email, research the company, get Loginom info",
    argsSchema: {
        email: z.string().describe("Raw email text from the client"),
    },
}, ({ email }) => ({
    messages: [
        {
            role: "user",
            content: {
                type: "text",
                text: `You are an assistant for Loginom Company that processes incoming software development orders.

Follow these steps in order:
1. Call parse_order with all fields extracted from the email below
2. Call get_page_links on the client's website to discover pages
3. Call fetch_page on 2-3 relevant pages (about, services, contacts) to understand the business
4. Find the company's INN on their website, then call get_company_finances with it
5. Call get_loginom_info to load our company profile

After all steps, provide a brief summary: client overview, their financials, and whether Loginom should pursue this order.

Email:
${email}`,
            },
        },
    ],
}));
const transport = new StdioServerTransport();
await server.connect(transport);
