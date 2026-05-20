import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ORDERS_FILE = path.join(__dirname, "..", "orders.json");
const VERDICTS_FILE = path.join(__dirname, "..", "verdicts.json");
async function loadOrders() {
    try {
        const data = await fs.readFile(ORDERS_FILE, "utf-8");
        return JSON.parse(data);
    }
    catch {
        return [];
    }
}
async function saveOrders(orders) {
    await fs.writeFile(ORDERS_FILE, JSON.stringify(orders, null, 2));
}
const server = new McpServer({
    name: "order-parser",
    version: "1.0.0",
});
server.registerTool("parse_order", {
    description: `Extract structured fields from a client email and save the order.
Read the email text carefully and populate all fields you can find.
Call this tool once per email with everything extracted.`,
    inputSchema: {
        company_name: z.string().describe("Name of the client company"),
        contact_name: z.string().describe("Full name of the contact person"),
        contact_email: z.string().describe("Contact email address"),
        contact_phone: z.string().optional().describe("Contact phone number"),
        description: z.string().describe("What software/system the client wants to build"),
        budget_from: z.number().optional().describe("Minimum budget in rubles"),
        budget_to: z.number().optional().describe("Maximum budget in rubles"),
        deadline: z.string().optional().describe("Project deadline or duration, e.g. '8 months'"),
        raw_email: z.string().optional().describe("Original email text verbatim"),
    },
}, async (args) => {
    const orders = await loadOrders();
    const order = {
        id: Date.now(),
        created_at: new Date().toISOString(),
        ...args,
    };
    orders.push(order);
    await saveOrders(orders);
    return {
        content: [
            {
                type: "text",
                text: JSON.stringify({ success: true, order }, null, 2),
            },
        ],
    };
});
server.registerTool("list_orders", {
    description: "Return all saved orders as JSON",
    inputSchema: {},
}, async () => {
    const orders = await loadOrders();
    return {
        content: [
            {
                type: "text",
                text: JSON.stringify(orders, null, 2),
            },
        ],
    };
});
server.registerTool("get_order", {
    description: "Return a single order by its id",
    inputSchema: {
        id: z.number().describe("Order id"),
    },
}, async ({ id }) => {
    const orders = await loadOrders();
    const order = orders.find((o) => o.id === id);
    if (!order) {
        return {
            content: [{ type: "text", text: `Order ${id} not found` }],
            isError: true,
        };
    }
    return {
        content: [{ type: "text", text: JSON.stringify(order, null, 2) }],
    };
});
server.registerTool("handle_email", {
    description: `ENTRY POINT. Call this first when you receive any client email.
Returns step-by-step instructions for the full processing pipeline.`,
    inputSchema: {
        email_text: z.string().describe("Raw email text from the client"),
    },
}, async ({ email_text }) => {
    const instructions = `You are an assistant for Loginom Company (loginom.ru).
Loginom Company is a SOFTWARE DEVELOPMENT company — they build custom analytics and automation systems for clients on a project basis.
They also develop their own platform product called "Loginom", but THIS IS IRRELEVANT to client orders.
IMPORTANT: Any pricing you see on loginom.ru (Enterprise, Standard, etc.) refers to self-service platform licenses — it has NOTHING to do with custom development project costs.
When evaluating a client order, assess only: project complexity, team fit, domain expertise, and timeline. Do not mention platform license prices.

Process this client email through the full pipeline:

STEP 1 — Call parse_order with all fields you can extract from the email.

STEP 2 — Call get_page_links on the client website URL found in the email to discover its pages.

STEP 3 — Call fetch_page on 2-3 relevant pages (look for: about, company, services, contacts in the URL). Find the company INN on these pages.

STEP 4 — Call get_company_finances with the INN you found.

STEP 5 — Call get_loginom_info to load our company profile.

STEP 6 — Call save_verdict with: decision, summary, fit_arguments, risks, solution_sketch, client_financials.

STEP 7 — If decision is "pursue", call generate_presentation using all data gathered. Fields:
- client_name, client_industry, order_description, budget, deadline, pilot_format
- problem_context: why this problem exists and why it's hard to solve
- pain_points: 2-3 consequences of not solving it
- solution_description: from solution_sketch
- fit_arguments: from save_verdict
- relevant_cases: 1-3 cases from loginom.ru that match the client's industry
- loginom_about: brief company description from loginom.ru cache
- contact_name, contact_email: from the email

EMAIL:
${email_text}`;
    return {
        content: [{ type: "text", text: instructions }],
    };
});
server.registerTool("save_verdict", {
    description: "Save the final assessment verdict for a client order. Call this after completing research and forming a recommendation.",
    inputSchema: {
        order_id: z.number().describe("ID of the order from parse_order"),
        decision: z.enum(["pursue", "decline", "pending"]).describe("Final decision"),
        summary: z.string().describe("2-3 sentence summary of the client and their task"),
        fit_arguments: z.array(z.string()).min(2).max(5).describe("Why Loginom fits this task"),
        risks: z.array(z.string()).optional().describe("Key risks or concerns"),
        solution_sketch: z.string().optional().describe("Brief description of the proposed technical solution"),
        client_financials: z.object({
            revenue: z.string().optional(),
            assets: z.string().optional(),
            rating: z.string().optional(),
            employees: z.string().optional(),
        }).optional().describe("Key financial figures found during research"),
    },
}, async (args) => {
    let verdicts = [];
    try {
        const raw = await fs.readFile(VERDICTS_FILE, "utf-8");
        verdicts = JSON.parse(raw);
    }
    catch {
        verdicts = [];
    }
    const verdict = { ...args, created_at: new Date().toISOString() };
    verdicts.push(verdict);
    await fs.writeFile(VERDICTS_FILE, JSON.stringify(verdicts, null, 2));
    return {
        content: [{ type: "text", text: JSON.stringify({ success: true, verdict }, null, 2) }],
    };
});
const transport = new StdioServerTransport();
await server.connect(transport);
