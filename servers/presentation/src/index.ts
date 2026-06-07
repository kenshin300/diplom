import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { createRequire } from "node:module";
import { exec } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";

const execAsync = promisify(exec);
const SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
import {
  slide1Cover,
  slide2Task,
  slide3Solution,
  slide4WhyUs,
  slide5Cases,
  slide6About,
  slide7Contacts,
  type PresentationInput,
} from "./slides.js";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const PptxGenJS = createRequire(import.meta.url)("pptxgenjs");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.join(__dirname, "..", "output");

const server = new McpServer({ name: "presentation", version: "1.0.0" });

server.registerTool(
  "generate_presentation",
  {
    description:
      "Generate a PPTX pitch deck using data gathered from order-parser and research servers. " +
      "Call this after save_verdict. All content must come from research — do not invent data.",
    inputSchema: {
      client_name: z.string().describe("Full client company name"),
      client_industry: z.string().describe("Client industry, e.g. 'Розничная торговля'"),

      order_description: z.string().describe("What the client wants to build (from order)"),
      budget: z.string().describe("Budget range, e.g. '5–8 млн руб.'"),
      deadline: z.string().describe("Timeline, e.g. '10 месяцев'"),
      pilot_format: z.string().optional().describe("Pilot scope if mentioned, e.g. 'Пилот на одном регионе'"),

      problem_context: z.string().describe("Why this problem exists and why it's hard to solve"),
      pain_points: z.array(z.string()).min(1).max(3).describe("What happens if the problem is not solved"),

      solution_description: z.string().describe("Proposed solution architecture and approach — 3-5 sentences overview"),

      solution_phases: z.array(z.object({
        title: z.string().describe("Phase title, e.g. 'Интеграция источников данных'"),
        description: z.string().describe("What happens in this phase — 2-4 sentences"),
      })).min(2).max(4).describe("Key implementation phases of the proposed solution (2-4 steps)"),

      fit_arguments: z.array(z.string()).min(3).max(5).describe("Why Loginom specifically fits this task"),

      relevant_cases: z
        .array(z.object({
          title: z.string().describe("Client name or case title"),
          description: z.string().describe("What was built"),
          result: z.string().describe("Measurable outcome"),
        }))
        .min(1)
        .max(3)
        .describe("Relevant cases from loginom.ru that match the client's industry or task"),

      loginom_about: z.string().describe("Company description from loginom.ru cache — history, scale, expertise, key products"),

      loginom_stats: z.array(z.object({
        label: z.string().describe("Metric name, e.g. 'Лет на рынке', 'Клиентов', 'Специалистов'"),
        value: z.string().describe("Metric value with units, e.g. '20+', '500+', '200+'"),
      })).min(2).max(4).describe("Key Loginom company statistics from loginom.ru — years on market, client count, team size, project count"),

      contact_name: z.string().describe("Client contact person name"),
      contact_email: z.string().describe("Client contact email"),

      output_filename: z.string().optional().describe("Filename without extension"),
    },
  },
  async (args) => {
    const input: PresentationInput = {
      ...args,
      date: new Date().toLocaleDateString("ru-RU", { day: "numeric", month: "long", year: "numeric" }),
    };

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";

    slide1Cover(pptx, input);
    slide2Task(pptx, input);
    slide3Solution(pptx, input);
    slide4WhyUs(pptx, input);
    slide5Cases(pptx, input);
    slide6About(pptx, input);
    slide7Contacts(pptx, input);

    const filename = args.output_filename ?? `${args.client_name.replaceAll(/\s+/g, "_")}_pitch`;
    await fs.mkdir(OUTPUT_DIR, { recursive: true });
    const filepath = path.join(OUTPUT_DIR, `${filename}.pptx`);
    await pptx.writeFile({ fileName: filepath });

    // Convert to PNG slides via LibreOffice
    const pngDir = path.join(OUTPUT_DIR, filename);
    await fs.mkdir(pngDir, { recursive: true });
    await execAsync(`"${SOFFICE}" --headless --convert-to png --outdir "${pngDir}" "${filepath}"`);

    // Read generated PNGs sorted by name
    const allFiles = await fs.readdir(pngDir);
    const pngFiles = allFiles
      .filter((f) => f.endsWith(".png"))
      .sort()
      .map((f) => path.join(pngDir, f));

    const imageItems = await Promise.all(
      pngFiles.map(async (f) => {
        const buf = await fs.readFile(f);
        return {
          type: "image" as const,
          data: buf.toString("base64"),
          mimeType: "image/png" as const,
        };
      })
    );

    return {
      content: [
        {
          type: "text",
          text: `Презентация сгенерирована: ${filepath}\nСлайдов: ${imageItems.length}`,
        },
        ...imageItems,
      ],
    };
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
