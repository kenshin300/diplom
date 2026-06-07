const C = {
  dark: "1E2433",      // Loginom dark navy
  darkAlt: "252B3B",   // slightly lighter navy
  red: "D4312A",       // Loginom red 
  white: "FFFFFF",
  light: "F4F5F7",     // light gray
  gray: "8A93A2",      
};

export interface RelevantCase {
  title: string;
  description: string;
  result: string;
}

export interface SolutionPhase {
  title: string;
  description: string;
}

export interface CompanyStat {
  label: string;
  value: string;
}

export interface PresentationInput {
  // Slide 1 — Cover
  client_name: string;
  client_industry: string;
  date: string;

  order_description: string;
  budget: string;
  deadline: string;
  pilot_format?: string;
  problem_context: string;
  pain_points: string[];

  // Slide 3 — Solution
  solution_description: string;
  solution_phases: SolutionPhase[];

  // Slide 4 — Why Loginom
  fit_arguments: string[];

  // Slide 5 — Cases
  relevant_cases: RelevantCase[];

  // Slide 6 — About Loginom
  loginom_about: string;
  loginom_stats: CompanyStat[];

  // Slide 7 — Contacts
  contact_name: string;
  contact_email: string;
}


function blueHeader(// eslint-disable-next-line @typescript-eslint/no-explicit-any
slide: any, title: string) {
  slide.addShape("rect", { x: 0, y: 0, w: "100%", h: 1.1, fill: { color: C.dark } });
  slide.addText(title, {
    x: 0.5, y: 0.15, w: 12.3, h: 0.8,
    fontSize: 22, bold: true, color: C.white,
  });
}

function tag(// eslint-disable-next-line @typescript-eslint/no-explicit-any
slide: any, text: string, x: number, y: number) {
  slide.addShape("rect", { x, y, w: text.length * 0.13 + 0.4, h: 0.38, fill: { color: C.light }, line: { color: C.red, width: 1 } });
  slide.addText(text, { x: x + 0.12, y: y + 0.04, w: text.length * 0.13 + 0.2, h: 0.3, fontSize: 11, color: C.red, bold: true });
}

//Slide 1 — Cover 

export function slide1Cover(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();

  s.addShape("rect", { x: 0, y: 0, w: "100%", h: "100%", fill: { color: C.dark } });
  s.addShape("rect", { x: 0, y: 5.6, w: "100%", h: 1.9, fill: { color: C.darkAlt } });

  s.addText("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", {
    x: 0.7, y: 0.8, w: 12, h: 0.55,
    fontSize: 13, color: "9BA3B0", bold: true, charSpacing: 3,
  });

  s.addText(`для ${d.client_name}`, {
    x: 0.7, y: 1.5, w: 11.5, h: 1.6,
    fontSize: 38, bold: true, color: C.white,
  });

  s.addText(d.client_industry, {
    x: 0.7, y: 3.2, w: 10, h: 0.55,
    fontSize: 18, color: "9BA3B0", italic: true,
  });

  s.addShape("rect", { x: 0.7, y: 3.95, w: 11, h: 0.04, fill: { color: C.darkAlt } });

  s.addText("Loginom Company", {
    x: 0.7, y: 4.15, w: 7, h: 0.5,
    fontSize: 16, color: C.white, bold: true,
  });
  s.addText("loginom.ru", {
    x: 0.7, y: 4.65, w: 7, h: 0.4,
    fontSize: 13, color: "9BA3B0",
  });
  s.addText(d.date, {
    x: 9.5, y: 4.15, w: 3, h: 0.5,
    fontSize: 13, color: "9BA3B0", align: "right",
  });
}

//Slide 2 — Understanding the task 

export function slide2Task(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();
  blueHeader(s, "Понимание задачи");

  s.addText("Запрос клиента", {
    x: 0.4, y: 1.25, w: 5.8, h: 0.4,
    fontSize: 13, bold: true, color: C.red,
  });
  s.addShape("rect", { x: 0.4, y: 1.65, w: 5.8, h: 2.2, fill: { color: C.light }, line: { color: "E2E5EA", width: 1 } });
  s.addText(d.order_description, {
    x: 0.6, y: 1.75, w: 5.4, h: 2,
    fontSize: 12, color: C.dark,
  });

  tag(s, `Бюджет: ${d.budget}`, 0.4, 4);
  tag(s, `Срок: ${d.deadline}`, 0.4 + d.budget.length * 0.13 + 1.1, 4);
  if (d.pilot_format) tag(s, d.pilot_format, 0.4, 4.5);

  s.addText("Контекст и боль", {
    x: 6.8, y: 1.25, w: 6.3, h: 0.4,
    fontSize: 13, bold: true, color: C.red,
  });
  s.addText(d.problem_context, {
    x: 6.8, y: 1.65, w: 6.3, h: 1.5,
    fontSize: 12, color: C.dark,
  });

  s.addText("Что будет, если не решить:", {
    x: 6.8, y: 3.25, w: 6.3, h: 0.35,
    fontSize: 12, bold: true, color: C.dark,
  });
  d.pain_points.slice(0, 3).forEach((point, i) => {
    s.addShape("rect", { x: 6.8, y: 3.65 + i * 0.72, w: 0.35, h: 0.35, fill: { color: C.red } });
    s.addText(point, {
      x: 7.25, y: 3.65 + i * 0.72, w: 5.85, h: 0.55,
      fontSize: 12, color: C.dark,
    });
  });
}

//Slide 3 — Solution

export function slide3Solution(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();
  blueHeader(s, "Предлагаемое решение");

  s.addShape("rect", { x: 0, y: 1.1, w: 4.5, h: 6.4, fill: { color: C.dark } });
  s.addShape("rect", { x: 0.4, y: 1.5, w: 1.8, h: 0.06, fill: { color: C.red } });
  s.addText("Концепция", {
    x: 0.4, y: 1.65, w: 3.7, h: 0.5,
    fontSize: 14, bold: true, color: C.red,
  });
  s.addText(d.solution_description, {
    x: 0.4, y: 2.3, w: 3.7, h: 5,
    fontSize: 11, color: "C0C6D0", lineSpacingMultiple: 1.55,
  });

  const phases = d.solution_phases.slice(0, 4);
  const rightX = 4.85;
  const rightW = 8.1;
  const gap = 0.18;
  const cols = phases.length === 4 ? 2 : 1;
  const rows = Math.ceil(phases.length / cols);
  const cardW = cols === 2 ? (rightW - gap) / 2 : rightW;
  const cardH = (6.1 - gap * (rows - 1)) / rows;

  phases.forEach((phase, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = rightX + col * (cardW + gap);
    const y = 1.22 + row * (cardH + gap);

    s.addShape("rect", { x, y, w: cardW, h: cardH, fill: { color: C.light }, line: { color: "E2E5EA", width: 1 } });
    s.addShape("rect", { x, y, w: cardW, h: 0.07, fill: { color: C.red } });

    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.2, y: y + 0.18, w: 0.65, h: 0.55,
      fontSize: 22, bold: true, color: C.red,
    });
    s.addText(phase.title, {
      x: x + 1, y: y + 0.18, w: cardW - 1.2, h: 0.6,
      fontSize: 12, bold: true, color: C.dark, valign: "middle",
    });
    s.addText(phase.description, {
      x: x + 0.2, y: y + 0.88, w: cardW - 0.4, h: cardH - 1.05,
      fontSize: 11, color: C.dark, lineSpacingMultiple: 1.4,
    });
  });
}

//Slide 4 — Why Loginom 

export function slide4WhyUs(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();
  blueHeader(s, "Почему Loginom");

  s.addShape("rect", { x: 0, y: 1.1, w: "100%", h: 6.4, fill: { color: "F4F5F7" } });

  const args = d.fit_arguments.slice(0, 5);
  const cardH = 1.08;
  const gap = 0.12;

  args.forEach((arg, i) => {
    const y = 1.3 + i * (cardH + gap);

    s.addShape("rect", { x: 0.5, y, w: 12.3, h: cardH, fill: { color: C.white }, line: { color: "E8EAED", width: 1 } });

    s.addShape("rect", { x: 0.5, y, w: 0.07, h: cardH, fill: { color: C.red } });

    s.addShape("rect", { x: 0.77, y: y + 0.19, w: 0.7, h: 0.7, fill: { color: C.dark } });
    s.addText(String(i + 1).padStart(2, "0"), {
      x: 0.77, y: y + 0.19, w: 0.7, h: 0.7,
      fontSize: 14, bold: true, color: C.red, align: "center", valign: "middle",
    });

    s.addText(arg, {
      x: 1.65, y: y + 0.09, w: 11, h: cardH - 0.18,
      fontSize: 13, color: C.dark, valign: "middle",
    });
  });
}

//Slide 5 — Cases 

export function slide5Cases(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();
  blueHeader(s, `Наш опыт в вашей отрасли`);

  const cases = d.relevant_cases.slice(0, 3);
  let w: number;
  if (cases.length === 1) w = 12.5;
  else if (cases.length === 2) w = 6;
  else w = 3.9;

  cases.forEach((c, i) => {
    const x = 0.4 + i * (w + 0.25);

    s.addShape("rect", { x, y: 1.2, w, h: 5.5, fill: { color: C.light }, line: { color: "E2E5EA", width: 1 } });
    s.addShape("rect", { x, y: 1.2, w, h: 0.55, fill: { color: C.dark } });
    s.addText(c.title, {
      x: x + 0.15, y: 1.22, w: w - 0.3, h: 0.5,
      fontSize: 13, bold: true, color: C.white, valign: "middle",
    });
    s.addText(c.description, {
      x: x + 0.15, y: 1.9, w: w - 0.3, h: 2.8,
      fontSize: 12, color: C.dark,
    });
    s.addShape("rect", { x, y: 5.1, w, h: 0.04, fill: { color: C.dark } });
    s.addText(`Результат: ${c.result}`, {
      x: x + 0.15, y: 5.2, w: w - 0.3, h: 1.2,
      fontSize: 12, bold: true, color: C.red,
    });
  });
}

//Slide 6 — About Loginom 

export function slide6About(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();
  blueHeader(s, "О компании Loginom");

  s.addShape("rect", { x: 0, y: 1.1, w: 4.5, h: 6.4, fill: { color: C.dark } });
  s.addShape("rect", { x: 0.4, y: 1.5, w: 1.8, h: 0.06, fill: { color: C.red } });
  s.addText("Ключевые факты", {
    x: 0.4, y: 1.65, w: 3.7, h: 0.45,
    fontSize: 13, bold: true, color: C.red,
  });

  const stats = d.loginom_stats.slice(0, 4);
  stats.forEach((stat, i) => {
    const y = 2.3 + i * 1.25;
    s.addText(stat.value, {
      x: 0.4, y, w: 3.7, h: 0.72,
      fontSize: 30, bold: true, color: C.white,
    });
    s.addText(stat.label, {
      x: 0.4, y: y + 0.72, w: 3.7, h: 0.35,
      fontSize: 10.5, color: C.gray,
    });
    if (i < stats.length - 1) {
      s.addShape("rect", { x: 0.4, y: y + 1.15, w: 3.2, h: 0.03, fill: { color: "2A3144" } });
    }
  });

  s.addShape("rect", { x: 4.75, y: 1.2, w: 8.2, h: 6.1, fill: { color: C.light }, line: { color: "E2E5EA", width: 1 } });
  s.addText(d.loginom_about, {
    x: 5.05, y: 1.4, w: 7.6, h: 5.7,
    fontSize: 12, color: C.dark, lineSpacingMultiple: 1.6,
  });
}

//Slide 7 — Next steps 

export function slide7Contacts(// eslint-disable-next-line @typescript-eslint/no-explicit-any
pptx: any, d: PresentationInput) {
  const s = pptx.addSlide();
  blueHeader(s, "Следующие шаги и контакты");

  const steps = [
    "Ответный звонок для уточнения деталей задачи",
    "Подготовка технического предложения (2 недели)",
    "Согласование условий и подписание договора",
    "Старт проекта",
  ];

  steps.forEach((step, i) => {
    const y = 1.3 + i * 1.1;
    s.addShape("rect", { x: 0.4, y, w: 0.75, h: 0.75, fill: { color: C.dark } });
    s.addText(`${i + 1}`, {
      x: 0.4, y, w: 0.75, h: 0.75,
      fontSize: 20, bold: true, color: C.white, align: "center", valign: "middle",
    });
    s.addText(step, {
      x: 1.35, y: y + 0.05, w: 7.5, h: 0.65,
      fontSize: 14, color: C.dark, valign: "middle",
    });
  });

  s.addShape("rect", { x: 9.3, y: 1.3, w: 3.7, h: 4.3, fill: { color: C.dark } });
  s.addText("Loginom Company", {
    x: 9.5, y: 1.55, w: 3.3, h: 0.6,
    fontSize: 16, bold: true, color: C.white,
  });
  s.addText("sales@loginom.ru\n+7 (383) 206-09-40\nloginom.ru", {
    x: 9.5, y: 2.3, w: 3.3, h: 1.6,
    fontSize: 13, color: "9BA3B0",
  });
  s.addShape("rect", { x: 9.5, y: 4.1, w: 3.1, h: 0.03, fill: { color: C.darkAlt } });
  s.addText(`Подготовлено для:\n${d.contact_name}\n${d.contact_email}`, {
    x: 9.5, y: 4.25, w: 3.3, h: 1.1,
    fontSize: 11, color: "9BA3B0",
  });
}
