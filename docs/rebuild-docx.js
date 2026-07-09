const { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, Table, TableRow, TableCell, AlignmentType, PageBreak, BorderStyle, TableLayoutType, WidthType, ShadingType, LevelFormat, NumberFormat, convertInchesToTwip } = require("docx");
const fs = require("fs");
const path = require("path");

const docsDir = "C:/Users/Administrator/Desktop/SkillLink/docs";
const outputPath = path.join(docsDir, "SkillLink_User_Research_Report_V2.docx");
const figmaUrl = "https://www.figma.com/design/AVLBzCgnVANAHBGL9kL0VR/Untitled?node-id=0-1&t=WV8dvdlUtMqJ865a-1";

function readImg(name) {
  const p = path.join(docsDir, name);
  return fs.existsSync(p) ? fs.readFileSync(p) : null;
}

const workflowImg = readImg("Workflow.png");
const empathyImg = readImg("Empathy map.png");
const ssImg = readImg("SS.png");

// Reusable table cell
function cell(text, opts = {}) {
  const children = [];
  if (text) {
    children.push(new Paragraph({
      children: [new TextRun({ text, bold: !!opts.bold, size: opts.size || 22, color: opts.color || "000000" })],
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 40, after: 40 }
    }));
  }
  return new TableCell({
    children: children.length ? children : [new Paragraph({ children: [new TextRun("")] })],
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    shading: opts.shading ? { type: ShadingType.CLEAR, fill: opts.shading } : undefined,
    verticalAlign: "center"
  });
}

function headerCell(text, width) {
  return cell(text, { bold: true, width, shading: "3880FF", color: "ffffff", size: 22 });
}

function bodyCell(text, width) {
  return cell(text, { width, size: 22 });
}

function emptyCell(width) {
  return cell(null, { width });
}

// Build document
const children = [];

// ── TITLE PAGE ──
children.push(new Paragraph({ spacing: { before: 3000 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({ text: "SkillLink", size: 48, bold: true, color: "3880FF" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: "User Research Report", size: 36, color: "212529" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({ text: "A Mobile Freelance Job Matching Platform for Pakistan", size: 24, color: "6C757D" })]
}));
children.push(new Paragraph({ spacing: { before: 100 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Product Design Team  |  July 2026  |  Version 2.0", size: 22, color: "6C757D" })]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ── TABLE OF CONTENTS ──
children.push(new Paragraph({
  spacing: { before: 300, after: 200 },
  children: [new TextRun({ text: "Table of Contents", size: 28, bold: true, color: "3880FF" })]
}));
const toc = [
  "1. Executive Summary",
  "2. User Persona",
  "    2.1 Overview",
  "    2.2 Goals",
  "    2.3 Pain Points",
  "3. Empathy Map",
  "    3.1 Says",
  "    3.2 Thinks",
  "    3.3 Does",
  "    3.4 Feels",
  "4. Problem Statement",
  "    4.1 Context",
  "5. How Might We (HMW) Questions",
  "    5.1 HMW 1 - Trust & Verification",
  "    5.2 HMW 2 - Hyper-Local Job Matching",
  "6. User Workflow",
  "7. Design System Addendum",
  "    7.1 Colour Palette",
  "    7.2 Typography Scale",
  "Appendix A - Methodology",
  "8. Interactive Prototype"
];
for (const t of toc) {
  children.push(new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text: t, size: 22, color: t.startsWith("    ") ? "495057" : "212529" })]
  }));
}
children.push(new Paragraph({ children: [new PageBreak()] }));

// ── 1. EXECUTIVE SUMMARY ──
children.push(secH1("1. Executive Summary"));
children.push(body("Pakistan's freelance economy is expanding rapidly, yet freelancers lack a dedicated local platform that connects them with organizations offering project-based work. Global platforms such as Upwork and Fiverr dominate the market but impose high commission fees, lack regional context, and offer inadequate trust and dispute resolution mechanisms tailored to the Pakistani ecosystem. This report documents the core user persona, empathy map, problem statement, and strategic How Might We questions that will guide the design and development of SkillLink."));

// ── 2. USER PERSONA ──
children.push(secH1("2. User Persona"));
children.push(secH2("2.1 Overview"));
children.push(makeTable(
  ["Attribute", "Detail"],
  [
    ["Name", "Ahmad Khan"],
    ["Age", "25"],
    ["Occupation", "Freelance Graphic Designer"],
    ["Location", "Lahore, Pakistan"],
    ["Education", "Bachelors in Visual Communication Design"],
    ["Tech Comfort Level", "High - smartphone-dependent, proficient with mobile apps, social media, and digital payment platforms"],
    ["Platforms Used", "Upwork, Fiverr, Facebook Groups, Instagram"],
    ["Income Range", "PKR 30,000 - 60,000/month (variable)"],
    ["Payment Methods", "JazzCash, Easypaisa, Bank Transfer"]
  ],
  [30, 70]
));

children.push(secH2("2.2 Goals"));
for (const g of [
  "Find 3-4 reliable local clients for consistent monthly income",
  "Reduce dependence on high-commission international platforms",
  "Build a long-term reputation within the Pakistani freelance ecosystem"
]) {
  children.push(bullet(g));
}

children.push(secH2("2.3 Pain Points"));
children.push(makeTable(
  ["Pain Point", "Impact"],
  [
    ["Inconsistent project flow", "Income instability, stress during dry periods"],
    ["20% commission on global platforms", "Significant reduction in take-home earnings"],
    ["No local discovery channel", "Missed opportunities with nearby organizations"],
    ["Client ghosting post-delivery", "Unpaid work, no legal or platform-level recourse"],
    ["Cross-border timezone friction", "Delayed communication and feedback cycles"]
  ],
  [40, 60]
));

// ── 3. EMPATHY MAP ──
children.push(secH1("3. Empathy Map"));
children.push(body("The empathy map captures Ahmad's current experience across four dimensions: what he says, thinks, does, and feels in relation to finding freelance work."));

if (empathyImg) {
  children.push(imgPara(empathyImg, 500, 300));
}

children.push(secH2("3.1 Says"));
for (const { quote, context } of [
  { quote: "I wish there was a local platform where I can find genuine clients.", context: "Ahmad openly expresses frustration with the lack of a dedicated Pakistani freelance marketplace. He frequently discusses the gap with fellow freelancers in online communities." },
  { quote: "Payment delays kill my motivation.", context: "Delayed or missing payments are a recurring topic in his conversations, affecting his enthusiasm for taking on new projects." },
  { quote: "I keep refreshing the same platforms but see the same old projects.", context: "He notices stagnation in available opportunities, particularly on global platforms where competition is high and project rotation is slow." }
]) {
  children.push(quotePara(quote, context));
}

children.push(secH2("3.2 Thinks"));
for (const { quote, context } of [
  { quote: "Upwork takes too much commission, and most clients are international with timezone issues.", context: "Ahmad internally calculates the cost of using global platforms and finds the value proposition weakening over time." },
  { quote: "If I could just find 3-4 reliable local clients, I would be set.", context: "He believes that a small, consistent client base would solve most of his income stability problems." },
  { quote: "There must be other freelancers like me struggling to find work.", context: "He senses a shared pain point across the community and wonders why no solution has emerged yet." }
]) {
  children.push(quotePara(quote, context));
}

children.push(secH2("3.3 Does"));
for (const d of [
  "Checks Upwork and Fiverr 3-4 times daily for new project listings",
  "Actively browses Facebook freelance groups for leads",
  "Maintains an Instagram portfolio to attract direct client inquiries",
  "Uses JazzCash and Easypaisa for payment transactions",
  "Networks with other freelancers in local co-working spaces and WhatsApp groups"
]) {
  children.push(bullet(d));
}

children.push(secH2("3.4 Feels"));
children.push(makeTable(
  ["Emotion", "Context"],
  [
    ["Anxious", "About next month's income and irregular cash flow"],
    ["Frustrated", "With high platform fees and lack of local options"],
    ["Hopeful", "That a better, localized solution may emerge"],
    ["Undervalued", "When local clients bargain aggressively on pricing"],
    ["Isolated", "Competing in a global marketplace without regional support"]
  ],
  [30, 70]
));

// ── 4. PROBLEM STATEMENT ──
children.push(secH1("4. Problem Statement"));
children.push(body("[A freelance professional in Pakistan] needs a way [to discover and apply for local project opportunities] because [existing global platforms lack regional context, trust mechanisms, and localized payment solutions, making it hard to find consistent, legitimate work]."));
children.push(secH2("4.1 Context"));
children.push(body("The problem sits at the intersection of three gaps in the current market:"));
children.push(bullet("Discovery Gap", "No centralized platform curates local projects specifically for Pakistani freelancers."));
children.push(bullet("Trust Gap", "No local identity verification or reputation system tailored to the Pakistani market exists, leading to ghosting and fraud."));
children.push(bullet("Payment Gap", "Global platforms do not integrate with local payment rails (JazzCash, Easypaisa, bank transfers), causing withdrawal friction and delays."));

// ── 5. HMW QUESTIONS ──
children.push(secH1("5. How Might We (HMW) Questions"));
children.push(body("The following HMW questions translate the problem statement into actionable design challenges:"));
children.push(secH2("5.1 HMW 1 - Trust & Verification"));
children.push(body("How might we create a trust and verification layer that authenticates both freelancers and clients to reduce ghosting and fraud in the Pakistani freelance market?"));
children.push(body("Design Focus Areas:", true));
for (const d of [
  "Identity verification via national ID (CNIC) and business registration",
  "Client rating system with verified project completion data",
  "Escrow-based payment protection integrated with local payment providers",
  "Dispute resolution workflow managed within the platform"
]) {
  children.push(bullet(d));
}

children.push(secH2("5.2 HMW 2 - Hyper-Local Job Matching"));
children.push(body("How might we design a hyper-local job matching experience that surfaces relevant opportunities based on skills, location, and past work history while functioning smoothly on low-bandwidth mobile connections?"));
children.push(body("Design Focus Areas:", true));
for (const d of [
  "Algorithmic matching using skill tags, location proximity, and project history",
  "Offline-first architecture for areas with unstable internet connectivity",
  "Location-aware filters to prioritize nearby organizations",
  "Lightweight app with minimal data consumption for core workflows"
]) {
  children.push(bullet(d));
}

// ── 6. USER WORKFLOW ──
children.push(secH1("6. User Workflow"));
children.push(body("The following diagram illustrates the freelancer's journey from browsing projects to submitting an application, including the key decision point where requirements are evaluated."));
if (workflowImg) {
  children.push(imgPara(workflowImg, 550, 400));
}

// ── 7. DESIGN SYSTEM ADDENDUM ──
children.push(secH1("7. Design System Addendum"));
children.push(body("To ensure a cohesive user experience across the SkillLink mobile platform, the following design system specifications have been established based on the low-fidelity wireframe prototypes. These choices intentionally address the regional constraints, trust gaps, and user behavior outlined in this research report."));

children.push(secH2("7.1 Colour Palette"));
children.push(makeTable(
  ["Token", "Hex Code", "Justification"],
  [
    ["Brand Primary", "#3880FF", "A vivid, tech-forward blue that establishes platform authority and highlights active navigation states or active interactive filters like hyper-local matching filters."],
    ["Success & Trust", "#2B9348", "A secure green utilized specifically for positive feedback loops, verified client milestones, budget highlights, and escrow safety indicators to alleviate user anxiety regarding payment delivery."],
    ["Neutral Dark (Text)", "#212529", "High-contrast charcoal black ensuring deep readability for critical hierarchy text, such as project titles and main confirmation headers on mobile layouts."],
    ["Neutral Muted", "#6C757D", "Slate gray reserved for supporting metadata, passive labels, layout borders, and inactive navigation items to reduce cognitive load."],
    ["Surface Light", "#F8F9FA", "A crisp, low-intensity off-white that acts as the canvas for card elements, giving structure to distinct content segments without straining the user's eye on mobile screens."]
  ],
  [20, 15, 65]
));

children.push(secH2("7.2 Typography Scale"));
children.push(makeTable(
  ["Style", "Font Family", "Size", "Weight", "Sample Text"],
  [
    ["H1", "'Segoe UI', Arial, sans-serif", "22px", "Bold", "Application Submitted!"],
    ["Body", "'Segoe UI', Arial, sans-serif", "14px", "Regular", "Search local projects (e.g., Graphic Design)..."],
    ["Caption", "'Segoe UI', Arial, sans-serif", "11px", "Semi-Bold", "Lahore"]
  ],
  [15, 30, 12, 15, 28]
));

// ── APPENDIX A ──
children.push(secH1("Appendix A - Methodology"));
children.push(body("This report was developed through:"));
for (const d of [
  "Secondary research on Pakistan's freelance economy statistics",
  "Competitive analysis of Upwork, Fiverr, and regional platforms",
  "Informal interviews with 5 freelance professionals in Lahore and Karachi",
  "Synthesis using the Persona + Empathy Map framework from Design Thinking"
]) {
  children.push(bullet(d));
}

// ── 8. INTERACTIVE PROTOTYPE ──
children.push(secH1("8. Interactive Prototype"));
children.push(body("The following Figma prototype demonstrates the interactive clickable prototype flow:"));
children.push(new Paragraph({
  spacing: { before: 100, after: 150 },
  children: [
    new TextRun({ text: "Figma Interactive Prototype: ", size: 22, bold: true, color: "212529" }),
    new TextRun({ text: figmaUrl, size: 22, color: "3880FF", underline: {} })
  ]
}));

children.push(body("Below is a screenshot of the prototype screens connected in the flow:"));
if (ssImg) {
  children.push(imgPara(ssImg, 550, 350));
}

children.push(new Paragraph({
  spacing: { before: 300, after: 200 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "--- End of Report ---", size: 22, bold: true, color: "6C757D" })]
}));

// ── BUILD ──
const doc = new Document({
  sections: [{ children }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outputPath, buf);
  console.log("Done! Rebuilt docx: " + (buf.length / 1024).toFixed(1) + " KB");
}).catch(console.error);


// ── HELPERS ──

function secH1(text) {
  return new Paragraph({
    spacing: { before: 300, after: 150 },
    children: [new TextRun({ text, size: 32, bold: true, color: "212529" })]
  });
}

function secH2(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, size: 26, bold: true, color: "212529" })]
  });
}

function body(text, bold = false) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, bold, color: "212529" })]
  });
}

function bullet(text, sub = null) {
  const runs = [new TextRun({ text: "- " + text, size: 22, color: "212529" })];
  if (sub) {
    runs.push(new TextRun({ text: "\n  " + sub, size: 20, color: "6C757D" }));
  }
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: convertInchesToTwip(0.3) },
    children: runs
  });
}

function quotePara(quote, context) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: convertInchesToTwip(0.3) },
    children: [
      new TextRun({ text: "\"" + quote + "\"", size: 22, italics: true, color: "495057" }),
      new TextRun({ text: "\n" + context, size: 21, color: "212529" })
    ]
  });
}

function imgPara(buffer, width, height) {
  if (!buffer) return new Paragraph({ children: [new TextRun({ text: "[Image not found]", size: 20, italics: true, color: "999999" })] });
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 100 },
    children: [
      new ImageRun({
        data: buffer,
        transformation: { width, height }
      })
    ]
  });
}

function makeTable(headers, rows, widths) {
  const hCells = headers.map((h, i) => headerCell(h, widths[i]));
  const headerRow = new TableRow({ children: hCells });

  const dataRows = rows.map(r => {
    const row = r.map((c, i) => bodyCell(c, widths[i]));
    return new TableRow({ children: row });
  });

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: "DEE2E6" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "DEE2E6" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "DEE2E6" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "DEE2E6" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "DEE2E6" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "DEE2E6" }
    }
  });
}
