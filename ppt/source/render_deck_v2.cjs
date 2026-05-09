const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const ASSET_DIR = path.join(ROOT, "assets");
const PROTO_DIR = path.join(ASSET_DIR, "prototype");
const WORKSHOP_DIR = path.join(ASSET_DIR, "workshop");
const USER_TESTING_DIR = path.join(ASSET_DIR, "user-testing");
const SLIDES_DIR = path.join(ROOT, "slides-v2");
const QA_DIR = path.join(ROOT, "qa");
const OUTPUT_DIR = path.join(ROOT, "output");
for (const dir of [ASSET_DIR, SLIDES_DIR, QA_DIR, OUTPUT_DIR]) fs.mkdirSync(dir, { recursive: true });

const W = 1920;
const H = 1080;
const OUT = path.join(OUTPUT_DIR, "the-earth-written-by-us-radical-earth.html");
const PPTX_OUT = path.join(OUTPUT_DIR, "the-earth-written-by-us-earth-story-v2.pptx");

const C = {
  ink: "#10212b",
  slate: "#53636a",
  paper: "#f6f0e6",
  white: "#fffaf1",
  green: "#4f8f68",
  mint: "#d8eadc",
  amber: "#e6b85c",
  coral: "#e86f5c",
  blue: "#577d91",
  red: "#a44942",
  line: "#d2c3ad",
};

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function imgData(file) {
  const ext = path.extname(file).slice(1) || "png";
  return `data:image/${ext};base64,${fs.readFileSync(file).toString("base64")}`;
}

function wrap(text, maxChars) {
  const words = String(text).split(/\s+/);
  const lines = [];
  let line = "";
  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (next.length > maxChars && line) {
      lines.push(line);
      line = word;
    } else {
      line = next;
    }
  }
  if (line) lines.push(line);
  return lines;
}

function t(text, x, y, w, opts = {}) {
  const size = opts.size || 32;
  const lines = Array.isArray(text) ? text : wrap(text, opts.maxChars || Math.max(18, Math.floor(w / (size * 0.52))));
  const lh = opts.lineHeight || Math.round(size * 1.24);
  const fill = opts.fill || C.ink;
  const weight = opts.weight || 450;
  const anchor = opts.anchor || "start";
  return `<text x="${x}" y="${y}" fill="${fill}" font-family="Inter, Avenir Next, Helvetica Neue, Arial, sans-serif" font-size="${size}" font-weight="${weight}" text-anchor="${anchor}">${lines
    .map((line, i) => `<tspan x="${x}" dy="${i === 0 ? 0 : lh}">${esc(line)}</tspan>`)
    .join("")}</text>`;
}

function bg() {
  return `
  <defs>
    <radialGradient id="paperGlow" cx="22%" cy="10%" r="95%">
      <stop offset="0%" stop-color="#fffaf1"/>
      <stop offset="62%" stop-color="#f4ebdc"/>
      <stop offset="100%" stop-color="#e8dccb"/>
    </radialGradient>
    <pattern id="grain" width="96" height="96" patternUnits="userSpaceOnUse">
      <circle cx="12" cy="18" r="1.4" fill="#d8c9b4" opacity=".26"/>
      <circle cx="44" cy="52" r="1.1" fill="#cdbb9f" opacity=".18"/>
      <circle cx="82" cy="31" r="1.2" fill="#e3d4bd" opacity=".3"/>
      <path d="M0 78 Q24 70 48 80 T96 79" stroke="#dacbb5" stroke-width="1" opacity=".12" fill="none"/>
    </pattern>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="15" stdDeviation="18" flood-color="#10212b" flood-opacity=".13"/>
    </filter>
    <marker id="arrow" markerWidth="13" markerHeight="13" refX="10" refY="6" orient="auto">
      <path d="M1,1 L11,6 L1,11 Z" fill="${C.ink}"/>
    </marker>
  </defs>
  <rect width="${W}" height="${H}" fill="url(#paperGlow)"/>
  <rect width="${W}" height="${H}" fill="url(#grain)" opacity=".72"/>
  <path d="M-30 1010 C300 920 520 1060 850 980 C1210 890 1450 990 1950 910 L1950 1100 L-30 1100 Z" fill="#dfe9df" opacity=".44"/>
  <path d="M1500 -130 C1710 25 1775 220 1950 310" stroke="#d7c8ae" stroke-width="2" fill="none" opacity=".42"/>`;
}

function chrome(n) {
  return `<text x="98" y="70" font-family="Inter, Arial, sans-serif" font-size="24" fill="${C.slate}" letter-spacing="1.8">RADICAL EARTH PUBLISHING FORM</text><text x="1822" y="70" font-family="Inter, Arial, sans-serif" font-size="24" fill="${C.slate}" text-anchor="end">${String(n).padStart(2, "0")} / 16</text>`;
}

function title(n, headline, sub = "") {
  const lines = wrap(headline, 31);
  const subY = 165 + (lines.length - 1) * 76 + 98;
  return `${chrome(n)}${t(lines, 96, 165, 1120, { size: 68, weight: 790, lineHeight: 76 })}${sub ? t(sub, 100, subY, 1180, { size: 28, fill: C.slate, maxChars: 76, lineHeight: 36 }) : ""}`;
}

function card(x, y, w, h, body, fill = C.white) {
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="18" fill="${fill}" stroke="${C.line}" stroke-width="2" filter="url(#shadow)"/>${body}`;
}

function pill(x, y, label, fill = C.mint, stroke = C.green, w = null) {
  const width = w || Math.max(120, label.length * 15 + 48);
  return `<rect x="${x}" y="${y}" width="${width}" height="54" rx="27" fill="${fill}" stroke="${stroke}" stroke-width="2"/><text x="${x + width / 2}" y="${y + 35}" font-family="Inter, Arial" font-size="23" font-weight="700" fill="${C.ink}" text-anchor="middle">${esc(label)}</text>`;
}

function arrow(x1, y1, x2, y2, color = C.ink) {
  return `<path d="M${x1} ${y1} C${(x1 + x2) / 2} ${y1}, ${(x1 + x2) / 2} ${y2}, ${x2} ${y2}" stroke="${color}" stroke-width="5" fill="none" marker-end="url(#arrow)"/>`;
}

function globe(cx, cy, r, state = "repair") {
  const base = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#cfe2d3" stroke="${C.ink}" stroke-width="6"/><path d="M${cx - r * .52} ${cy - r * .06} C${cx - r * .26} ${cy - r * .48} ${cx + r * .14} ${cy - r * .38} ${cx + r * .22} ${cy - r * .1} C${cx + r * .38} ${cy + r * .22} ${cx - r * .07} ${cy + r * .2} ${cx - r * .03} ${cy + r * .5}" fill="${C.green}" opacity=".9"/><path d="M${cx + r * .2} ${cy - r * .62} C${cx + r * .62} ${cy - r * .35} ${cx + r * .44} ${cy + r * .1} ${cx + r * .72} ${cy + r * .28}" stroke="${C.blue}" stroke-width="${r * .13}" stroke-linecap="round" fill="none" opacity=".42"/>`;
  const extras = {
    repair: `<path d="M${cx - r * .62} ${cy + r * .12} C${cx - r * .2} ${cy - r * .3} ${cx + r * .26} ${cy + r * .22} ${cx + r * .64} ${cy - r * .16}" stroke="${C.amber}" stroke-width="${r * .055}" stroke-linecap="round" fill="none"/><circle cx="${cx + r * .64}" cy="${cy - r * .16}" r="${r * .08}" fill="${C.amber}"/>`,
    crack: `<path d="M${cx - r * .12} ${cy - r * .56} L${cx + r * .08} ${cy - r * .15} L${cx - r * .06} ${cy + r * .16} L${cx + r * .2} ${cy + r * .52}" stroke="${C.red}" stroke-width="${r * .045}" fill="none"/>`,
    bloom: `<circle cx="${cx - r * .36}" cy="${cy - r * .36}" r="${r * .08}" fill="${C.green}"/><circle cx="${cx + r * .38}" cy="${cy + r * .24}" r="${r * .09}" fill="${C.green}"/><circle cx="${cx + r * .12}" cy="${cy - r * .55}" r="${r * .07}" fill="${C.green}"/>`,
  };
  return `<g filter="url(#shadow)">${base}${extras[state] || extras.repair}</g>`;
}

const titleHero = path.join(ASSET_DIR, "title-earth-story.png");
const earthStates = path.join(ASSET_DIR, "earth-states-generated.png");

const slides = [
  () => {
    const hero = imgData(titleHero);
    return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}<image href="${hero}" x="0" y="0" width="${W}" height="${H}" preserveAspectRatio="xMidYMid slice"/><rect width="${W}" height="${H}" fill="#10212b" opacity=".48"/><text x="98" y="104" font-family="Inter, Arial" font-size="27" fill="#f6f0e6" letter-spacing="2">THE BRIEF</text>${t(["Publishing Forms:", "Radical Earth"], 96, 255, 900, { size: 92, fill: C.white, weight: 790, lineHeight: 100 })}${t("The Earth, Written by Us", 100, 512, 760, { size: 42, fill: "#efe4ce", weight: 760, maxChars: 36, lineHeight: 50 })}${t("An Earth Story You Can Change", 100, 578, 720, { size: 32, fill: "#efe4ce", maxChars: 36, lineHeight: 42 })}<rect x="100" y="820" width="620" height="72" rx="36" fill="#f6f0e6" opacity=".92"/><text x="410" y="866" font-family="Inter, Arial" font-size="27" font-weight="700" fill="${C.ink}" text-anchor="middle">Prototype + publishing strategy</text></svg>`;
  },
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(2, "Brief Interpretation", "Publishing beyond the static printed page.")}
    ${card(115, 375, 500, 385, t("The brief asks for a resolved publishing platform or event, not only a book.", 160, 455, 390, { size: 35, weight: 700, maxChars: 24, lineHeight: 45 }), "#fff8ea")}
    ${card(710, 340, 520, 465, `${t("Publication becomes:", 760, 420, 360, { size: 34, weight: 780 })}${["interface", "system", "event", "hybrid format"].map((d, i) => pill(760, 485 + i * 70, d, i % 2 ? "#f8dfd9" : C.mint, i % 2 ? C.coral : C.green, 310)).join("")}`, "#fbf5ea")}
    ${card(1320, 375, 490, 385, `${t("My response", 1370, 455, 350, { size: 34, weight: 780 })}${t("A web publication where each reader word triggers story text, shifts Earth variables, and outputs an ending card.", 1370, 525, 355, { size: 30, maxChars: 26, lineHeight: 39 })}`, "#f2f8f0")}
    ${arrow(615, 570, 710, 555)}${arrow(1230, 555, 1320, 570)}</svg>`,
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(3, "Given Content Analysis", "This page makes the source text and editing strategy explicit.")}
    ${card(100, 350, 535, 520, `${t("Original content", 150, 420, 380, { size: 34, weight: 800 })}${t("[TO FILL] Title / author / source of the given text.", 150, 500, 390, { size: 29, fill: C.red, maxChars: 31, lineHeight: 37 })}${t("[TO FILL] Core argument, tone, key phrases and recurring motifs.", 150, 625, 390, { size: 29, fill: C.red, maxChars: 31, lineHeight: 37 })}`, "#fff8ea")}
    ${card(690, 350, 535, 520, `${t("My interpretation", 740, 420, 390, { size: 34, weight: 800 })}${t("The text becomes a branching ecological story. Key ideas become prompts, keywords, Earth states and endings.", 740, 500, 390, { size: 30, maxChars: 31, lineHeight: 39 })}${pill(740, 715, "text → keyword", C.mint, C.green, 250)}${pill(1010, 715, "choice → consequence", "#fff1d6", C.amber, 310)}`, "#f2f8f0")}
    ${card(1280, 350, 535, 520, `${t("Editing strategy", 1330, 420, 390, { size: 34, weight: 800 })}${t("[TO FILL] Confirm whether the final form uses the full text, selected fragments, rewritten chapters, or an intentional edited version.", 1330, 500, 390, { size: 29, fill: C.red, maxChars: 31, lineHeight: 37 })}`, "#fff4ef")}</svg>`,
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(4, "Core Concept: reading changes the Earth", "The reader does not just consume the story; they alter the system.")}
    ${card(105, 350, 760, 560, `${t("Living Earth", 150, 420, 300, { size: 36, weight: 800 })}${globe(485, 650, 170, "repair")}${t("Health, biodiversity, care, justice, waste and hope are stored as hidden variables.", 185, 855, 560, { size: 28, fill: C.slate, maxChars: 52, lineHeight: 35 })}`, "#f8fbf4")}
    ${card(1020, 350, 790, 560, `<rect x="1095" y="430" width="640" height="320" rx="12" fill="#fffdf6" stroke="${C.line}" stroke-width="3"/><line x1="1415" y1="430" x2="1415" y2="750" stroke="${C.line}" stroke-width="3"/>${t("Chapter 03", 1145, 500, 230, { size: 28, fill: C.coral, weight: 800 })}${t("A word enters the page. The story answers with consequence.", 1145, 560, 230, { size: 29, maxChars: 17, lineHeight: 38 })}${t("keyword", 1470, 520, 170, { size: 25, fill: C.slate, weight: 700 })}<rect x="1470" y="545" width="210" height="58" rx="29" fill="${C.mint}" stroke="${C.green}" stroke-width="2"/><text x="1575" y="583" text-anchor="middle" font-family="Inter, Arial" font-size="26" font-weight="700" fill="${C.ink}">repair</text>${t("Storybook Dialogue", 1095, 830, 480, { size: 36, weight: 800 })}`, "#fff7e8")}
    <path d="M880 635 C950 572 965 572 1010 635" stroke="${C.amber}" stroke-width="9" fill="none" marker-end="url(#arrow)"/></svg>`,
  () => {
    const floor = imgData(path.join(WORKSHOP_DIR, "workshop-floor-cards.jpg"));
    const group = imgData(path.join(WORKSHOP_DIR, "group-systems-game.jpg"));
    const tipping = imgData(path.join(WORKSHOP_DIR, "positive-tipping-report.jpg"));
    const feedback = imgData(path.join(WORKSHOP_DIR, "responsibility-freedom-loop.jpg"));
    return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(5, "World Earth Day research evidence", "Power of Positive Tipping Points / Systems Games Workshop, 22 April 2026.")}
    <image href="${floor}" x="95" y="348" width="520" height="358" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    <image href="${group}" x="95" y="730" width="520" height="235" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    ${card(665, 350, 500, 615, `${t("Experiment Log 04", 715, 420, 340, { size: 32, weight: 800 })}${t("Power of Positive Tipping Points", 715, 475, 350, { size: 29, weight: 700, maxChars: 30, lineHeight: 37 })}${t("Format: Earth Day workshop / field participation", 715, 575, 350, { size: 27, fill: C.slate, maxChars: 31, lineHeight: 35 })}${t("Focus: systems thinking, tipping points, adaptation and feedback.", 715, 675, 350, { size: 27, fill: C.slate, maxChars: 31, lineHeight: 35 })}${t("Insight: the room became the interface. Rules, timing and misunderstanding made system relations visible.", 715, 790, 350, { size: 28, fill: C.ink, maxChars: 31, lineHeight: 36 })}`, "#fffaf1")}
    <image href="${tipping}" x="1215" y="350" width="290" height="282" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    <image href="${feedback}" x="1528" y="350" width="290" height="282" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    ${card(1215, 665, 603, 300, `${t("Translation into the web publication", 1260, 730, 440, { size: 32, weight: 800 })}${pill(1260, 790, "physical system game", "#fff1d6", C.amber, 310)}${pill(1585, 790, "keyword loop", C.mint, C.green, 220)}${t("Small repeated actions become interaction logic: word → feedback → changed Earth state → next choice.", 1260, 880, 470, { size: 28, fill: C.slate, maxChars: 45, lineHeight: 36 })}`, "#f2f8f0")}
    </svg>`;
  },
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(6, "Audience and encounter context", "A narrower audience makes the publishing format more convincing.")}
    ${card(125, 365, 760, 420, `${t("Primary audience", 180, 435, 420, { size: 36, weight: 800 })}${t("WSA design students and World Earth Day workshop participants.", 180, 515, 570, { size: 34, maxChars: 38, lineHeight: 44 })}${t("They already share the brief context and can test whether the story system makes climate action feel participatory.", 180, 650, 570, { size: 28, fill: C.slate, maxChars: 47, lineHeight: 36 })}`, "#f2f8f0")}
    ${card(1035, 365, 760, 420, `${t("Secondary audience", 1090, 435, 420, { size: 36, weight: 800 })}${t("Exhibition visitors and online readers interested in climate storytelling.", 1090, 515, 570, { size: 34, maxChars: 37, lineHeight: 44 })}${t("Encounter: type one word, receive a story response, watch Earth change, save or take away an ending card.", 1090, 650, 570, { size: 28, fill: C.slate, maxChars: 48, lineHeight: 36 })}`, "#fff8ea")}
    ${pill(470, 850, "studio exhibition", "#e7eef1", C.blue, 310)}${pill(805, 850, "Earth Day event", C.mint, C.green, 300)}${pill(1130, 850, "online publication", "#fff1d6", C.amber, 330)}</svg>`,
  () => {
    const shot = imgData(path.join(PROTO_DIR, "03-repair.png"));
    return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(7, "Example reading path", "A complete reader action shows the website mechanics and text quality.")}
    <image href="${shot}" x="96" y="360" width="800" height="535" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    ${card(980, 345, 820, 560, `${[
      ["1", "Reader types: “repair”"],
      ["2", "Story response: “A crack softens. A tool is shared...”"],
      ["3", "Variables change: +Health, +Care, -Waste, +Hope"],
      ["4", "Earth state becomes: Repairing Earth"],
      ["5", "Next choice opens: listen / plant / share"],
      ["6", "Ending path can become: Regenerative Earth"],
    ].map((d, i) => `<circle cx="1038" cy="${420 + i * 75}" r="22" fill="${i % 2 ? C.amber : C.green}"/><text x="1038" y="${429 + i * 75}" font-family="Inter, Arial" font-size="22" font-weight="800" text-anchor="middle" fill="${C.white}">${d[0]}</text><text x="1085" y="${430 + i * 75}" font-family="Inter, Arial" font-size="30" font-weight="650" fill="${C.ink}">${esc(d[1])}</text>`).join("")}`, "#fffaf1")}</svg>`;
  },
  () => {
    const start = imgData(path.join(PROTO_DIR, "02-story-start.png"));
    const ending = imgData(path.join(PROTO_DIR, "07-organise.png"));
    return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(8, "Interface design", "Testing showed some labels were unclear, so the next iteration uses state-first language.")}
    <image href="${start}" x="95" y="350" width="800" height="535" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    <image href="${ending}" x="1025" y="350" width="800" height="535" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    ${pill(190, 900, "Living Earth + variables", C.mint, C.green, 360)}${pill(1125, 900, "Keyword panel + ending result", "#fff1d6", C.amber, 440)}</svg>`;
  },
  () => {
    const img = imgData(earthStates);
    return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(9, "Earth states", "Six public-facing states keep visual feedback consistent across the prototype.")}
      <image href="${img}" x="105" y="350" width="1710" height="550" preserveAspectRatio="xMidYMid meet" filter="url(#shadow)"/>
      ${["Fragile", "Warming", "Repairing", "Blooming", "Collective", "Collapsing"].map((l, i) => `<text x="${235 + i * 291}" y="938" text-anchor="middle" font-family="Inter, Arial" font-size="27" font-weight="760" fill="${i === 5 ? C.red : C.ink}">${l} Earth</text>`).join("")}
    </svg>`;
  },
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(10, "Keyword and hidden variable system", "The presentation highlights 13 keywords; the engine can expand the action library.")}
    ${card(105, 360, 520, 430, `${t("Positive keywords", 150, 430, 360, { size: 34, weight: 800 })}${["repair", "plant", "share", "listen", "teach", "reduce", "organise"].map((d, i) => pill(150 + (i % 2) * 220, 485 + Math.floor(i / 2) * 72, d, C.mint, C.green, 190)).join("")}`, "#f5fbf1")}
    ${card(700, 360, 520, 430, `${t("Negative keywords", 745, 430, 360, { size: 34, weight: 800 })}${["extract", "consume", "ignore", "waste", "delay", "exploit"].map((d, i) => pill(745 + (i % 2) * 220, 485 + Math.floor(i / 2) * 72, d, "#f8dfd9", C.coral, 190)).join("")}`, "#fff4ef")}
    ${card(1290, 330, 520, 515, `${t("Hidden variables", 1338, 400, 360, { size: 34, weight: 800 })}${["Earth Health", "Biodiversity", "Community Care", "Justice", "Waste", "Hope"].map((d, i) => `<text x="1340" y="${465 + i * 61}" font-family="Inter, Arial" font-size="25" font-weight="650" fill="${C.ink}">${esc(d)}</text><rect x="1590" y="${445 + i * 61}" width="150" height="18" rx="9" fill="#e0d3bd"/><rect x="1590" y="${445 + i * 61}" width="${[108, 120, 132, 92, 56, 118][i]}" height="18" rx="9" fill="${i === 4 ? C.coral : C.green}"/>`).join("")}`, "#fbf7ee")}</svg>`,
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(11, "Branching story map", "Multiple paths produce multiple futures.")}
    ${pill(135, 585, "reader keyword", "#fff1d6", C.amber, 300)}
    ${[["repair", 520, 365, "community care", 880, 325, "regenerative ending", 1260, 325, C.green], ["plant", 520, 525, "biodiversity", 880, 525, "blooming ending", 1260, 525, C.green], ["consume", 520, 685, "waste", 880, 725, "collapse ending", 1260, 725, C.coral], ["technology", 520, 845, "efficiency", 880, 895, "unequal green future", 1260, 895, C.blue]].map(([a, ax, ay, b, bx, by, c, cx, cy, col]) => `${pill(ax, ay, a, "#fffaf1", col, 235)}${pill(bx, by, b, "#fffaf1", col, 300)}${pill(cx, cy, c, "#fffaf1", col, 390)}${arrow(435, 612, ax, ay + 27, col)}${arrow(ax + 235, ay + 27, bx, by + 27, col)}${arrow(bx + 300, by + 27, cx, cy + 27, col)}`).join("")}</svg>`,
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(12, "Possible endings / ending card", "The final output is a personal publication record.")}
    ${[["Regenerative Earth", "Repair becomes a repeated pattern strong enough to change the weather of the story.", C.green], ["Fragile Balance", "The Earth remains between two futures, still listening for what repeats.", C.amber], ["Technological Green Future", "Cleaner machines help, but justice and care still decide who benefits.", C.blue], ["Unequal Survival", "Some places are protected while others carry the heat.", C.coral], ["Collapse", "The last page records the cost of delayed care.", C.red]].map((e, i) => {
      const x = i < 3 ? 110 + (i % 3) * 590 : 110 + (i - 3) * 910;
      const y = i < 3 ? 350 : 675;
      const w = i < 3 ? 520 : 800;
      return card(x, y, w, 240, `<circle cx="${x + 54}" cy="${y + 58}" r="18" fill="${e[2]}"/><text x="${x + 92}" y="${y + 70}" font-family="Inter, Arial" font-size="30" font-weight="780" fill="${C.ink}">${esc(e[0])}</text>${t(e[1], x + 52, y + 132, w - 100, { size: 27, fill: C.slate, maxChars: i < 3 ? 33 : 52, lineHeight: 35 })}`, "#fffaf1");
    }).join("")}</svg>`,
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(13, "Why this is a publication", "The website publishes through action, output and archive.")}
    ${["Story text is triggered, not simply displayed.", "The reader is a co-author of the sequence.", "Each keyword is a publishing action.", "The ending card becomes a personal publication.", "The archive becomes a collective publication.", "The website is a living story system."].map((d, i) => {
      const x = 140 + (i % 2) * 820;
      const y = 360 + Math.floor(i / 2) * 165;
      return card(x, y, 700, 115, `<text x="${x + 46}" y="${y + 70}" font-family="Inter, Arial" font-size="30" font-weight="650" fill="${C.ink}">${esc(d)}</text>`, i % 2 ? "#f3f8f0" : "#fff8ea");
    }).join("")}</svg>`,
  () => {
    const landing = imgData(path.join(USER_TESTING_DIR, "user-test-landing-slide.jpg"));
    const flow = imgData(path.join(USER_TESTING_DIR, "user-test-interface-slide.jpg"));
    return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(14, "User testing / iteration evidence", "Two readers tested the prototype and identified unclear interface language.")}
    <image href="${landing}" x="95" y="380" width="790" height="405" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    <image href="${flow}" x="1035" y="380" width="790" height="405" preserveAspectRatio="xMidYMid slice" filter="url(#shadow)"/>
    ${card(150, 825, 760, 180, `${t("Observed issue", 195, 885, 500, { size: 31, weight: 800 })}${t("Interface labels and ending wording felt ambiguous.", 195, 940, 610, { size: 26, fill: C.slate, maxChars: 49, lineHeight: 33 })}`, "#fff8ea")}
    ${card(1010, 825, 760, 180, `${t("Iteration decision", 1055, 885, 500, { size: 31, weight: 800 })}${t("Rename “Collapse Story” to “Collapse” and make state labels clearer.", 1055, 940, 610, { size: 26, fill: C.slate, maxChars: 50, lineHeight: 33 })}`, "#f2f8f0")}</svg>`;
  },
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title(15, "Final deliverables checklist", "This page now maps directly to the brief requirements.")}
    ${["Interactive website prototype: local build + deployment link / QR", "Demo video: 1–2 minutes showing a complete reading path", "Story system map: keyword → variable → branch → ending", "UI screens: desktop and exhibition encounter view", "Ending cards / receipts: digital and/or printed mockups", "Process document: sketches, workshop notes, iterations", "Written reflection: audience, context, format, Radical Earth response"].map((d, i) => {
      const x = i < 4 ? 130 : 1040;
      const y = (i < 4 ? 345 : 420) + (i % 4) * 122;
      return `<rect x="${x}" y="${y}" width="720" height="82" rx="16" fill="${i % 2 ? "#fff8ea" : "#f2f8f0"}" stroke="${C.line}" stroke-width="2"/><circle cx="${x + 45}" cy="${y + 41}" r="16" fill="${i < 2 ? C.red : C.green}"/><text x="${x + 84}" y="${y + 52}" font-family="Inter, Arial" font-size="27" font-weight="650" fill="${C.ink}">${esc(d)}</text>`;
    }).join("")}
    ${t("Red dots = still needs real material or confirmation before final submission.", 1220, 930, 520, { size: 26, fill: C.red, maxChars: 42 })}</svg>`,
  () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}<circle cx="1550" cy="600" r="360" fill="${C.mint}" opacity=".55"/><circle cx="1550" cy="600" r="250" fill="${C.green}" opacity=".20"/>${globe(1550, 600, 210, "repair")}<text x="98" y="92" font-family="Inter, Arial" font-size="24" fill="${C.slate}" letter-spacing="1.8">REFLECTION / RADICAL EARTH RESPONSE</text>${t("Small words, big systems", 98, 220, 810, { size: 76, weight: 790, maxChars: 18, lineHeight: 84 })}${t("This project treats publishing as a responsive system rather than a fixed page. Designed for World Earth Day participants, WSA design students and exhibition visitors, the website turns reading into a sequence of small ecological actions.", 104, 395, 860, { size: 32, fill: C.ink, maxChars: 56, lineHeight: 42 })}${t("Each keyword changes the story, the visual state of the Earth and the reader’s possible ending. The format is digital, but it can become spatial through a projected interface and printed ending cards. In response to Radical Earth, Earth is framed as a living system shaped by language, responsibility, repair and collective consequence.", 104, 610, 860, { size: 31, fill: C.slate, maxChars: 58, lineHeight: 40 })}</svg>`,
];

async function renderSlides() {
  for (let i = 0; i < slides.length; i += 1) {
    await sharp(Buffer.from(slides[i]())).png({ compressionLevel: 9 }).toFile(path.join(SLIDES_DIR, `slide-${String(i + 1).padStart(2, "0")}.png`));
  }
}

async function contactSheet() {
  const tw = 480;
  const th = 270;
  const comps = [];
  for (let i = 1; i <= slides.length; i += 1) {
    const input = await sharp(path.join(SLIDES_DIR, `slide-${String(i).padStart(2, "0")}.png`)).resize(tw, th).png().toBuffer();
    comps.push({ input, left: ((i - 1) % 4) * tw, top: Math.floor((i - 1) / 4) * th });
  }
  await sharp({ create: { width: tw * 4, height: th * 4, channels: 4, background: C.paper } }).composite(comps).png().toFile(path.join(QA_DIR, "contact-sheet-v2.png"));
}

async function buildPptx() {
  const pptx = new pptxgen();
  pptx.author = "Codex";
  pptx.subject = "Radical Earth interactive web publication";
  pptx.title = "The Earth, Written by Us";
  pptx.company = "MACD";
  pptx.lang = "en-GB";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "en-GB",
  };
  pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333333, height: 7.5 });
  pptx.layout = "CUSTOM_WIDE";
  for (let i = 1; i <= slides.length; i += 1) {
    const slide = pptx.addSlide();
    slide.background = { color: "F6F0E6" };
    slide.addImage({
      path: path.join(SLIDES_DIR, `slide-${String(i).padStart(2, "0")}.png`),
      x: 0,
      y: 0,
      w: 13.333333,
      h: 7.5,
    });
  }
  await pptx.writeFile({ fileName: PPTX_OUT });
}

function html() {
  const slideItems = slides
    .map((_, i) => {
      const n = String(i + 1).padStart(2, "0");
      return `<section class="slide" id="slide-${n}" aria-label="Slide ${i + 1} of ${slides.length}">
        <img src="../slides-v2/slide-${n}.png" alt="Slide ${i + 1}" />
      </section>`;
    })
    .join("\n");

  const navItems = slides
    .map((_, i) => {
      const n = String(i + 1).padStart(2, "0");
      return `<a href="#slide-${n}" aria-label="Go to slide ${i + 1}">${i + 1}</a>`;
    })
    .join("");

  const doc = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Publishing Forms: Radical Earth — The Earth, Written by Us</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #181f22;
      --paper: #f6f0e6;
      --ink: #10212b;
      --muted: #6b767a;
    }
    * { box-sizing: border-box; }
    html { scroll-behavior: smooth; }
    body {
      margin: 0;
      font-family: Inter, "Avenir Next", "Helvetica Neue", Arial, sans-serif;
      background: var(--bg);
      color: var(--paper);
    }
    .topbar {
      position: fixed;
      z-index: 5;
      top: 0;
      left: 0;
      right: 0;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      padding: 12px 18px;
      background: rgba(16, 33, 43, 0.82);
      backdrop-filter: blur(10px);
      border-bottom: 1px solid rgba(246, 240, 230, 0.12);
    }
    .topbar strong { font-size: 15px; letter-spacing: .02em; }
    .topbar span { color: rgba(246, 240, 230, .74); font-size: 13px; }
    .deck {
      min-height: 100vh;
      scroll-snap-type: y mandatory;
      overflow-y: auto;
    }
    .slide {
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 58px 18px 26px;
      scroll-snap-align: start;
    }
    .slide img {
      width: min(100%, calc((100vh - 94px) * 16 / 9));
      max-height: calc(100vh - 94px);
      aspect-ratio: 16 / 9;
      object-fit: contain;
      border-radius: 10px;
      box-shadow: 0 24px 70px rgba(0,0,0,.42);
      background: var(--paper);
    }
    .nav {
      position: fixed;
      z-index: 6;
      left: 50%;
      bottom: 14px;
      transform: translateX(-50%);
      display: flex;
      gap: 5px;
      padding: 8px;
      border-radius: 999px;
      background: rgba(16, 33, 43, 0.82);
      backdrop-filter: blur(10px);
    }
    .nav a {
      width: 26px;
      height: 26px;
      display: grid;
      place-items: center;
      border-radius: 50%;
      text-decoration: none;
      color: var(--paper);
      font-size: 12px;
      font-weight: 700;
    }
    .nav a:hover, .nav a:focus-visible {
      outline: none;
      background: #f6f0e6;
      color: var(--ink);
    }
    .hint { display: inline; }
    @media (max-width: 760px) {
      .topbar { align-items: flex-start; flex-direction: column; padding: 10px 14px; }
      .hint { display: none; }
      .slide { padding-top: 76px; }
      .nav { max-width: calc(100vw - 24px); overflow-x: auto; }
    }
    @media print {
      body { background: white; }
      .topbar, .nav { display: none; }
      .deck { overflow: visible; }
      .slide {
        min-height: auto;
        padding: 0;
        page-break-after: always;
      }
      .slide img {
        width: 100%;
        max-height: none;
        border-radius: 0;
        box-shadow: none;
      }
    }
  </style>
</head>
<body>
  <header class="topbar">
    <strong>Publishing Forms: Radical Earth</strong>
    <span>The Earth, Written by Us <span class="hint">· use ← / → keys to navigate</span></span>
  </header>
  <main class="deck">
${slideItems}
  </main>
  <nav class="nav" aria-label="Slide navigation">${navItems}</nav>
  <script>
    const slides = Array.from(document.querySelectorAll(".slide"));
    function currentIndex() {
      const center = window.scrollY + window.innerHeight / 2;
      let best = 0;
      let bestDistance = Infinity;
      slides.forEach((slide, index) => {
        const middle = slide.offsetTop + slide.offsetHeight / 2;
        const distance = Math.abs(center - middle);
        if (distance < bestDistance) {
          best = index;
          bestDistance = distance;
        }
      });
      return best;
    }
    window.addEventListener("keydown", (event) => {
      if (!["ArrowRight", "ArrowDown", "PageDown", "ArrowLeft", "ArrowUp", "PageUp", "Home", "End"].includes(event.key)) return;
      event.preventDefault();
      let next = currentIndex();
      if (["ArrowRight", "ArrowDown", "PageDown"].includes(event.key)) next += 1;
      if (["ArrowLeft", "ArrowUp", "PageUp"].includes(event.key)) next -= 1;
      if (event.key === "Home") next = 0;
      if (event.key === "End") next = slides.length - 1;
      slides[Math.max(0, Math.min(slides.length - 1, next))].scrollIntoView({ behavior: "smooth" });
    });
  </script>
</body>
</html>`;
  fs.writeFileSync(OUT, doc);
}

function missingList() {
  const md = `# Updated Missing Material To Fill

Already added:
- Earth Day Experiment Log 04: Power of Positive Tipping Points / Systems Games Workshop.
- Workshop date/place/focus: 22 April 2026, West Side Lecture Theatre, systems thinking, tipping points, adaptation, feedback.
- Workshop process photos showing floor-card participation, group activity, positive tipping points, feedback loops and complex-system slides.
- User testing photos showing the prototype being used on laptop.
- User feedback added: some interface meanings were unclear; simplify ending language from "Collapse Story" to "Collapse".

1. Original given content
   - Title / author / source of the assigned text.
   - 3-5 core ideas or arguments.
   - Tone of voice, key phrases, recurring motifs, narrative rhythm.
   - Confirmation: full text, selected fragments, rewritten chapters, or intentional edited version.

2. Source-to-story mapping
   - Which exact lines or ideas from the original text become keywords, chapters, Earth states and endings.
   - A short paragraph explaining why the text is edited into an interactive story rather than reproduced in full.

3. Process and iteration evidence
   - Early sketches, failed UI versions, system diagrams, paper prototypes beyond the Earth Day workshop and user-testing evidence.
   - More detailed user testing notes: who tested, task given, exact quotes, what changed after each test.
   - Before/after screenshots showing what changed.

4. Final publication evidence
   - Public prototype link or QR code.
   - Demo video still or 1-2 minute demo video.
   - Optional exhibition mock installation view.
   - Printed/digital ending-card mockups.

5. Optional strengthening
   - Direct participant quotes or recorded workshop keywords, if available.
   - Your own wording for audience, context, format, and Radical Earth response.
   - Any citations or references from the source text / workshop.
`;
  fs.writeFileSync(path.join(ROOT, "missing-materials-to-fill.md"), md);
}

(async () => {
  missingList();
  await renderSlides();
  await contactSheet();
  await buildPptx();
  html();
  console.log(OUT);
  console.log(PPTX_OUT);
})();
