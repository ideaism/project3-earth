const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const pptxgen = require("pptxgenjs");

const ROOT = __dirname;
const ASSET_DIR = path.join(ROOT, "assets");
const SLIDES_DIR = path.join(ROOT, "slides");
const PREVIEW_DIR = path.join(ROOT, "preview");
const QA_DIR = path.join(ROOT, "qa");
const OUTPUT_DIR = path.join(ROOT, "output");
const OUT = path.join(OUTPUT_DIR, "a-website-that-tips-earth-story.pptx");

for (const dir of [ASSET_DIR, SLIDES_DIR, PREVIEW_DIR, QA_DIR, OUTPUT_DIR]) {
  fs.mkdirSync(dir, { recursive: true });
}

const W = 1920;
const H = 1080;
const C = {
  ink: "#10212b",
  ink2: "#17313c",
  green: "#4f8f68",
  moss: "#8fb37f",
  mint: "#d8eadc",
  paper: "#f6f0e6",
  paper2: "#ece0cf",
  coral: "#e86f5c",
  amber: "#e6b85c",
  blue: "#577d91",
  slate: "#4d5c62",
  smoke: "#6f6f68",
  white: "#fffaf1",
  red: "#a44942",
};

function esc(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function imgData(file) {
  const ext = path.extname(file).toLowerCase().replace(".", "") || "png";
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

function textBlock(text, x, y, width, opts = {}) {
  const size = opts.size || 38;
  const fill = opts.fill || C.ink;
  const weight = opts.weight || 400;
  const maxChars = opts.maxChars || Math.max(16, Math.floor(width / (size * 0.48)));
  const lineHeight = opts.lineHeight || Math.round(size * 1.22);
  const anchor = opts.anchor || "start";
  const family = opts.family || "Inter, Avenir Next, Helvetica Neue, Arial, sans-serif";
  const lines = Array.isArray(text) ? text : wrap(text, maxChars);
  return `<text x="${x}" y="${y}" fill="${fill}" font-family="${family}" font-size="${size}" font-weight="${weight}" text-anchor="${anchor}">${lines
    .map((line, i) => `<tspan x="${x}" dy="${i === 0 ? 0 : lineHeight}">${esc(line)}</tspan>`)
    .join("")}</text>`;
}

function bg() {
  return `
    <defs>
      <radialGradient id="paperGlow" cx="28%" cy="12%" r="92%">
        <stop offset="0%" stop-color="#fffaf1"/>
        <stop offset="58%" stop-color="#f4ebdc"/>
        <stop offset="100%" stop-color="#e8dccb"/>
      </radialGradient>
      <pattern id="grain" width="96" height="96" patternUnits="userSpaceOnUse">
        <rect width="96" height="96" fill="none"/>
        <circle cx="12" cy="18" r="1.4" fill="#d8c9b4" opacity=".28"/>
        <circle cx="44" cy="52" r="1.1" fill="#cdbb9f" opacity=".20"/>
        <circle cx="82" cy="31" r="1.2" fill="#e3d4bd" opacity=".32"/>
        <path d="M0 78 Q24 70 48 80 T96 79" stroke="#dacbb5" stroke-width="1" opacity=".12" fill="none"/>
      </pattern>
      <filter id="softShadow" x="-20%" y="-20%" width="140%" height="140%">
        <feDropShadow dx="0" dy="18" stdDeviation="22" flood-color="#10212b" flood-opacity=".16"/>
      </filter>
      <linearGradient id="darkGrad" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="#10212b" stop-opacity=".82"/>
        <stop offset="58%" stop-color="#10212b" stop-opacity=".28"/>
        <stop offset="100%" stop-color="#f6f0e6" stop-opacity=".08"/>
      </linearGradient>
    </defs>
    <rect width="${W}" height="${H}" fill="url(#paperGlow)"/>
    <rect width="${W}" height="${H}" fill="url(#grain)" opacity=".72"/>
    <path d="M-30 1010 C300 920 520 1060 850 980 C1210 890 1450 990 1950 910 L1950 1100 L-30 1100 Z" fill="#dfe9df" opacity=".44"/>
    <path d="M1500 -130 C1710 25 1775 220 1950 310" stroke="#d7c8ae" stroke-width="2" fill="none" opacity=".42"/>
  `;
}

function chrome(n, kicker = "Radical Earth Publishing Form") {
  return `
    <text x="98" y="70" font-family="Inter, Avenir Next, Helvetica Neue, Arial, sans-serif" font-size="24" fill="${C.slate}" letter-spacing="1.8">${esc(kicker.toUpperCase())}</text>
    <text x="1822" y="70" font-family="Inter, Avenir Next, Helvetica Neue, Arial, sans-serif" font-size="24" fill="${C.slate}" text-anchor="end">${String(n).padStart(2, "0")} / 16</text>
  `;
}

function title(t, st, n) {
  const titleLines = wrap(t, 26);
  const subY = 168 + (titleLines.length - 1) * 82 + 106;
  return `${chrome(n)}${textBlock(titleLines, 96, 168, 980, { size: 74, weight: 760, lineHeight: 82 })}${st ? textBlock(st, 100, subY, 1180, { size: 29, fill: C.slate, maxChars: 72, lineHeight: 37 }) : ""}`;
}

function pill(x, y, label, fill = C.mint, stroke = C.green, w = null) {
  const ww = w || Math.max(112, label.length * 15 + 42);
  return `<rect x="${x}" y="${y}" width="${ww}" height="54" rx="27" fill="${fill}" stroke="${stroke}" stroke-width="2"/><text x="${x + ww / 2}" y="${y + 35}" font-family="Inter, Avenir Next, Helvetica Neue, Arial, sans-serif" font-size="24" font-weight="650" fill="${C.ink}" text-anchor="middle">${esc(label)}</text>`;
}

function card(x, y, w, h, body, opts = {}) {
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="18" fill="${opts.fill || C.white}" stroke="${opts.stroke || "#d8cbb6"}" stroke-width="2" filter="${opts.shadow === false ? "" : "url(#softShadow)"}"/>${body}`;
}

function globe(cx, cy, r, state = "repair") {
  const base = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#cfe2d3" stroke="${C.ink}" stroke-width="6"/>
    <path d="M${cx - r * .55} ${cy - r * .08} C${cx - r * .3} ${cy - r * .5} ${cx + r * .12} ${cy - r * .4} ${cx + r * .22} ${cy - r * .12} C${cx + r * .38} ${cy + r * .22} ${cx - r * .08} ${cy + r * .2} ${cx - r * .04} ${cy + r * .5}" fill="${C.green}" opacity=".9"/>
    <path d="M${cx + r * .2} ${cy - r * .62} C${cx + r * .62} ${cy - r * .35} ${cx + r * .44} ${cy + r * .1} ${cx + r * .72} ${cy + r * .28}" stroke="${C.blue}" stroke-width="24" stroke-linecap="round" fill="none" opacity=".42"/>`;
  const variants = {
    fragile: `<path d="M${cx - 25} ${cy - 95} L${cx + 12} ${cy - 40} L${cx - 10} ${cy + 20} L${cx + 38} ${cy + 86}" stroke="${C.red}" stroke-width="7" fill="none"/>`,
    warming: `<circle cx="${cx}" cy="${cy}" r="${r - 8}" fill="${C.coral}" opacity=".22"/><path d="M${cx - 100} ${cy + 112} Q${cx} ${cy + 70} ${cx + 100} ${cy + 112}" stroke="${C.coral}" stroke-width="10" fill="none"/>`,
    repair: `<path d="M${cx - 118} ${cy + 20} C${cx - 38} ${cy - 54} ${cx + 50} ${cy + 46} ${cx + 124} ${cy - 32}" stroke="${C.amber}" stroke-width="9" stroke-linecap="round" fill="none"/><circle cx="${cx + 118}" cy="${cy - 32}" r="13" fill="${C.amber}"/>`,
    bloom: `<g fill="${C.moss}"><circle cx="${cx - 74}" cy="${cy - 72}" r="15"/><circle cx="${cx + 70}" cy="${cy + 38}" r="18"/><circle cx="${cx + 25}" cy="${cy - 105}" r="13"/></g>`,
    collective: `<g stroke="${C.amber}" stroke-width="4" fill="${C.amber}"><circle cx="${cx - 85}" cy="${cy - 70}" r="9"/><circle cx="${cx + 75}" cy="${cy - 40}" r="9"/><circle cx="${cx - 30}" cy="${cy + 90}" r="9"/><path d="M${cx - 85} ${cy - 70} L${cx + 75} ${cy - 40} L${cx - 30} ${cy + 90} Z" fill="none"/></g>`,
    collapse: `<circle cx="${cx}" cy="${cy}" r="${r - 6}" fill="#263238" opacity=".35"/><path d="M${cx - 110} ${cy - 62} L${cx - 22} ${cy + 10} L${cx - 72} ${cy + 92} M${cx + 24} ${cy - 96} L${cx + 72} ${cy + 4} L${cx + 38} ${cy + 96}" stroke="#2c3234" stroke-width="9" fill="none"/>`,
  };
  return `<g filter="url(#softShadow)">${base}${variants[state] || variants.repair}</g>`;
}

function arrow(x1, y1, x2, y2, color = C.ink) {
  return `<path d="M${x1} ${y1} C${(x1 + x2) / 2} ${y1}, ${(x1 + x2) / 2} ${y2}, ${x2} ${y2}" stroke="${color}" stroke-width="5" fill="none" marker-end="url(#arrow)"/>`;
}

const defsArrow = `<defs><marker id="arrow" markerWidth="13" markerHeight="13" refX="10" refY="6" orient="auto"><path d="M1,1 L11,6 L1,11 Z" fill="${C.ink}"/></marker></defs>`;

const slides = [
  {
    n: 1,
    svg: () => {
      const hero = imgData(path.join(ASSET_DIR, "title-earth-story.png"));
      return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}<image href="${hero}" x="0" y="0" width="${W}" height="${H}" preserveAspectRatio="xMidYMid slice"/><rect width="${W}" height="${H}" fill="url(#darkGrad)"/><text x="98" y="104" font-family="Inter, Avenir Next, Helvetica Neue, Arial, sans-serif" font-size="27" fill="#f6f0e6" letter-spacing="2">INTERACTIVE WEB PUBLICATION</text>${textBlock("A Website That Tips", 96, 265, 790, { size: 95, fill: C.white, weight: 780, maxChars: 15, lineHeight: 100 })}${textBlock("An Earth Story You Can Change", 100, 472, 700, { size: 38, fill: "#efe4ce", maxChars: 34, lineHeight: 46 })}<rect x="100" y="820" width="500" height="72" rx="36" fill="#f6f0e6" opacity=".92"/><text x="350" y="866" font-family="Inter, Avenir Next, Helvetica Neue, Arial, sans-serif" font-size="27" font-weight="700" fill="${C.ink}" text-anchor="middle">Project presentation deck</text></svg>`;
    },
  },
  {
    n: 2,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Brief Interpretation", "Publishing beyond the static printed page.", 2)}
      ${card(98, 395, 520, 360, `${textBlock("The brief asks us to publish and activate content beyond a fixed printed object.", 140, 470, 430, { size: 38, weight: 650, maxChars: 23, lineHeight: 48 })}`, { fill: "#fff9ee" })}
      ${card(690, 330, 520, 440, `${textBlock("Publication becomes:", 742, 410, 400, { size: 34, weight: 750, maxChars: 21 })}${["interface", "system", "event", "hybrid format"].map((d, i) => pill(742, 475 + i * 70, d, i % 2 ? "#f6dfd8" : C.mint, i % 2 ? C.coral : C.green, 310)).join("")}`, { fill: "#fbf5ea" })}
      ${card(1280, 395, 540, 360, `${textBlock("My response", 1330, 465, 400, { size: 34, weight: 760, maxChars: 20 })}${textBlock("An interactive web publication changed by reader input: every word triggers a story branch and an Earth-state shift.", 1330, 535, 420, { size: 32, maxChars: 26, lineHeight: 42 })}`, { fill: "#f2f8f0" })}
      <path d="M620 575 L690 550 M1210 550 L1280 575" stroke="${C.ink}" stroke-width="5" marker-end="url(#arrow)"/>${defsArrow}</svg>`,
  },
  {
    n: 3,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("What if reading could change the Earth?", "Concept: story input becomes visible ecological consequence.", 3)}
      ${card(102, 390, 760, 520, `${textBlock("Living Earth", 150, 455, 300, { size: 36, weight: 760 })}${globe(475, 660, 168, "repair")}${textBlock("Earth state responds to health, biodiversity, care, justice, waste and hope.", 180, 842, 585, { size: 27, fill: C.slate, maxChars: 52, lineHeight: 34 })}`, { fill: "#f8fbf4" })}
      ${card(1010, 390, 808, 520, `<rect x="1080" y="455" width="660" height="315" rx="12" fill="#fffdf6" stroke="#cbbda6" stroke-width="3"/><line x1="1410" y1="455" x2="1410" y2="770" stroke="#ddd0bb" stroke-width="3"/>${textBlock("Chapter 03", 1135, 520, 230, { size: 28, fill: C.coral, weight: 760 })}${textBlock("The reader types a word. The story answers as a system, not a page.", 1135, 580, 230, { size: 29, maxChars: 17, lineHeight: 38 })}${textBlock("keyword", 1470, 540, 170, { size: 25, fill: C.slate, weight: 700 })}<rect x="1470" y="565" width="210" height="58" rx="29" fill="${C.mint}" stroke="${C.green}" stroke-width="2"/><text x="1575" y="603" text-anchor="middle" font-family="Inter, Arial" font-size="26" font-weight="700" fill="${C.ink}">repair</text>${textBlock("Storybook Dialogue", 1080, 835, 480, { size: 36, weight: 760 })}`, { fill: "#fff7e8" })}
      <path d="M875 635 C945 572 957 572 1000 635" stroke="${C.amber}" stroke-width="9" fill="none" marker-end="url(#arrow)"/>${defsArrow}</svg>`,
  },
  {
    n: 4,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("From World Earth Day Workshop to Story System", "Research context translated into interaction mechanics.", 4)}
      <circle cx="960" cy="610" r="150" fill="${C.ink}" opacity=".95" filter="url(#softShadow)"/><text x="960" y="595" text-anchor="middle" font-family="Inter, Arial" font-size="34" font-weight="760" fill="${C.white}">Story</text><text x="960" y="640" text-anchor="middle" font-family="Inter, Arial" font-size="34" font-weight="760" fill="${C.white}">System</text>
      ${[
        ["positive tipping points", 430, 360, C.mint, C.green],
        ["systems thinking", 1240, 360, "#e7eef1", C.blue],
        ["small actions", 350, 755, "#fff1d6", C.amber],
        ["collective change", 1215, 755, "#f8dfd9", C.coral],
      ].map(([l, x, y, f, s]) => `${pill(x, y, l, f, s, 360)}<path d="M${x + 180} ${y + 58} Q960 ${y < 500 ? 470 : 760} 960 610" stroke="${s}" stroke-width="4" fill="none" opacity=".72"/>`).join("")}
      ${textBlock("Workshop language becomes game logic: small choices accumulate, feedback is visible, and collective repair is treated as a possible tipping point.", 380, 890, 1160, { size: 33, fill: C.slate, maxChars: 70, lineHeight: 42 })}</svg>`,
  },
  {
    n: 5,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Who is this for?", "Audience and context define how the publication is encountered.", 5)}
      <text x="130" y="385" font-family="Inter, Arial" font-size="35" font-weight="760" fill="${C.ink}">Audience</text>
      ${["design students", "workshop participants", "climate-aware young people", "exhibition visitors"].map((d, i) => card(110, 435 + i * 112, 760, 82, `<text x="158" y="${487 + i * 112}" font-family="Inter, Arial" font-size="31" font-weight="650" fill="${C.ink}">${esc(d)}</text>`, { fill: i % 2 ? "#f9f2e6" : "#f2f8f0", shadow: false })).join("")}
      <text x="1050" y="385" font-family="Inter, Arial" font-size="35" font-weight="760" fill="${C.ink}">Context</text>
      ${["World Earth Day event", "studio exhibition", "online publication", "interactive installation"].map((d, i) => card(1030, 435 + i * 112, 760, 82, `<text x="1078" y="${487 + i * 112}" font-family="Inter, Arial" font-size="31" font-weight="650" fill="${C.ink}">${esc(d)}</text>`, { fill: i % 2 ? "#fff1d6" : "#e7eef1", shadow: false })).join("")}
      <path d="M930 418 L930 890" stroke="#c8bca8" stroke-width="3" stroke-dasharray="10 12"/></svg>`,
  },
  {
    n: 6,
    svg: () => {
      const steps = ["Enter", "Read prologue", "Type keyword", "Earth changes", "Story branches", "Repeat", "Ending", "Archive"];
      return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${defsArrow}${title("How the reader enters the story", "User experience flow.", 6)}
      ${steps.map((s, i) => {
        const x = 115 + (i % 4) * 435;
        const y = i < 4 ? 455 : 705;
        const f = [C.mint, "#fff1d6", "#f8dfd9", "#e7eef1"][i % 4];
        return `${card(x, y, 310, 118, `<text x="${x + 155}" y="${y + 72}" text-anchor="middle" font-family="Inter, Arial" font-size="28" font-weight="760" fill="${C.ink}">${esc(s)}</text>`, { fill: f, shadow: false })}${i < steps.length - 1 ? arrow(x + 310, y + 59, (i === 3 ? 115 : x + 435), i === 3 ? 764 : y + 59) : ""}`;
      }).join("")}
      ${textBlock("The loop matters: the publication is experienced through repeated reading, input, feedback and revision.", 330, 935, 1260, { size: 32, fill: C.slate, maxChars: 72 })}</svg>`;
    },
  },
  {
    n: 7,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Two-part Interface", "Left: Living Earth. Right: Storybook Dialogue.", 7)}
      ${card(110, 330, 760, 600, `${textBlock("Living Earth", 155, 395, 300, { size: 33, weight: 760 })}${globe(490, 635, 165, "collective")}${pill(222, 835, "Earth state visual", C.mint, C.green, 360)}${pill(535, 415, "Progress indicator", "#fff1d6", C.amber, 280)}`, { fill: "#f6fbf2" })}
      ${card(1010, 330, 800, 600, `<rect x="1085" y="405" width="650" height="350" rx="12" fill="#fffdf6" stroke="#cbbda6" stroke-width="3"/>${textBlock("Chapter title", 1132, 470, 240, { size: 27, weight: 760, fill: C.coral })}${textBlock("Story text answers the reader's keyword and changes the next choices.", 1132, 530, 510, { size: 30, maxChars: 38, lineHeight: 40 })}<rect x="1132" y="682" width="390" height="58" rx="29" fill="#f7efe2" stroke="#d2c3ad" stroke-width="2"/><text x="1160" y="720" font-family="Inter, Arial" font-size="25" fill="${C.slate}">type a keyword...</text>${pill(1560, 682, "repair", C.mint, C.green, 150)}${pill(1118, 805, "Action chips", "#fff1d6", C.amber, 230)}${pill(1385, 805, "Ending status", "#f8dfd9", C.coral, 250)}`, { fill: "#fff7e8" })}</svg>`,
  },
  {
    n: 8,
    svg: () => {
      const img = imgData(path.join(ASSET_DIR, "earth-states-generated.png"));
      const labels = ["Fragile", "Warming", "Repairing", "Blooming", "Collective", "Collapsing"];
      return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("The Earth as a Living Page", "Each state has a distinct visual, colour and story tone.", 8)}
        <image href="${img}" x="105" y="350" width="1710" height="550" preserveAspectRatio="xMidYMid meet" filter="url(#softShadow)"/>
        ${labels.map((l, i) => `<text x="${235 + i * 291}" y="938" text-anchor="middle" font-family="Inter, Arial" font-size="27" font-weight="760" fill="${i === 5 ? C.red : C.ink}">${l} Earth</text>`).join("")}
      </svg>`;
    },
  },
  {
    n: 9,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Every Word Changes the System", "Keywords affect hidden variables behind the story.", 9)}
      ${card(110, 360, 520, 440, `${textBlock("Positive keywords", 155, 428, 360, { size: 34, weight: 760 })}${["repair", "plant", "share", "listen", "teach", "reduce", "organise"].map((d, i) => pill(155 + (i % 2) * 220, 480 + Math.floor(i / 2) * 72, d, C.mint, C.green, 190)).join("")}`, { fill: "#f5fbf1" })}
      ${card(700, 360, 520, 440, `${textBlock("Negative keywords", 745, 428, 360, { size: 34, weight: 760 })}${["extract", "consume", "ignore", "waste", "delay", "exploit"].map((d, i) => pill(745 + (i % 2) * 220, 480 + Math.floor(i / 2) * 72, d, "#f8dfd9", C.coral, 190)).join("")}`, { fill: "#fff4ef" })}
      ${card(1290, 330, 520, 515, `${textBlock("Hidden variables", 1338, 400, 360, { size: 34, weight: 760 })}${["Earth Health", "Biodiversity", "Community Care", "Justice", "Waste", "Hope"].map((d, i) => `<text x="1340" y="${465 + i * 61}" font-family="Inter, Arial" font-size="25" font-weight="650" fill="${C.ink}">${esc(d)}</text><rect x="1590" y="${445 + i * 61}" width="150" height="18" rx="9" fill="#e0d3bd"/><rect x="1590" y="${445 + i * 61}" width="${[108, 120, 132, 92, 56, 118][i]}" height="18" rx="9" fill="${i === 4 ? C.coral : C.green}"/>`).join("")}`, { fill: "#fbf7ee" })}</svg>`,
  },
  {
    n: 10,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${defsArrow}${title("Multiple Paths, Multiple Futures", "A simple branching model turns language into consequences.", 10)}
      ${pill(160, 585, "reader keyword", "#fff1d6", C.amber, 300)}
      ${[["repair", 560, 375, "community care", 920, 330, "regenerative ending", 1290, 330, C.green], ["plant", 560, 535, "biodiversity", 920, 535, "blooming ending", 1290, 535, C.moss], ["consume", 560, 695, "waste", 920, 740, "collapse ending", 1290, 740, C.coral], ["technology", 560, 855, "efficiency", 920, 900, "unequal green future", 1290, 900, C.blue]].map(([a, ax, ay, b, bx, by, c, cx, cy, col]) => `${pill(ax, ay, a, "#fffaf1", col, 230)}${pill(bx, by, b, "#fffaf1", col, 285)}${pill(cx, cy, c, "#fffaf1", col, 360)}${arrow(460, 612, ax, ay + 27, col)}${arrow(ax + 230, ay + 27, bx, by + 27, col)}${arrow(bx + 285, by + 27, cx, cy + 27, col)}`).join("")}</svg>`,
  },
  {
    n: 11,
    svg: () => {
      const endings = [
        ["Regenerative Earth", "A repaired river becomes a meeting place; each voice adds another root.", C.green],
        ["Fragile Balance", "The Earth keeps breathing, but every future choice matters more than before.", C.amber],
        ["Technological Green Future", "Efficient machines cool the city, while the question of access remains open.", C.blue],
        ["Unequal Survival", "Some shelters glow at night; outside them, the weather is still writing.", C.coral],
        ["Collapse Story", "The last page is not silent. It records the cost of delayed care.", C.red],
      ];
      return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Possible Endings", "The publication outputs a personal ending card.", 11)}
      ${endings.map((e, i) => {
        const x = 110 + (i % 3) * 590;
        const y = i < 3 ? 355 : 690;
        const w = i < 3 ? 520 : 800;
        return card(x, y, w, 245, `<circle cx="${x + 54}" cy="${y + 58}" r="18" fill="${e[2]}"/><text x="${x + 92}" y="${y + 70}" font-family="Inter, Arial" font-size="30" font-weight="780" fill="${C.ink}">${esc(e[0])}</text>${textBlock(e[1], x + 52, y + 132, w - 100, { size: 28, fill: C.slate, maxChars: i < 3 ? 31 : 48, lineHeight: 36 })}`, { fill: "#fffaf1", shadow: false });
      }).join("")}</svg>`;
    },
  },
  {
    n: 12,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Why this is a publication", "Publishing strategy: text becomes action, output and archive.", 12)}
      ${["Story text is triggered, not simply displayed.", "The reader is a co-author of the sequence.", "Each keyword is a publishing action.", "The ending card becomes a personal publication.", "The archive becomes a collective publication.", "The website is a living story system."].map((d, i) => {
        const x = 140 + (i % 2) * 820;
        const y = 360 + Math.floor(i / 2) * 165;
        return card(x, y, 700, 115, `<text x="${x + 46}" y="${y + 70}" font-family="Inter, Arial" font-size="30" font-weight="650" fill="${C.ink}">${esc(d)}</text>`, { fill: i % 2 ? "#f3f8f0" : "#fff8ea", shadow: false });
      }).join("")}</svg>`,
  },
  {
    n: 13,
    svg: () => {
      const tiles = [
        ["storybook page", "#fff8ea"],
        ["ecological globe", "#edf8ef"],
        ["soft paper texture", "#f1e3cf"],
        ["system diagram", "#e8eef1"],
        ["glowing repair marks", "#fff1d6"],
        ["collapse cracks", "#f6ded7"],
      ];
      return `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Storybook x Earth System", "Visual direction moodboard.", 13)}
      ${tiles.map((t, i) => {
        const x = 120 + (i % 3) * 570;
        const y = 335 + Math.floor(i / 3) * 275;
        const icons = [
          `<rect x="${x + 64}" y="${y + 64}" width="130" height="150" rx="8" fill="#fffdf6" stroke="#cbbda6" stroke-width="5"/><line x1="${x + 129}" y1="${y + 64}" x2="${x + 129}" y2="${y + 214}" stroke="#d9cbb6" stroke-width="4"/>`,
          globe(x + 130, y + 140, 76, "bloom"),
          `<path d="M${x + 55} ${y + 110} C${x + 145} ${y + 80} ${x + 190} ${y + 160} ${x + 285} ${y + 115}" stroke="#cbbda6" stroke-width="3" fill="none"/><path d="M${x + 62} ${y + 162} C${x + 148} ${y + 128} ${x + 222} ${y + 208} ${x + 310} ${y + 170}" stroke="#d7c8ae" stroke-width="3" fill="none"/>`,
          `<circle cx="${x + 100}" cy="${y + 105}" r="16" fill="${C.blue}"/><circle cx="${x + 230}" cy="${y + 82}" r="16" fill="${C.green}"/><circle cx="${x + 260}" cy="${y + 188}" r="16" fill="${C.coral}"/><path d="M${x + 100} ${y + 105} L${x + 230} ${y + 82} L${x + 260} ${y + 188} L${x + 100} ${y + 105}" stroke="${C.ink}" stroke-width="4" fill="none"/>`,
          `<path d="M${x + 70} ${y + 170} C${x + 145} ${y + 72} ${x + 220} ${y + 218} ${x + 310} ${y + 100}" stroke="${C.amber}" stroke-width="13" fill="none" stroke-linecap="round"/><circle cx="${x + 310}" cy="${y + 100}" r="18" fill="${C.amber}"/>`,
          `<path d="M${x + 85} ${y + 60} L${x + 160} ${y + 135} L${x + 120} ${y + 218} M${x + 245} ${y + 55} L${x + 210} ${y + 150} L${x + 300} ${y + 225}" stroke="${C.red}" stroke-width="8" fill="none"/>`,
        ][i];
        return card(x, y, 480, 220, `${icons}<text x="${x + 54}" y="${y + 185}" font-family="Inter, Arial" font-size="27" font-weight="760" fill="${C.ink}">${esc(t[0])}</text>`, { fill: t[1], shadow: false });
      }).join("")}
      ${textBlock("Visual keywords: fragile, living, poetic, systemic, participatory.", 280, 975, 1360, { size: 34, fill: C.slate, maxChars: 64 })}</svg>`;
    },
  },
  {
    n: 14,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("MVP Prototype", "A focused build plan for the final publishing form.", 14)}
      <line x1="250" y1="590" x2="1660" y2="590" stroke="${C.ink}" stroke-width="6"/>
      ${[["desktop prototype", "screen layout + input loop"], ["5 earth states", "visual feedback system"], ["12 keywords", "positive and negative actions"], ["3-5 endings", "story response library"], ["ending card", "shareable personal output"], ["optional archive", "collective publication"]].map((d, i) => {
        const x = 255 + i * 280;
        return `<circle cx="${x}" cy="590" r="25" fill="${i % 2 ? C.amber : C.green}" stroke="${C.ink}" stroke-width="4"/><text x="${x}" y="680" text-anchor="middle" font-family="Inter, Arial" font-size="28" font-weight="760" fill="${C.ink}">${esc(d[0])}</text>${textBlock(d[1], x - 105, 730, 210, { size: 23, fill: C.slate, anchor: "start", maxChars: 16, lineHeight: 30 })}`;
      }).join("")}</svg>`,
  },
  {
    n: 15,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}${title("Final Publishing Form", "Deliverables for assessment and exhibition.", 15)}
      ${["interactive website prototype", "UI screens", "story system map", "ending cards / receipts", "process documentation", "demo video", "reflection"].map((d, i) => {
        const angle = -90 + i * (360 / 7);
        const r = 280;
        const x = 960 + Math.cos(angle * Math.PI / 180) * r;
        const y = 640 + Math.sin(angle * Math.PI / 180) * r;
        return `<path d="M960 640 L${x} ${y}" stroke="#c9bca6" stroke-width="4"/><circle cx="${x}" cy="${y}" r="82" fill="${i % 3 === 0 ? C.mint : i % 3 === 1 ? "#fff1d6" : "#f8dfd9"}" stroke="${C.ink}" stroke-width="4"/><text x="${x}" y="${y - 8}" text-anchor="middle" font-family="Inter, Arial" font-size="23" font-weight="760" fill="${C.ink}">${wrap(d, 14).map((line, j) => `<tspan x="${x}" dy="${j === 0 ? 0 : 28}">${esc(line)}</tspan>`).join("")}</text>`;
      }).join("")}
      <circle cx="960" cy="640" r="118" fill="${C.ink}" filter="url(#softShadow)"/><text x="960" y="626" text-anchor="middle" font-family="Inter, Arial" font-size="32" font-weight="780" fill="${C.white}">Publishing</text><text x="960" y="670" text-anchor="middle" font-family="Inter, Arial" font-size="32" font-weight="780" fill="${C.white}">System</text></svg>`,
  },
  {
    n: 16,
    svg: () => `<svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg">${bg()}<circle cx="1550" cy="600" r="360" fill="${C.mint}" opacity=".55"/><circle cx="1550" cy="600" r="250" fill="${C.green}" opacity=".20"/>${globe(1550, 600, 210, "repair")}<text x="98" y="92" font-family="Inter, Arial" font-size="24" fill="${C.slate}" letter-spacing="1.8">REFLECTION / RADICAL EARTH RESPONSE</text>${textBlock("Small words, big systems", 98, 235, 810, { size: 82, weight: 780, maxChars: 18, lineHeight: 90 })}${textBlock("This project does not use Earth as a background pattern. It treats Earth as a system changed by choice, language and action.", 104, 430, 820, { size: 38, fill: C.ink2, maxChars: 42, lineHeight: 50 })}${textBlock("By typing words, readers participate in a publishing experiment about repair, consumption, responsibility and possible futures.", 104, 635, 790, { size: 34, fill: C.slate, maxChars: 46, lineHeight: 44 })}</svg>`,
  },
];

async function renderSlides() {
  for (const slide of slides) {
    const svg = slide.svg();
    const out = path.join(SLIDES_DIR, `slide-${String(slide.n).padStart(2, "0")}.png`);
    await sharp(Buffer.from(svg)).png({ compressionLevel: 9 }).toFile(out);
  }
}

async function renderContactSheet() {
  const thumbs = [];
  const tw = 480;
  const th = 270;
  for (const slide of slides) {
    const file = path.join(SLIDES_DIR, `slide-${String(slide.n).padStart(2, "0")}.png`);
    const buffer = await sharp(file).resize(tw, th).png().toBuffer();
    thumbs.push({ input: buffer, top: Math.floor((slide.n - 1) / 4) * th, left: ((slide.n - 1) % 4) * tw });
  }
  await sharp({
    create: {
      width: tw * 4,
      height: th * 4,
      channels: 4,
      background: C.paper,
    },
  }).composite(thumbs).png().toFile(path.join(QA_DIR, "contact-sheet.png"));
}

async function buildPptx() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Codex";
  pptx.subject = "Interactive web publication project presentation";
  pptx.title = "A Website That Tips";
  pptx.company = "MACD";
  pptx.lang = "en-GB";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "en-GB",
  };
  pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333333, height: 7.5 });
  pptx.layout = "CUSTOM_WIDE";
  for (const slide of slides) {
    const s = pptx.addSlide();
    s.background = { color: "F6F0E6" };
    s.addImage({
      path: path.join(SLIDES_DIR, `slide-${String(slide.n).padStart(2, "0")}.png`),
      x: 0,
      y: 0,
      w: 13.333333,
      h: 7.5,
    });
  }
  await pptx.writeFile({ fileName: OUT });
}

async function writeNotes() {
  const notes = `Task mode: create
Primary deck profile: consumer-retail / image-led design presentation
Source: /Users/beijixinfei/Downloads/03_Publishing_Forms(1) (1).pdf

Brief interpretation:
- Publish and activate content beyond the static printed page.
- Treat publishing as an event, interface, space, system, performance or hybrid format.
- Make the audience encounter, respond, assemble, annotate, perform or participate.
- Align the publishing form with Radical Earth: ecology, justice, repair and collective change.

Generated image assets copied from Codex imagegen:
- assets/title-earth-story.png
- assets/earth-states-generated.png

Deck structure follows the user-provided 16-slide outline.
`;
  fs.writeFileSync(path.join(ROOT, "source-notes.txt"), notes);
}

(async () => {
  await writeNotes();
  await renderSlides();
  await renderContactSheet();
  await buildPptx();
  console.log(OUT);
})();
