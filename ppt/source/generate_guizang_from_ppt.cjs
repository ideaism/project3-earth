const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const pptxgen = require("pptxgenjs");
const { chromium } = require("playwright");

const ROOT = path.resolve(__dirname, "..");
const TEMPLATE = "/Users/beijixinfei/.claude/skills/guizang-ppt-skill/assets/template.html";
const OUTPUT_DIR = path.join(ROOT, "output");
const SLIDE_DIR = path.join(ROOT, "slides-guizang");
const QA_DIR = path.join(ROOT, "qa");
const SOURCE_PPTX = path.join(OUTPUT_DIR, "the-earth-written-by-us-earth-story-v2.pptx");
const OUT_HTML = path.join(OUTPUT_DIR, "the-earth-written-by-us-revised_2-editable.html");
const OUT_PPTX = path.join(OUTPUT_DIR, "the-earth-written-by-us-guizang-regenerated.pptx");
const EXTRACTED_MD = path.join(ROOT, "ppt-extracted-content.md");
const WEBSITE = "https://project3-earth.vercel.app/";

for (const dir of [OUTPUT_DIR, SLIDE_DIR, QA_DIR]) fs.mkdirSync(dir, { recursive: true });

const total = 21;
const rel = (p) => path.relative(OUTPUT_DIR, p).replaceAll(path.sep, "/");
const img = {
  title: rel(path.join(ROOT, "assets/title-earth-story.png")),
  earthStates: rel(path.join(ROOT, "assets/earth-states-generated.png")),
  choiceFlow: rel(path.join(ROOT, "..", "outputs/choice-translation/choice-translation-flow.png")),
  turningEarth: rel(path.join(ROOT, "..", "public/assets/earth/earth-turning.png")),
  formSketch: rel(path.join(ROOT, "assets/form-finding-sketch.png")),
  floor: rel(path.join(ROOT, "assets/workshop/workshop-floor-cards.jpg")),
  group: rel(path.join(ROOT, "assets/workshop/group-systems-game.jpg")),
  tipping: rel(path.join(ROOT, "assets/workshop/positive-tipping-report.jpg")),
  feedback: rel(path.join(ROOT, "assets/workshop/responsibility-freedom-loop.jpg")),
  repair: rel(path.join(ROOT, "assets/prototype/03-repair.png")),
  start: rel(path.join(ROOT, "assets/prototype/02-story-start.png")),
  organise: rel(path.join(ROOT, "assets/prototype/07-organise.png")),
  userLanding: rel(path.join(ROOT, "assets/user-testing/user-test-landing-slide.jpg")),
  userInterface: rel(path.join(ROOT, "assets/user-testing/user-test-interface-slide.jpg")),
};

const slidesMeta = [
  ["Publishing Forms: Radical Earth", "The brief is the frame; the project is the answer."],
  ["Brief Interpretation", "The response is an interactive web publication, not a fixed printed page."],
  ["Given Content Analysis", "The assigned text is translated into thresholds, feedback loops, keywords, Earth states and ending cards."],
  ["Core Concept: Reading Changes The Earth", "Reader input changes a living Earth state and story panel."],
  ["World Earth Day Research Evidence", "Workshop feedback loops become the website interaction loop."],
  ["Audience And Encounter Context", "Primary audience: WSA design students and Earth Day workshop participants."],
  ["Example Reading Path", "A keyword such as repair produces story response, variable change and next choices."],
  ["Interface Design", "Left side shows Earth state; right side holds keyword input and ending result."],
  ["Earth States", "Six public-facing states keep feedback legible."],
  ["Keyword And Hidden Variable System", "Words are actions that shift health, biodiversity, care, justice, waste and hope."],
  ["Choice Translation Flow", "Reader-facing options are normalized into keywords, routed through chapters, reflected as Earth states and resolved into endings."],
  ["Chapter Routing And Ending Logic", "Five short chapters interpret the same keyword differently and accumulate toward a scored ending."],
  ["Branching Story Map", "Different words lead toward different futures."],
  ["Possible Endings / Ending Card", "The reader receives a personal publication record."],
  ["User Testing / Iteration Evidence", "Testing led to clearer labels and the change from Collapse Story to Collapse."],
  ["Final Web Prototype", "The deployed website is linked directly in the deck."],
  ["Why This Is Radical Earth", "The publication turns awareness into a quiet rehearsal for futures that might do tomorrow."],
  ["Choosing The Form", "The project rejects a static print zine and chooses a repeatable web publication interface."],
  ["Three Iterations Of The Story Engine", "The keyword-variable-state engine was narrowed across three rounds until the reader path became legible."],
];

function chrome(section, n) {
  return `<div class="chrome"><div>${section}</div><div>${String(n).padStart(2, "0")} / ${total}</div></div>`;
}

function foot(label) {
  return `<div class="foot"><div class="title">The Earth, Written by Us</div><div>${label}</div></div>`;
}

function frameImg(src, classes = "r-16x10", cap = "") {
  return `<figure class="tile"><div class="frame-img ${classes}"><img src="${src}" alt=""></div>${cap ? `<figcaption class="frame-cap"><span class="pf">${cap}</span><span class="idx">source asset</span></figcaption>` : ""}</figure>`;
}

const sections = [
`<section class="slide hero dark" data-animate="hero">
  ${chrome("The Brief · Radical Earth", 1)}
  <div class="frame grid-2-7-5" style="padding-top:5vh; align-items:center">
    <div class="col">
      <div class="kicker" data-anim>interactive web publication</div>
      <h1 class="h-hero" style="font-size:7.4vw" data-anim>Publishing Forms:<br>Radical Earth</h1>
      <p class="lead" style="max-width:46vw" data-anim>The Earth, Written by Us turns a reading assignment into a system where words alter the page.</p>
      <div class="meta-row" data-anim><span>Prototype</span><span>·</span><span>publishing strategy</span><span>·</span><span>story engine</span></div>
      <p class="meta" data-anim><a href="${WEBSITE}" style="color:inherit;text-decoration:none;border-bottom:1px solid currentColor">project3-earth.vercel.app</a></p>
    </div>
    ${frameImg(img.title, "r-4x3", "generated hero image")}
  </div>
  ${foot("opening frame")}
</section>`,
`<section class="slide light">
  ${chrome("Act I · Brief", 2)}
  <div class="frame" style="padding-top:6vh">
    <div class="kicker" data-anim>from static page to active system</div>
    <h2 class="h-xl" data-anim>Brief Interpretation</h2>
    <div class="grid-3" style="margin-top:7vh">
      <div class="stat-card" data-anim><div class="stat-label">brief asks for</div><div class="stat-note" style="font-size:1.25vw">a publishing platform, event, system or hybrid form.</div></div>
      <div class="stat-card" data-anim><div class="stat-label">publication becomes</div><div class="stat-note" style="font-size:1.25vw">interface · system · event · hybrid format.</div></div>
      <div class="stat-card" data-anim><div class="stat-label">my response</div><div class="stat-note" style="font-size:1.25vw">a web publication where each reader action triggers text, variables and an ending.</div></div>
    </div>
  </div>
  ${foot("brief logic")}
</section>`,
`<section class="slide dark">
  ${chrome("Act I · Source", 3)}
  <div class="frame" style="display:grid;grid-template-columns:repeat(3,1fr);padding-top:6vh;gap:3vw;align-items:start">
    <div class="col" data-anim>
      <div class="kicker">source base</div>
      <h2 class="h-md">Original Content</h2>
      <p class="body-zh" style="font-size:.88vw;line-height:1.55;font-weight:700">Radical Earth: Incubator for Utopian Publishing.</p>
      <p class="body-zh" style="font-size:.88vw;line-height:1.55">Winchester School of Art, Department of Design, in conjunction with the Royal Society for Arts and the Future Cities Community Hub. World Earth Day, 21–22 April 2026.</p>
      <p class="body-zh" style="font-size:.88vw;line-height:1.55">The brief frames Earth as a system of thresholds, feedback loops, and shared consequence — and asks how publishing might activate, rather than merely describe, ecological responsibility.</p>
      <div class="callout" style="padding:2vh 1.5vw"><span class="q-big" style="font-size:1.12vw">Tone: urgent, systemic, speculative, but not nihilistic.</span></div>
    </div>
    <div class="col" data-anim>
      <div class="kicker">interpretation</div>
      <h2 class="h-md">My Interpretation</h2>
      <p class="body-zh" style="font-size:.88vw;line-height:1.55">The brief becomes a branching ecological story. Its prompts become reader inputs; its motifs become keywords; climate thresholds become Earth states; accumulated choices become personal ending cards.</p>
      <p class="meta" style="margin-top:1.5vh">motifs</p>
      <p class="body-zh" style="font-size:.88vw;line-height:1.55">threshold · tipping point · heat · soil · river · waste · repair · archive · collective action</p>
    </div>
    <div class="col" data-anim>
      <div class="kicker">editing decision</div>
      <h2 class="h-md">Strategy</h2>
      <p class="body-zh" style="font-size:.88vw;line-height:1.55">Use an intentional edited version. The brief is rewritten as five short interactive chapters; external research supports the system logic.</p>
      <div class="rowline" style="padding:1vh 0"><div class="k" style="font-size:1vw">Core 01</div><div class="v" style="font-size:.78vw">small repeated choices produce systemic consequences.</div><div class="m">choice</div></div>
      <div class="rowline" style="padding:1vh 0"><div class="k" style="font-size:1vw">Core 02</div><div class="v" style="font-size:.78vw">climate futures are unequal; justice shapes who is protected.</div><div class="m">justice</div></div>
      <div class="rowline" style="padding:1vh 0"><div class="k" style="font-size:1vw">Core 03</div><div class="v" style="font-size:.78vw">positive tipping points emerge through repair, sharing and care.</div><div class="m">repair</div></div>
      <div class="rowline" style="padding:1vh 0"><div class="k" style="font-size:1vw">Core 04</div><div class="v" style="font-size:.78vw">publishing can be a participatory interface, not a fixed page.</div><div class="m">form</div></div>
    </div>
  </div>
  ${foot("source pack filled")}
</section>`,
`<section class="slide light">
  ${chrome("Act II · Concept", 4)}
  <div class="frame grid-2-7-5" style="padding-top:6vh">
    <div class="col" style="justify-content:space-between" data-anim="left">
      <div>
        <div class="kicker">core mechanism</div>
        <h2 class="h-xl">Reading changes<br>the Earth</h2>
        <p class="lead" style="margin-top:4vh">The reader types a word. The interface answers through two linked parts: a living Earth state and a story panel.</p>
      </div>
      <div class="callout"><span class="q-big">The page becomes a feedback system.</span><span class="cite">conceptual translation</span></div>
    </div>
    <div class="pipeline-section" data-anim="right">
      <div class="pipeline-label">reader loop</div>
      <div class="pipeline" data-cols="4">
        <div class="step"><div class="step-nb">01</div><div class="step-title">Word</div><div class="step-desc">reader input</div></div>
        <div class="step"><div class="step-nb">02</div><div class="step-title">Variables</div><div class="step-desc">hidden shifts</div></div>
        <div class="step"><div class="step-nb">03</div><div class="step-title">Earth State</div><div class="step-desc">visible feedback</div></div>
        <div class="step"><div class="step-nb">04</div><div class="step-title">Ending</div><div class="step-desc">personal output</div></div>
      </div>
    </div>
  </div>
  ${foot("system concept")}
</section>`,
`<section class="slide hero light" data-animate="hero">
  ${chrome("Research · Earth Day", 5)}
  <div class="frame" style="display:grid; gap:5vh; align-content:center; min-height:80vh">
    <div class="kicker" data-anim>22 April 2026 · West Side Lecture Theatre</div>
    <h1 class="h-hero" style="font-size:7vw" data-anim>World Earth Day<br>research evidence</h1>
    <p class="lead" style="max-width:62vw" data-anim>The workshop made systems thinking physical: room, bodies, cards, timing and feedback became an interface.</p>
  </div>
  ${foot("research frame")}
</section>`,
`<section class="slide light">
  ${chrome("Research · Proof", 6)}
  <div class="frame grid-2-6-6" style="padding-top:6vh">
    <div class="grid-4" data-anim="left">
      ${frameImg(img.floor, "h-22", "floor card participation")}
      ${frameImg(img.group, "h-22", "group systems game")}
      ${frameImg(img.tipping, "h-22", "positive tipping points")}
      ${frameImg(img.feedback, "h-22", "feedback loop")}
    </div>
    <div class="col" data-anim="right">
      <div class="kicker">translation into the prototype</div>
      <h2 class="h-md">Physical feedback becomes web feedback</h2>
      <div class="rowline"><div class="k">Workshop</div><div class="v">people experience rules, misunderstanding and repeated action.</div><div class="m">field</div></div>
      <div class="rowline"><div class="k">Website</div><div class="v">word → feedback → changed Earth state → next choice.</div><div class="m">system</div></div>
    </div>
  </div>
  ${foot("evidence assets")}
</section>`,
`<section class="slide dark">
  ${chrome("Audience · Context", 7)}
  <div class="frame grid-2-6-6" style="padding-top:7vh">
    <div class="col" data-anim="left">
      <div class="kicker">primary audience</div>
      <h2 class="h-xl">WSA students<br>and workshop participants</h2>
      <p class="body-zh">They share the brief context and can test whether the story system makes climate action feel participatory.</p>
    </div>
    <div class="col" data-anim="right">
      <div class="kicker">encounter context</div>
      <div class="rowline"><div class="k">Studio</div><div class="v">type one word, receive a response, watch Earth change.</div><div class="m">exhibition</div></div>
      <div class="rowline"><div class="k">Earth Day</div><div class="v">connect workshop systems thinking to an individual reading path.</div><div class="m">event</div></div>
      <div class="rowline"><div class="k">Online</div><div class="v">save or share an ending card as a personal record.</div><div class="m">publication</div></div>
    </div>
  </div>
  ${foot("audience narrowing")}
</section>`,
`<section class="slide light">
  ${chrome("Prototype · Path", 8)}
  <div class="frame grid-2-7-5" style="padding-top:6vh">
    ${frameImg(img.repair, "r-16x10 fit-contain", "prototype screenshot: repair")}
    <div class="pipeline-section" data-animate="pipeline">
      <div class="pipeline-label">example reading path</div>
      <div class="pipeline" data-cols="3">
        <div class="step" data-anim="step"><div class="step-nb">01</div><div class="step-title">Reader types “repair”</div><div class="step-desc">the input starts the branch.</div></div>
        <div class="step" data-anim="step"><div class="step-nb">02</div><div class="step-title">Story responds</div><div class="step-desc">a crack softens; a tool is shared.</div></div>
        <div class="step" data-anim="step"><div class="step-nb">03</div><div class="step-title">Variables shift</div><div class="step-desc">health, care and hope increase.</div></div>
        <div class="step" data-anim="step"><div class="step-nb">04</div><div class="step-title">Earth changes</div><div class="step-desc">state becomes Repairing Earth.</div></div>
        <div class="step" data-anim="step"><div class="step-nb">05</div><div class="step-title">Next choices open</div><div class="step-desc">listen, plant or share.</div></div>
        <div class="step" data-anim="step"><div class="step-nb">06</div><div class="step-title">Ending path</div><div class="step-desc">can become Regenerative Earth.</div></div>
      </div>
    </div>
  </div>
  ${foot("micro-demo")}
</section>`,
`<section class="slide light">
  ${chrome("Prototype · Interface", 9)}
  <div class="frame grid-2-6-6" style="padding-top:6vh">
    ${frameImg(img.start, "r-16x10 fit-contain", "left: Earth state and variables")}
    ${frameImg(img.organise, "r-16x10 fit-contain", "right: keyword panel and ending")}
  </div>
  ${foot("interface design")}
</section>`,
`<section class="slide hero dark" data-animate="hero">
  ${chrome("Visual System · Earth States", 10)}
  <div class="frame" style="display:grid; gap:5vh; align-content:center; min-height:80vh">
    <div class="kicker" data-anim>the Earth is treated as a living page</div>
    <h1 class="h-hero" style="font-size:8.4vw" data-anim>Six states<br>make feedback legible</h1>
    <p class="lead" style="max-width:58vw" data-anim>Each state carries a different colour, tone and visual signal.</p>
  </div>
  ${foot("visual feedback system")}
</section>`,
`<section class="slide light">
  ${chrome("Visual System · States", 11)}
  <div class="frame" style="padding-top:6vh">
    <div class="kicker" data-anim>public-facing feedback states</div>
    <h2 class="h-xl" data-anim>Earth States</h2>
    <div style="margin-top:5vh" data-anim>${frameImg(img.earthStates, "r-16x9 fit-contain", "fragile · warming · repairing · blooming · collective · collapsing")}</div>
  </div>
  ${foot("state family")}
</section>`,
`<section class="slide dark">
  ${chrome("Engine · Keywords", 12)}
  <div class="frame grid-3" style="padding-top:7vh">
    <div class="stat-card" data-anim><div class="stat-label">positive keywords</div><div class="stat-note">repair · plant · share · listen · teach · reduce · organise</div></div>
    <div class="stat-card" data-anim><div class="stat-label">negative keywords</div><div class="stat-note">extract · consume · ignore · waste · delay · exploit</div></div>
    <div class="stat-card" data-anim><div class="stat-label">hidden variables</div><div class="stat-note">Earth Health · Biodiversity · Community Care · Justice · Waste · Hope</div></div>
  </div>
  ${foot("story engine")}
</section>`,
`<section class="slide light">
  ${chrome("Engine · Translation Flow", 13)}
  <div class="frame grid-2-7-5" style="padding-top:6vh">
    ${frameImg(img.choiceFlow, "r-16x9 fit-contain", "option → keyword → chapter → Earth state → ending")}
    <div class="col" data-anim="right">
      <div class="kicker">uploaded story logic</div>
      <h2 class="h-md">Choice becomes publishable action</h2>
      <div class="rowline"><div class="k">Option</div><div class="v">A visible chip, typed word or branch button starts the reader path.</div><div class="m">reader</div></div>
      <div class="rowline"><div class="k">Keyword</div><div class="v">Aliases normalize into canonical IDs such as repair, delay or technology.</div><div class="m">engine</div></div>
      <div class="rowline"><div class="k">Outcome</div><div class="v">The story response changes variables, Earth state and ending tendency.</div><div class="m">future</div></div>
    </div>
  </div>
  ${foot("choice translation")}
</section>`,
`<section class="slide light">
  ${chrome("Output · Ending Examples", 14)}
  <div class="frame grid-2-6-6" style="padding-top:5vh; gap:3vw">
    <div class="col" data-anim="left" style="gap:2vh">
      <div class="kicker">real ending card example</div>
      <div style="background:rgba(var(--paper-rgb),.72);border:1px solid rgba(var(--ink-rgb),.12);border-left:5px solid #4f6f50;border-radius:8px;padding:2.2vh 2.4vw;box-shadow:0 18px 48px rgba(var(--ink-rgb),.12)">
        <p class="meta" style="color:#a26542;font-weight:700;opacity:.95">ENDING CARD</p>
        <h2 class="h-xl" style="font-size:4.8vw;line-height:.9;margin:1.2vh 0 1.6vh">Regenerative<br>Earth</h2>
        <p class="body-zh" style="font-weight:700;color:#456b4b;opacity:1">Final state: Turning Earth</p>
        <p class="body-zh" style="margin-top:1.4vh;max-width:36vw;font-size:1.05vw;line-height:1.48">The Earth did not become perfect. It became connected again. Every small action found another, until repair became a pattern strong enough to change the weather of the story.</p>
        <div class="callout" style="margin-top:1.3vh;padding:2vh 1.6vw"><span class="q-big" style="font-size:1.35vw">I was not saved by one word. I was changed by what kept returning.</span></div>
        <p class="body-zh" style="font-weight:700;color:#a26542;opacity:1;margin-top:1.3vh;font-size:1vw">Which action could become a pattern in your own community?</p>
        <div class="meta-row" style="margin-top:1.2vh;font-size:.72vw;gap:.8vw"><span>repair</span><span>share</span><span>listen</span><span>organise</span><span>technology</span></div>
      </div>
    </div>
    <div class="col" data-anim="right" style="gap:2vh">
      <div class="kicker">archive record + final state</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:1.4vw">
        <div style="background:rgba(var(--paper-rgb),.76);border:1px solid rgba(var(--ink-rgb),.12);border-radius:8px;padding:2vh 1.5vw;box-shadow:0 16px 38px rgba(var(--ink-rgb),.10)">
          <p class="meta" style="color:#a26542;font-weight:700;opacity:.95">ARCHIVE OF STORIES</p>
          <h3 class="h-md" style="font-size:2.4vw;margin:1.2vh 0">Completed<br>Story Paths</h3>
          <div style="border:1px solid rgba(var(--ink-rgb),.14);border-radius:7px;padding:1.3vh 1vw">
            <p class="body-zh" style="font-weight:700;opacity:1">Regenerative Earth</p>
            <p class="body-zh" style="font-size:1vw;opacity:.72">repair -> share -> listen -> organise -> technology</p>
          </div>
        </div>
        <div style="background:#bfeee7;border:1px solid rgba(var(--ink-rgb),.12);border-radius:8px;padding:2vh 1.5vw;min-height:42vh;display:flex;flex-direction:column;align-items:center;justify-content:center;box-shadow:0 16px 38px rgba(var(--ink-rgb),.10)">
          <img src="${img.turningEarth}" alt="" style="width:56%;height:auto;filter:drop-shadow(0 18px 22px rgba(31,75,54,.25))">
          <div style="background:rgba(var(--paper-rgb),.82);border:1px solid rgba(var(--ink-rgb),.12);border-radius:7px;padding:1.2vh 1vw;margin-top:2vh;width:85%">
            <p class="meta" style="color:#a26542;font-weight:700;opacity:.95">CURRENT STATE</p>
            <p class="body-zh" style="font-weight:700;opacity:1">Turning Earth</p>
            <p class="body-zh" style="font-size:.95vw;opacity:.72">A future published by your repeated choices</p>
          </div>
        </div>
      </div>
      <p class="body-zh" style="opacity:.7">These examples use the real ending title, summary, archive path and final-state language from the web prototype.</p>
    </div>
  </div>
  ${foot("real ending card examples")}
</section>`,
`<section class="slide light">
  ${chrome("Engine · Branching", 15)}
  <div class="frame" style="padding-top:6vh">
    <div class="kicker" data-anim>multiple paths produce multiple futures</div>
    <h2 class="h-xl" data-anim>Branching story map</h2>
    <div class="pipeline-section" style="margin-top:6vh">
      <div class="pipeline-label">keyword → variable → ending</div>
      <div class="pipeline" data-cols="4">
        <div class="step" data-anim><div class="step-nb">repair</div><div class="step-title">Community Care</div><div class="step-desc">regenerative ending</div></div>
        <div class="step" data-anim><div class="step-nb">plant</div><div class="step-title">Biodiversity</div><div class="step-desc">blooming ending</div></div>
        <div class="step" data-anim><div class="step-nb">consume</div><div class="step-title">Waste</div><div class="step-desc">collapse ending</div></div>
        <div class="step" data-anim><div class="step-nb">technology</div><div class="step-title">Efficiency</div><div class="step-desc">unequal green future</div></div>
      </div>
    </div>
  </div>
  ${foot("branch logic")}
</section>`,
`<section class="slide dark">
  ${chrome("Output · Ending Card", 16)}
  <div class="frame grid-3-3" style="padding-top:7vh">
    ${["Regenerative Earth","Fragile Balance","Technological Green Future","Unequal Survival","Collapse"].map((name, i) => `<div class="stat-card" data-anim><div class="stat-label">${String(i + 1).padStart(2, "0")}</div><div class="stat-note" style="font-size:1.35vw">${name}</div></div>`).join("")}
    <div class="callout" data-anim><span class="q-big">The ending card is the reader’s personal publication.</span><span class="cite">output logic</span></div>
  </div>
  ${foot("possible endings")}
</section>`,
`<section class="slide light">
  ${chrome("Proof · Testing", 17)}
  <div class="frame grid-2-6-6" style="padding-top:6vh">
    <div class="col" data-anim="left">
      ${frameImg(img.userLanding, "r-16x10", "user test: landing")}
      <div class="callout"><span class="q-big">Observed issue: interface labels and ending wording felt ambiguous.</span><span class="cite">testing evidence</span></div>
    </div>
    <div class="col" data-anim="right">
      ${frameImg(img.userInterface, "r-16x10", "user test: interface")}
      <div class="callout"><span class="q-big">Iteration decision: rename “Collapse Story” to “Collapse” and make state labels clearer.</span><span class="cite">revision evidence</span></div>
    </div>
  </div>
  ${foot("iteration")}
</section>`,
`<section class="slide hero dark" data-animate="hero">
  ${chrome("Prototype · Live Website", 18)}
  <div class="frame" style="display:grid; gap:5vh; align-content:center; min-height:80vh">
    <div class="kicker" data-anim>deployed publication</div>
    <h1 class="h-hero" style="font-size:7.3vw" data-anim>Open the<br>living story</h1>
    <p class="lead" style="max-width:62vw" data-anim>The final prototype is available as a web publication. Use the link directly from the deck during presentation or assessment review.</p>
    <p class="lead" style="max-width:64vw;font-size:1.65vw" data-anim><a href="${WEBSITE}" style="color:inherit;text-decoration:none;border-bottom:1px solid currentColor">https://project3-earth.vercel.app/</a></p>
  </div>
  ${foot("clickable prototype link")}
</section>`,
`<section class="slide hero light" data-animate="hero">
  ${chrome("Reflection · Radical Earth", 19)}
  <div class="frame" style="display:grid; gap:5vh; align-content:center; min-height:80vh">
    <div class="kicker" data-anim>final argument</div>
    <h1 class="h-hero" style="font-size:7.8vw" data-anim>Small words,<br>big systems</h1>
    <p class="lead" style="max-width:64vw" data-anim>The project treats Earth as a living system shaped by language and collective consequence, not as background decoration.</p>
    <div class="meta-row" data-anim><span>trigger</span><span>·</span><span>action</span><span>·</span><span>output</span><span>·</span><span>archive</span></div>
    <p class="lead" style="max-width:64vw;font-size:1.45vw" data-anim>Live web prototype: <a href="${WEBSITE}" style="color:inherit;text-decoration:none;border-bottom:1px solid currentColor">https://project3-earth.vercel.app/</a></p>
  </div>
  ${foot("closing reflection")}
</section>`,
];

function buildExtractedMarkdown() {
  const lines = [
    "# Extracted PPT Content",
    "",
    `Source PPTX: \`${SOURCE_PPTX}\``,
    "",
    "The source PPTX stores each slide as a full-slide image rather than editable text. The extraction therefore combines PPTX structure verification with the project content map and slide renders.",
    "",
    "## Slide Spine",
    "",
    ...slidesMeta.map(([title, claim], i) => `### ${String(i + 1).padStart(2, "0")} · ${title}\n${claim}\n`),
    "## Added Story Logic",
    "",
    "Source: `/Users/beijixinfei/project3设计/outputs/choice-translation/choice-translation-map.md`",
    "",
    "- Core translation rule: reader option -> normalized keyword -> story response -> variable changes -> Earth state -> ending tendency.",
    "- Chapter route: The Earth Waits -> The First Ripple -> The System Listens -> The Turning Point -> The Story Becomes a Future.",
    "- Ending logic: regenerative, fragile, technological, unequal, or collapse endings are selected from accumulated positive, negative and ambiguous signals.",
    "",
    "## Added Slide 3 Source Pack",
    "",
    "Source: `/Users/beijixinfei/Downloads/slide3_source_pack_the_earth_written_by_us.docx`",
    "",
    "- Slide 3 now uses the provided source-base wording, tone, motifs, interpretation and intentional editing strategy.",
    "- The assigned text title/author/source remains bracketed as `[Assigned text title, author, source]` because the source pack explicitly warns not to invent those details.",
    "",
    `Live prototype: ${WEBSITE}`,
    "",
    "## Regeneration Mode",
    "",
    "Generated with guizang-ppt-skill template: electronic magazine / e-ink style, horizontal navigation, WebGL background, alternating light/dark rhythm, and rendered PPTX backup.",
    "",
  ];
  fs.writeFileSync(EXTRACTED_MD, lines.join("\n"), "utf8");
}

function buildHtml() {
  let html = fs.readFileSync(TEMPLATE, "utf8");
  html = html.replace("[必填] 替换为 PPT 标题 · Deck Title", "The Earth, Written by Us · Radical Earth");
  html = html.replace(
    /--ink:#0a0a0b;\n    --ink-rgb:10,10,11;\n    --paper:#f1efea;\n    --paper-rgb:241,239,234;\n    --paper-tint:#e8e5de;\n    --ink-tint:#18181a;/,
    "--ink:#1a2e1f;\n    --ink-rgb:26,46,31;\n    --paper:#f5f1e8;\n    --paper-rgb:245,241,232;\n    --paper-tint:#ece7da;\n    --ink-tint:#253d2c;"
  );
  html = html.replace("<!-- SLIDES_HERE -->", sections.join("\n\n"));
  fs.writeFileSync(OUT_HTML, html, "utf8");
}

async function renderAndPackage() {
  const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
  const browser = await chromium.launch({
    headless: true,
    executablePath: fs.existsSync(chromePath) ? chromePath : undefined,
  });
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
  await page.goto(pathToFileURL(OUT_HTML).href, { waitUntil: "networkidle" });
  await page.addStyleTag({ content: "#nav,#hint{display:none!important} [data-anim]{opacity:1!important;transform:none!important}" });
  for (let i = 0; i < total; i += 1) {
    await page.evaluate((n) => window.go(n), i);
    await page.waitForTimeout(900);
    await page.screenshot({ path: path.join(SLIDE_DIR, `slide-${String(i + 1).padStart(2, "0")}.png`) });
  }
  await browser.close();

  const pptx = new pptxgen();
  pptx.author = "Codex";
  pptx.company = "MACD";
  pptx.subject = "Regenerated with guizang-ppt-skill from extracted PPT content";
  pptx.title = "The Earth, Written by Us";
  pptx.lang = "en-GB";
  pptx.layout = "LAYOUT_WIDE";
  for (let i = 1; i <= total; i += 1) {
    const slide = pptx.addSlide();
    slide.background = { color: "F5F1E8" };
    slide.addImage({
      path: path.join(SLIDE_DIR, `slide-${String(i).padStart(2, "0")}.png`),
      x: 0,
      y: 0,
      w: 13.333333,
      h: 7.5,
    });
    if (i === 1 || i === 18 || i === 19) {
      slide.addText("https://project3-earth.vercel.app/", {
        x: i === 1 ? 0.48 : 0.92,
        y: i === 1 ? 6.42 : 5.02,
        w: i === 1 ? 3.25 : 4.8,
        h: 0.28,
        fontFace: "Aptos",
        fontSize: 8,
        color: i === 19 ? "1A2E1F" : "F5F1E8",
        transparency: 100,
        hyperlink: { url: WEBSITE, tooltip: "Open live web prototype" },
      });
    }
  }
  await pptx.writeFile({ fileName: OUT_PPTX });
}

(async () => {
  buildExtractedMarkdown();
  buildHtml();
  await renderAndPackage();
  console.log(JSON.stringify({ extracted: EXTRACTED_MD, html: OUT_HTML, pptx: OUT_PPTX, slides: SLIDE_DIR }, null, 2));
})();
