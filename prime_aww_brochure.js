const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "OpenAI";
pptx.company = "Prime Property Maintenance";
pptx.subject = "Prime Maintenance × AWM Alliance brochure";
pptx.title = "Prime Maintenance × AWM Alliance";
pptx.lang = "en-CA";
pptx.theme = {
  headFontFace: "Aptos",
  bodyFontFace: "Aptos",
  lang: "en-CA",
};

const C = {
  NAVY: "0F2233",
  NAVY_2: "17364D",
  TEAL: "3BB7B2",
  TEAL_SOFT: "DDF5F3",
  WHITE: "FFFFFF",
  GREY_1: "F5F7FA",
  GREY_2: "D9E2EC",
  GREY_3: "6B7C8F",
  TEXT: "22313F",
  TEXT_SOFT: "566575",
};

const ASSETS = {
  cover: path.join(__dirname, "prime back contrast logo.jpg"),
  interior: path.join(__dirname, "prime clean 4.jpeg"),
  exterior1: path.join(__dirname, "prime clean 3.jpeg"),
  exterior2: path.join(__dirname, "prime clean 2.jpeg"),
  windowView: path.join(__dirname, "prime clean 1.jpeg"),
  landscaping: path.join(__dirname, "prime photo 1.jpg"),
};

function ensureFile(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Missing asset: ${filePath}`);
  }
  return filePath;
}

Object.values(ASSETS).forEach(ensureFile);

function addFullBleedImage(slide, imgPath, x = 0, y = 0, w = 13.333, h = 7.5) {
  slide.addImage({
    path: imgPath,
    x,
    y,
    w,
    h,
  });
}

function addTopAccent(slide) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.06,
    line: { color: C.TEAL, transparency: 100 },
    fill: { color: C.TEAL },
  });
}

function addFooter(slide, pageNum) {
  slide.addText(`${pageNum}`, {
    x: 12.65,
    y: 7.05,
    w: 0.25,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 9,
    color: C.GREY_3,
    align: "right",
    margin: 0,
  });
}

function addSectionEyebrow(slide, text, x, y, w) {
  slide.addText(text.toUpperCase(), {
    x,
    y,
    w,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 8,
    bold: true,
    color: C.TEAL,
    charSpace: 2.5,
    margin: 0,
  });
}

function addBulletList(slide, items, x, y, w, fontSize = 12, color = C.TEXT, gap = 0.44) {
  items.forEach((item, i) => {
    const yy = y + i * gap;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: yy + 0.08,
      w: 0.09,
      h: 0.09,
      line: { color: C.TEAL, transparency: 100 },
      fill: { color: C.TEAL },
    });
    slide.addText(item, {
      x: x + 0.18,
      y: yy,
      w: w - 0.18,
      h: 0.28,
      fontFace: "Aptos",
      fontSize,
      color,
      margin: 0,
      breakLine: false,
    });
  });
}

function addCard(slide, opts) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rectRadius: 0.06,
    line: { color: opts.lineColor || C.GREY_2, width: 1 },
    fill: { color: opts.fillColor || C.WHITE },
  });

  if (opts.accent !== false) {
    slide.addShape(pptx.ShapeType.rect, {
      x: opts.x,
      y: opts.y,
      w: opts.w,
      h: 0.06,
      line: { color: C.TEAL, transparency: 100 },
      fill: { color: C.TEAL },
    });
  }
}

// PAGE 1 — COVER
{
  const s = pptx.addSlide();
  addFullBleedImage(s, ASSETS.cover);
  addTopAccent(s);

  s.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    line: { color: C.NAVY, transparency: 100 },
    fill: { color: C.NAVY, transparency: 36 },
  });

  s.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 6.8,
    h: 7.5,
    line: { color: C.NAVY, transparency: 100 },
    fill: { color: C.NAVY, transparency: 18 },
  });

  s.addText("PRIME MAINTENANCE × AWM ALLIANCE", {
    x: 0.7,
    y: 1.0,
    w: 5.3,
    h: 0.4,
    fontFace: "Aptos",
    fontSize: 10,
    bold: true,
    color: C.TEAL,
    charSpace: 2.2,
    margin: 0,
  });

  s.addText("A Strategic Property\nServices Partnership", {
    x: 0.7,
    y: 1.55,
    w: 5.6,
    h: 1.5,
    fontFace: "Aptos",
    fontSize: 27,
    bold: true,
    color: C.WHITE,
    margin: 0,
    valign: "mid",
    breakLine: false,
  });

  s.addShape(pptx.ShapeType.rect, {
    x: 0.72,
    y: 3.35,
    w: 1.15,
    h: 0.06,
    line: { color: C.TEAL, transparency: 100 },
    fill: { color: C.TEAL },
  });

  s.addText(
    "Integrated building operations designed to enhance asset value, elevate tenant experience, and deliver consistent execution across modern residential and mixed-use properties.",
    {
      x: 0.7,
      y: 3.62,
      w: 5.2,
      h: 1.15,
      fontFace: "Aptos",
      fontSize: 13,
      color: "E8EEF5",
      margin: 0,
      valign: "mid",
    }
  );

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 5.45,
    w: 3.8,
    h: 0.72,
    rectRadius: 0.05,
    line: { color: C.TEAL, width: 1 },
    fill: { color: C.TEAL, transparency: 82 },
  });

  s.addText("Executive Brochure Presentation", {
    x: 1.0,
    y: 5.68,
    w: 3.2,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 11,
    bold: true,
    color: C.WHITE,
    align: "center",
    margin: 0,
  });

  addFooter(s, 1);
}

// PAGE 2 — WHY PRIME
{
  const s = pptx.addSlide();
  s.background = { color: C.WHITE };
  addTopAccent(s);

  addSectionEyebrow(s, "More Than a Vendor", 0.7, 0.6, 2.8);

  s.addText("Built to Operate Like an Extension of Your Team.", {
    x: 0.7,
    y: 0.95,
    w: 6.0,
    h: 0.65,
    fontFace: "Aptos",
    fontSize: 24,
    bold: true,
    color: C.NAVY,
    margin: 0,
  });

  s.addText(
    "Prime Property Maintenance brings janitorial, building management, concierge support, and exterior care into one accountable operating model. Instead of coordinating multiple disconnected vendors, AWM gains a responsive partner focused on consistent standards, proactive oversight, and portfolio-ready execution.",
    {
      x: 0.7,
      y: 1.8,
      w: 5.6,
      h: 1.3,
      fontFace: "Aptos",
      fontSize: 12.5,
      color: C.TEXT_SOFT,
      margin: 0,
    }
  );

  s.addImage({
    path: ASSETS.interior,
    x: 7.55,
    y: 0.65,
    w: 5.05,
    h: 2.85,
  });

  s.addShape(pptx.ShapeType.rect, {
    x: 7.55,
    y: 0.65,
    w: 5.05,
    h: 2.85,
    line: { color: C.NAVY, transparency: 100 },
    fill: { color: C.NAVY, transparency: 72 },
  });

  s.addText("Operational consistency,\nvisible execution.", {
    x: 7.92,
    y: 2.35,
    w: 3.8,
    h: 0.7,
    fontFace: "Aptos",
    fontSize: 20,
    bold: true,
    color: C.WHITE,
    margin: 0,
  });

  const cards = [
    {
      title: "Integrated Operations",
      body: "Janitorial, building management, concierge, and exterior support delivered under one connected service model.",
      x: 0.7,
      y: 3.6,
    },
    {
      title: "Single Accountability",
      body: "One point of contact, one escalation path, and one standard across the property.",
      x: 4.53,
      y: 3.6,
    },
    {
      title: "Systems-Driven Execution",
      body: "Checkpoint tracking, documentation, and clear reporting reduce management friction and improve visibility.",
      x: 8.36,
      y: 3.6,
    },
  ];

  cards.forEach((card) => {
    addCard(s, {
      x: card.x,
      y: card.y,
      w: 3.45,
      h: 1.75,
      fillColor: C.GREY_1,
      lineColor: C.GREY_2,
    });

    s.addText(card.title, {
      x: card.x + 0.18,
      y: card.y + 0.24,
      w: 2.9,
      h: 0.3,
      fontFace: "Aptos",
      fontSize: 15,
      bold: true,
      color: C.NAVY,
      margin: 0,
    });

    s.addText(card.body, {
      x: card.x + 0.18,
      y: card.y + 0.68,
      w: 3.0,
      h: 0.72,
      fontFace: "Aptos",
      fontSize: 10.5,
      color: C.TEXT_SOFT,
      margin: 0,
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 5.85,
    w: 11.9,
    h: 0.8,
    rectRadius: 0.04,
    line: { color: C.TEAL, width: 1 },
    fill: { color: C.TEAL_SOFT },
  });

  s.addText(
    "Reduced management burden. Greater consistency across sites. Stronger tenant experience. Better long-term operating control.",
    {
      x: 1.0,
      y: 6.08,
      w: 11.3,
      h: 0.2,
      fontFace: "Aptos",
      fontSize: 12,
      bold: true,
      color: C.NAVY,
      align: "center",
      margin: 0,
    }
  );

  addFooter(s, 2);
}

// PAGE 3 — SERVICES
{
  const s = pptx.addSlide();
  s.background = { color: C.GREY_1 };
  addTopAccent(s);

  addSectionEyebrow(s, "Core Services", 0.7, 0.55, 2.4);

  s.addText("Three Integrated Service Pillars.\nOne Accountable Partner.", {
    x: 0.7,
    y: 0.9,
    w: 6.4,
    h: 0.95,
    fontFace: "Aptos",
    fontSize: 24,
    bold: true,
    color: C.NAVY,
    margin: 0,
  });

  s.addImage({ path: ASSETS.exterior1, x: 8.1, y: 0.7, w: 2.0, h: 1.55 });
  s.addImage({ path: ASSETS.exterior2, x: 10.25, y: 0.7, w: 2.0, h: 1.55 });
  s.addImage({ path: ASSETS.landscaping, x: 8.1, y: 2.4, w: 4.15, h: 1.85 });

  const serviceCards = [
    {
      title: "Janitorial Services",
      subtitle: "Consistency",
      body: [
        "Common-area cleaning and sanitization",
        "Lobby, amenity, stairwell, and parkade upkeep",
        "Floor care, garbage handling, and routine detailing",
      ],
      x: 0.7,
      y: 2.15,
    },
    {
      title: "Building Management",
      subtitle: "Efficiency",
      body: [
        "Routine inspections and operational oversight",
        "Vendor coordination and issue escalation",
        "Preventative maintenance support and reporting",
      ],
      x: 4.55,
      y: 2.15,
    },
    {
      title: "Concierge Support",
      subtitle: "Experience",
      body: [
        "Resident-facing professionalism at the front line",
        "Visitor, parcel, and access coordination",
        "Amenity and day-to-day service support",
      ],
      x: 8.4,
      y: 4.6,
    },
  ];

  serviceCards.forEach((card, idx) => {
    const w = idx < 2 ? 3.4 : 3.85;
    const h = idx < 2 ? 3.2 : 1.75;

    addCard(s, {
      x: card.x,
      y: card.y,
      w,
      h,
      fillColor: C.WHITE,
      lineColor: C.GREY_2,
    });

    s.addText(card.title, {
      x: card.x + 0.18,
      y: card.y + 0.24,
      w: w - 0.36,
      h: 0.28,
      fontFace: "Aptos",
      fontSize: 16,
      bold: true,
      color: C.NAVY,
      margin: 0,
    });

    s.addText(card.subtitle, {
      x: card.x + 0.18,
      y: card.y + 0.58,
      w: 1.5,
      h: 0.2,
      fontFace: "Aptos",
      fontSize: 9,
      bold: true,
      color: C.TEAL,
      margin: 0,
    });

    addBulletList(s, card.body, card.x + 0.18, card.y + 0.95, w - 0.36, 10.5, C.TEXT_SOFT, 0.47);
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 5.95,
    w: 7.25,
    h: 0.85,
    rectRadius: 0.04,
    line: { color: C.GREY_2, width: 1 },
    fill: { color: C.WHITE },
  });

  s.addText("Supporting Services", {
    x: 0.95,
    y: 6.15,
    w: 1.8,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 10,
    bold: true,
    color: C.TEAL,
    margin: 0,
  });

  s.addText(
    "Pressure washing • Landscaping • Snow removal • Parking area maintenance • Carpet care • Handyman support • Garbage & recycling • Painting • Locksmith services",
    {
      x: 2.2,
      y: 6.13,
      w: 5.3,
      h: 0.22,
      fontFace: "Aptos",
      fontSize: 9.5,
      color: C.TEXT_SOFT,
      margin: 0,
    }
  );

  addFooter(s, 3);
}

// PAGE 4 — CLOSE / CONTACT
{
  const s = pptx.addSlide();
  s.background = { color: C.WHITE };
  addTopAccent(s);

  s.addImage({
    path: ASSETS.windowView,
    x: 0,
    y: 0,
    w: 4.9,
    h: 7.5,
  });

  s.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 4.9,
    h: 7.5,
    line: { color: C.NAVY, transparency: 100 },
    fill: { color: C.NAVY, transparency: 64 },
  });

  s.addText("Prime is ready to\nassess, deploy,\nand deliver.", {
    x: 0.5,
    y: 1.15,
    w: 3.7,
    h: 1.35,
    fontFace: "Aptos",
    fontSize: 25,
    bold: true,
    color: C.WHITE,
    margin: 0,
  });

  s.addText("Low-friction pilot.\nClear standards.\nPortfolio-ready scale.", {
    x: 0.5,
    y: 3.15,
    w: 3.2,
    h: 1.05,
    fontFace: "Aptos",
    fontSize: 14,
    color: "E8EEF5",
    margin: 0,
  });

  addSectionEyebrow(s, "Next Steps", 5.45, 0.7, 2.0);

  s.addText("Let’s Build This Together.", {
    x: 5.45,
    y: 1.05,
    w: 5.8,
    h: 0.5,
    fontFace: "Aptos",
    fontSize: 24,
    bold: true,
    color: C.NAVY,
    margin: 0,
  });

  s.addText(
    "Prime recommends starting with a focused property walkthrough, aligning on service scope and KPIs, and launching a pilot that can scale across the AWM portfolio with confidence.",
    {
      x: 5.45,
      y: 1.75,
      w: 6.6,
      h: 0.9,
      fontFace: "Aptos",
      fontSize: 12.5,
      color: C.TEXT_SOFT,
      margin: 0,
    }
  );

  const steps = [
    "Schedule site walkthrough",
    "Confirm pilot scope",
    "Align on KPIs",
    "Launch within 30 days",
  ];

  steps.forEach((step, i) => {
    const y = 2.9 + i * 0.62;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 5.45,
      y,
      w: 5.7,
      h: 0.42,
      rectRadius: 0.04,
      line: { color: C.GREY_2, width: 1 },
      fill: { color: i === 3 ? C.TEAL_SOFT : C.GREY_1 },
    });

    s.addText(`${i + 1}`, {
      x: 5.68,
      y: y + 0.1,
      w: 0.25,
      h: 0.15,
      fontFace: "Aptos",
      fontSize: 10,
      bold: true,
      color: C.TEAL,
      margin: 0,
      align: "center",
    });

    s.addText(step, {
      x: 6.05,
      y: y + 0.09,
      w: 4.6,
      h: 0.16,
      fontFace: "Aptos",
      fontSize: 11.5,
      bold: i === 3,
      color: C.NAVY,
      margin: 0,
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 5.45,
    y: 5.7,
    w: 6.55,
    h: 1.2,
    rectRadius: 0.05,
    line: { color: C.TEAL, width: 1 },
    fill: { color: C.NAVY },
  });

  s.addText("Contact", {
    x: 5.72,
    y: 5.95,
    w: 1.0,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 10,
    bold: true,
    color: C.TEAL,
    margin: 0,
  });

  s.addText("778-952-0882   |   c.chan@prime-maintenance.ca   |   prime-maintenance.ca", {
    x: 5.72,
    y: 6.25,
    w: 5.95,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 11,
    color: C.WHITE,
    margin: 0,
  });

  s.addText("WorkSafeBC aligned • $5M liability insurance • WHMIS trained • First Aid trained", {
    x: 5.45,
    y: 7.0,
    w: 6.6,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 8.5,
    color: C.GREY_3,
    margin: 0,
  });

  addFooter(s, 4);
}

pptx.writeFile({ fileName: "Prime_Maintenance_X_AWM_Alliance_Brochure.pptx" });
