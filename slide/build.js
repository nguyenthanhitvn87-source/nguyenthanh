const pptxgen = require("pptxgenjs");

const RED_LIGHT = "F0323C";   // left panel
const RED_DARK  = "AB0A2E";   // right panel
const CYAN      = "00A3DA";   // left edge stripe
const W = 13.33, H = 7.5;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

const slide = pres.addSlide();
slide.background = { color: RED_DARK };

// --- panels -----------------------------------------------------------
slide.addShape(pres.ShapeType.rect, {
  x: 0, y: 0, w: 0.42, h: H, fill: { color: CYAN }, line: { color: CYAN },
});
slide.addShape(pres.ShapeType.rect, {
  x: 0.42, y: 0, w: 4.78, h: H, fill: { color: RED_LIGHT }, line: { color: RED_LIGHT },
});
slide.addShape(pres.ShapeType.rect, {
  x: 5.20, y: 0, w: W - 5.20, h: H, fill: { color: RED_DARK }, line: { color: RED_DARK },
});

// --- DKSH wordmark (top right) ---------------------------------------
slide.addText("DKSH", {
  x: 9.9, y: 0.28, w: 2.9, h: 0.6, align: "right", margin: 0,
  fontFace: "Arial", fontSize: 34, bold: true, color: "FFFFFF", charSpacing: 1,
});

// --- icon: lightning bolt inside a ring of dots -----------------------
// centred a little above mid-panel so the icon + SAP read as one balanced group
const cx = 2.81, cy = 2.75;            // icon centre, inches

slide.addShape(pres.ShapeType.lightningBolt, {
  x: cx - 0.62, y: cy - 0.78, w: 1.24, h: 1.56,
  fill: { color: "FFFFFF" }, line: { color: "FFFFFF" },
});

// eight dots forming a rounded frame around the bolt
const dotR = 0.075;
const dots = [
  [-1.05, -0.95], [0.00, -0.95], [1.05, -0.95],
  [-1.05,  0.00],                [1.05,  0.00],
  [-1.05,  0.95], [0.00,  0.95], [1.05,  0.95],
];
dots.forEach(([dx, dy]) => {
  slide.addShape(pres.ShapeType.ellipse, {
    x: cx + dx - dotR, y: cy + dy - dotR, w: dotR * 2, h: dotR * 2,
    fill: { color: "FFFFFF" }, line: { color: "FFFFFF" },
  });
});

// --- SAP, directly under the icon ------------------------------------
slide.addText("SAP", {
  x: cx - 1.6, y: cy + 1.32, w: 3.2, h: 0.75, align: "center", margin: 0,
  fontFace: "Arial", fontSize: 40, bold: true, color: "FFFFFF", charSpacing: 3,
});

// --- title ------------------------------------------------------------
slide.addText("DEFAULT EXCLUDE 4PL FOR\nNEW CUSTOMER L' OREAL", {
  x: 5.9, y: 1.30, w: 6.9, h: 1.5, margin: 0,
  fontFace: "Arial", fontSize: 34, bold: true, color: "FFFFFF", lineSpacing: 40,
});

// --- owners -----------------------------------------------------------
slide.addText(
  [
    { text: "Business Owner: Luong Ngoc Truong, HEC Sales Manager", options: { breakLine: true } },
    { text: "Process Owner: Tran Quang Thuan, IT CSSC Vietnam", options: { breakLine: true } },
    { text: "Technical Owner: Nguyen Vu Thanh, IT CSSC Vietnam" },
  ],
  {
    x: 5.9, y: 3.45, w: 6.9, h: 1.1, margin: 0,
    fontFace: "Arial", fontSize: 16, color: "FFFFFF", lineSpacing: 24,
  }
);

slide.addText("HCMC, 10 Aug 2026", {
  x: 5.9, y: 4.80, w: 6.9, h: 0.4, margin: 0,
  fontFace: "Arial", fontSize: 16, color: "FFFFFF",
});

// --- tagline ----------------------------------------------------------
slide.addText("Delivering Growth – in Asia and Beyond.", {
  x: 5.9, y: 6.78, w: 6.9, h: 0.35, margin: 0,
  fontFace: "Arial", fontSize: 13, color: "FFFFFF",
});

pres.writeFile({ fileName: process.argv[2] || "slide.pptx" });
