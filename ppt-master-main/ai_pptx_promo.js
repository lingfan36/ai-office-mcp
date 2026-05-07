const PptxGenJS = require("C:/Users/zack_/AppData/Roaming/npm/node_modules/pptxgenjs");

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";
pres.title = "AI生成PPTX 产品宣传介绍";
pres.author = "PPT Master";

// ── Color palette ──────────────────────────────────────────
const C = {
  plum:    "160C28",
  purple:  "6D28D9",
  purpleL: "A855F7",
  rose:    "EC4899",
  indigo:  "4F46E5",
  cyan:    "06B6D4",
  green:   "10B981",
  amber:   "F59E0B",
  white:   "FFFFFF",
  lavender:"F5F3FF",
  textDark:"1E293B",
  textMid: "475569",
  textMute:"94A3B8",
  border:  "E8E3F8",
};

const mk = () => ({ type: "outer", color: "000000", blur: 10, offset: 3, angle: 135, opacity: 0.10 });

// ─────────────────────────────────────────────────────────────
// SLIDE 1 — Cover (split-screen)
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  // Left dark panel — 45% width
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 4.5, h: 5.625,
    fill: { color: C.plum }, line: { color: C.plum },
  });

  // Purple circle decoration (left panel)
  s.addShape(pres.shapes.OVAL, {
    x: -0.8, y: 3.5, w: 3.5, h: 3.5,
    fill: { color: C.purple, transparency: 80 },
    line: { color: C.purple, transparency: 70, width: 1 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 2.8, y: -0.8, w: 2.8, h: 2.8,
    fill: { color: C.purpleL, transparency: 85 },
    line: { color: C.purpleL, transparency: 75, width: 1 },
  });

  // Left: label
  s.addText("PRODUCT", {
    x: 0.42, y: 0.72, w: 3.6, h: 0.3,
    fontSize: 9.5, bold: true, color: C.purpleL,
    charSpacing: 5, fontFace: "Calibri", margin: 0,
  });
  s.addText("OVERVIEW", {
    x: 0.42, y: 0.98, w: 3.6, h: 0.3,
    fontSize: 9.5, bold: true, color: C.rose,
    charSpacing: 5, fontFace: "Calibri", margin: 0,
  });

  // Left: main brand
  s.addText("PPT\nMaster", {
    x: 0.42, y: 1.42, w: 3.7, h: 2.2,
    fontSize: 56, bold: true, color: C.white,
    fontFace: "Microsoft YaHei", margin: 0, valign: "top",
  });

  // Left: tagline
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.42, y: 3.75, w: 2.0, h: 0.04,
    fill: { color: C.purpleL }, line: { color: C.purpleL },
  });
  s.addText("AI 演示文稿生成引擎", {
    x: 0.42, y: 3.9, w: 3.7, h: 0.38,
    fontSize: 13, color: "C4B5FD",
    fontFace: "Microsoft YaHei", margin: 0,
  });

  // Left: version badge
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.42, y: 4.5, w: 1.1, h: 0.32,
    fill: { color: C.purple }, line: { color: C.purple },
  });
  s.addText("2026", {
    x: 0.42, y: 4.5, w: 1.1, h: 0.32,
    fontSize: 11, bold: true, color: C.white,
    fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });

  // Right panel: headline
  s.addText("从内容到精美\n演示文稿", {
    x: 4.85, y: 0.5, w: 4.9, h: 2.0,
    fontSize: 36, bold: true, color: C.textDark,
    fontFace: "Microsoft YaHei", margin: 0, valign: "top",
  });

  // Right: accent underline on first line only
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.85, y: 2.54, w: 3.8, h: 0.06,
    fill: { color: C.rose }, line: { color: C.rose },
  });

  s.addText("一键完成", {
    x: 4.85, y: 2.7, w: 4.9, h: 0.55,
    fontSize: 30, bold: true, color: C.purple,
    fontFace: "Microsoft YaHei", margin: 0,
  });

  s.addText("多角色 AI 协作 · 原生 DrawingML · 支持 PDF / DOCX / URL 输入", {
    x: 4.85, y: 3.42, w: 4.9, h: 0.5,
    fontSize: 12, color: C.textMute,
    fontFace: "Microsoft YaHei", wrap: true, margin: 0,
  });

  // Right: 3 feature pills
  const pills = [
    { t: "⚡ AI 三角色", c: C.purple },
    { t: "📄 多格式输入", c: C.indigo },
    { t: "✨ 可编辑输出", c: C.rose },
  ];
  pills.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.85 + i * 1.66, y: 4.12, w: 1.52, h: 0.38,
      fill: { color: C.lavender }, line: { color: p.c, width: 1 },
    });
    s.addText(p.t, {
      x: 4.85 + i * 1.66, y: 4.12, w: 1.52, h: 0.38,
      fontSize: 10, color: p.c, fontFace: "Microsoft YaHei",
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Bottom footer strip
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.5, y: 5.2, w: 5.5, h: 0.425,
    fill: { color: "F1EEF9" }, line: { color: C.border },
  });
  s.addText("PPT Master  ·  AI-Powered Presentation Engine", {
    x: 4.7, y: 5.22, w: 5.1, h: 0.38,
    fontSize: 9.5, color: C.textMute, fontFace: "Calibri", margin: 0,
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 2 — 三大痛点（水平条形风格）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lavender };

  // Title area
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: C.plum }, line: { color: C.plum },
  });
  s.addText("传统 PPT 制作的三大痛点", {
    x: 0.55, y: 0.1, w: 8, h: 0.82,
    fontSize: 26, bold: true, color: C.white,
    fontFace: "Microsoft YaHei", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 9.4, y: 0.3, w: 0.55, h: 0.45,
    fill: { color: C.rose }, line: { color: C.rose },
  });
  s.addText("×3", {
    x: 9.4, y: 0.3, w: 0.55, h: 0.45,
    fontSize: 14, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });

  const pain = [
    {
      num: "01",
      icon: "⏱",
      title: "耗时费力",
      desc: "排版、配色、图文对齐往往耗费数小时甚至整天，严重挤占核心工作时间",
      stat: "6h+",
      statLabel: "平均制作时间",
      color: "EF4444",
      bg: "FEF2F2",
    },
    {
      num: "02",
      icon: "🎨",
      title: "设计门槛高",
      desc: "非设计背景人员难以兼顾内容与美观，专业级视觉表达长期是业务痛点",
      stat: "73%",
      statLabel: "用户感到困难",
      color: "F59E0B",
      bg: "FFFBEB",
    },
    {
      num: "03",
      icon: "🔄",
      title: "修改成本大",
      desc: "内容变更牵一发动全身，逐页调整让协作效率极低，频繁返工耗尽精力",
      stat: "40%",
      statLabel: "时间用于修改",
      color: C.purple,
      bg: "F5F3FF",
    },
  ];

  pain.forEach((p, i) => {
    const y = 1.22 + i * 1.36;
    const h = 1.2;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h,
      fill: { color: p.bg }, line: { color: p.color, width: 1 },
      shadow: mk(),
    });

    // Left number
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.7, h,
      fill: { color: p.color }, line: { color: p.color },
    });
    s.addText(p.num, {
      x: 0.4, y, w: 0.7, h,
      fontSize: 15, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0, fontFace: "Calibri",
    });

    // Icon
    s.addText(p.icon, {
      x: 1.22, y: y + 0.28, w: 0.55, h: 0.55,
      fontSize: 24, align: "center", valign: "middle", margin: 0,
    });

    // Title
    s.addText(p.title, {
      x: 1.86, y: y + 0.08, w: 3.8, h: 0.44,
      fontSize: 17, bold: true, color: C.textDark,
      fontFace: "Microsoft YaHei", margin: 0,
    });

    // Desc
    s.addText(p.desc, {
      x: 1.86, y: y + 0.54, w: 4.8, h: 0.56,
      fontSize: 12.5, color: C.textMid,
      fontFace: "Microsoft YaHei", wrap: true,
    });

    // Right stat
    s.addShape(pres.shapes.RECTANGLE, {
      x: 7.8, y: y + 0.16, w: 1.6, h: 0.88,
      fill: { color: p.color, transparency: 88 },
      line: { color: p.color, transparency: 70, width: 1 },
    });
    s.addText(p.stat, {
      x: 7.8, y: y + 0.16, w: 1.6, h: 0.52,
      fontSize: 24, bold: true, color: p.color,
      fontFace: "Calibri", align: "center", margin: 0,
    });
    s.addText(p.statLabel, {
      x: 7.8, y: y + 0.68, w: 1.6, h: 0.3,
      fontSize: 9.5, color: p.color,
      fontFace: "Microsoft YaHei", align: "center",
    });
  });

  s.addText("PPT Master 一站式解决以上所有问题 →", {
    x: 0, y: 5.2, w: 10, h: 0.38,
    fontSize: 13, bold: true, color: C.purple,
    fontFace: "Microsoft YaHei", align: "center", margin: 0,
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 3 — 产品介绍（深色全版）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.plum };

  // Corner decoration
  s.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -1.5, w: 5.5, h: 5.5,
    fill: { color: C.purple, transparency: 88 },
    line: { color: C.purple, transparency: 80, width: 1 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: -1.2, y: 3.8, w: 4, h: 4,
    fill: { color: C.indigo, transparency: 90 },
    line: { color: C.indigo, transparency: 85, width: 1 },
  });

  s.addText("WHAT IS", {
    x: 0.6, y: 0.35, w: 8, h: 0.32,
    fontSize: 10, bold: true, color: C.purpleL,
    charSpacing: 5, fontFace: "Calibri", margin: 0,
  });
  s.addText("PPT Master", {
    x: 0.6, y: 0.65, w: 9, h: 1.0,
    fontSize: 48, bold: true, color: C.white,
    fontFace: "Microsoft YaHei", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.72, w: 1.8, h: 0.05,
    fill: { color: C.rose }, line: { color: C.rose },
  });

  s.addText(
    "PPT Master 是一款 AI 驱动的演示文稿自动生成引擎。它通过「策略师 → 图像生成师 → 执行师」三角色大模型协作流水线，将 PDF、DOCX、网页、Markdown 等任意格式的源内容，自动转化为专业级、原生可编辑的 PPTX 演示文稿。",
    {
      x: 0.6, y: 1.88, w: 8.8, h: 1.3,
      fontSize: 14.5, color: "C4B5FD",
      fontFace: "Microsoft YaHei", wrap: true, valign: "top",
    }
  );

  // Feature grid 2x2
  const feats = [
    { icon: "🧠", title: "智能内容理解", desc: "深度解析源文档，提炼结构化要点", color: C.purple },
    { icon: "🎨", title: "专业视觉设计", desc: "AI 自动配色、布局，媲美设计师", color: C.rose },
    { icon: "⚡", title: "高速批量生成", desc: "分钟级完成十余页专业 PPT", color: C.cyan },
    { icon: "✏️", title: "原生可编辑", desc: "DrawingML 格式，PowerPoint 完全兼容", color: C.amber },
  ];

  feats.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * 4.7;
    const y = 3.35 + row * 1.0;

    s.addShape(pres.shapes.OVAL, {
      x: x, y: y + 0.08, w: 0.54, h: 0.54,
      fill: { color: f.color, transparency: 30 }, line: { color: f.color, transparency: 20 },
    });
    s.addText(f.icon, {
      x: x, y: y + 0.08, w: 0.54, h: 0.54,
      fontSize: 16, align: "center", valign: "middle", margin: 0,
    });
    s.addText(f.title, {
      x: x + 0.66, y: y + 0.05, w: 3.7, h: 0.32,
      fontSize: 14, bold: true, color: C.white,
      fontFace: "Microsoft YaHei", margin: 0,
    });
    s.addText(f.desc, {
      x: x + 0.66, y: y + 0.38, w: 3.7, h: 0.32,
      fontSize: 12, color: "A78BFA",
      fontFace: "Microsoft YaHei", margin: 0,
    });
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 4 — 核心功能（左图标 + 右详情 交替列表）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  // Left sidebar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.22, h: 5.625,
    fill: { color: C.purple }, line: { color: C.purple },
  });

  s.addText("核心功能", {
    x: 0.45, y: 0.22, w: 5, h: 0.62,
    fontSize: 30, bold: true, color: C.textDark,
    fontFace: "Microsoft YaHei", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 0.84, w: 3.2, h: 0.05,
    fill: { color: C.purple }, line: { color: C.purple },
  });
  s.addText("四大核心能力，重新定义演示制作效率", {
    x: 0.45, y: 0.98, w: 7, h: 0.36,
    fontSize: 13.5, color: C.textMute,
    fontFace: "Microsoft YaHei", margin: 0,
  });

  const feats = [
    {
      icon: "🧠", num: "01",
      title: "智能内容理解与提炼",
      desc: "支持 PDF、DOCX、网页 URL、Markdown 等多格式输入，AI 自动解析文档结构，提炼核心论点与关键数据，无需手动整理",
      color: C.purple,
    },
    {
      icon: "🎨", num: "02",
      title: "专业视觉方案生成",
      desc: "策略师 AI 根据内容主题自动制定配色方案、版式结构与字体搭配，视觉效果媲美专业设计师出品",
      color: C.rose,
    },
    {
      icon: "⚡", num: "03",
      title: "全程自动化流水线",
      desc: "从内容分析到 PPT 输出全程无需人工干预，分钟级完成十余页演示文稿，效率提升超 10 倍",
      color: C.cyan,
    },
    {
      icon: "✏️", num: "04",
      title: "原生可编辑 PPTX 输出",
      desc: "输出标准 DrawingML 格式，每个文字、形状均为真实 PowerPoint 对象，完全保留二次编辑能力",
      color: C.amber,
    },
  ];

  feats.forEach((f, i) => {
    const y = 1.5 + i * 1.02;
    const hCard = 0.88;

    // Card bg (alternating)
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.45, y, w: 9.3, h: hCard,
      fill: { color: i % 2 === 0 ? C.lavender : C.white },
      line: { color: C.border, width: 1 },
      shadow: mk(),
    });

    // Color badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.45, y, w: 0.06, h: hCard,
      fill: { color: f.color }, line: { color: f.color },
    });

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.62, y: y + 0.14, w: 0.6, h: 0.6,
      fill: { color: f.color }, line: { color: f.color },
    });
    s.addText(f.num, {
      x: 0.62, y: y + 0.14, w: 0.6, h: 0.6,
      fontSize: 12, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0, fontFace: "Calibri",
    });

    // Icon
    s.addText(f.icon, {
      x: 1.38, y: y + 0.18, w: 0.46, h: 0.46,
      fontSize: 18, align: "center", valign: "middle", margin: 0,
    });

    // Title
    s.addText(f.title, {
      x: 1.94, y: y + 0.07, w: 3.2, h: 0.38,
      fontSize: 15, bold: true, color: C.textDark,
      fontFace: "Microsoft YaHei", margin: 0,
    });

    // Desc
    s.addText(f.desc, {
      x: 5.3, y: y + 0.04, w: 4.2, h: 0.8,
      fontSize: 11.5, color: C.textMid,
      fontFace: "Microsoft YaHei", wrap: true, valign: "middle",
    });
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 5 — 工作流程（横向流程）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lavender };

  s.addText("工作流程", {
    x: 0.55, y: 0.22, w: 9, h: 0.6,
    fontSize: 30, bold: true, color: C.textDark,
    fontFace: "Microsoft YaHei", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 0.82, w: 3.0, h: 0.05,
    fill: { color: C.purple }, line: { color: C.purple },
  });
  s.addText("五步全自动 · 无需人工干预 · 分钟级完成", {
    x: 0.55, y: 0.96, w: 9, h: 0.36,
    fontSize: 13.5, color: C.textMute,
    fontFace: "Microsoft YaHei", margin: 0,
  });

  const steps = [
    { num: "1", title: "导入\n内容", detail: "PDF / DOCX\nURL / Markdown", icon: "📥", color: C.purple },
    { num: "2", title: "AI\n策略", detail: "结构规划\n视觉方案", icon: "🧠", color: C.indigo },
    { num: "3", title: "图像\n设计", detail: "配色 / 布局\n图像生成", icon: "🎨", color: C.rose },
    { num: "4", title: "代码\n生成", detail: "DrawingML\nSVG 渲染", icon: "⚙️", color: C.cyan },
    { num: "5", title: "导出\nPPTX", detail: "原生可编辑\n标准格式", icon: "📤", color: C.green },
  ];

  const sW = 1.56, sH = 3.6, startX = 0.35, sY = 1.5;

  steps.forEach((st, i) => {
    const x = startX + i * (sW + 0.26);

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: sY, w: sW, h: sH,
      fill: { color: C.white }, line: { color: C.border, width: 1 },
      shadow: mk(),
    });

    // Top color fill
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: sY, w: sW, h: 1.18,
      fill: { color: st.color }, line: { color: st.color },
    });

    // Number (large, inside color area)
    s.addText(st.num, {
      x, y: sY + 0.04, w: sW, h: 0.58,
      fontSize: 32, bold: true, color: "FFFFFF",
      fontFace: "Calibri", align: "center", margin: 0,
      opacity: 0.3,
    });

    // Icon circle
    s.addShape(pres.shapes.OVAL, {
      x: x + (sW - 0.76) / 2, y: sY + 0.66, w: 0.76, h: 0.76,
      fill: { color: C.white }, line: { color: C.white },
    });
    s.addText(st.icon, {
      x: x + (sW - 0.76) / 2, y: sY + 0.66, w: 0.76, h: 0.76,
      fontSize: 22, align: "center", valign: "middle", margin: 0,
    });

    // Title inside white area
    s.addText(st.title, {
      x: x + 0.1, y: sY + 1.58, w: sW - 0.2, h: 0.78,
      fontSize: 14, bold: true, color: C.textDark,
      fontFace: "Microsoft YaHei", align: "center", margin: 0,
    });

    // Detail
    s.addText(st.detail, {
      x: x + 0.1, y: sY + 2.44, w: sW - 0.2, h: 0.9,
      fontSize: 11.5, color: C.textMute,
      fontFace: "Microsoft YaHei", align: "center", wrap: true,
    });

    // Arrow
    if (i < steps.length - 1) {
      s.addText("→", {
        x: x + sW + 0.04, y: sY + sH / 2 - 0.24, w: 0.24, h: 0.48,
        fontSize: 18, color: st.color, align: "center", margin: 0,
      });
    }
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 6 — 技术亮点（卡片 + 图标，深色）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.plum };

  // Deco
  s.addShape(pres.shapes.OVAL, {
    x: 8.5, y: 2.5, w: 4, h: 4,
    fill: { color: C.indigo, transparency: 90 },
    line: { color: C.indigo, transparency: 85, width: 1 },
  });

  s.addText("技术亮点", {
    x: 0.55, y: 0.22, w: 9, h: 0.6,
    fontSize: 30, bold: true, color: C.white,
    fontFace: "Microsoft YaHei", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 0.82, w: 1.6, h: 0.05,
    fill: { color: C.rose }, line: { color: C.rose },
  });
  s.addText("三大核心技术支柱", {
    x: 0.55, y: 0.96, w: 9, h: 0.36,
    fontSize: 13.5, color: "A78BFA",
    fontFace: "Microsoft YaHei", margin: 0,
  });

  const highs = [
    {
      icon: "🤖",
      title: "多角色 AI 协作",
      color: C.purple,
      items: ["策略师：规划内容结构与视觉方案", "图像生成师：创作定制视觉素材", "执行师：精准实现 DrawingML 代码", "质量门禁：自动校验每页输出"],
    },
    {
      icon: "📐",
      title: "DrawingML 原生输出",
      color: C.cyan,
      items: ["输出真实 PowerPoint 形状对象", "非截图、非图片，完整保留编辑能力", "兼容 Office 全版本及 WPS", "每个元素均可独立选中与修改"],
    },
    {
      icon: "📂",
      title: "全格式输入支持",
      color: C.amber,
      items: ["PDF 文档智能解析与排版还原", "DOCX / Word 直接结构化导入", "网页 URL 一键抓取与内容提炼", "Markdown 原生支持与语义理解"],
    },
  ];

  highs.forEach((h, i) => {
    const x = 0.35 + i * 3.22;
    const cW = 2.98;
    const cH = 4.3;
    const cY = 1.45;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cY, w: cW, h: cH,
      fill: { color: "1E1035" }, line: { color: h.color, width: 1 },
      shadow: mk(),
    });

    // Top accent line
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cY, w: cW, h: 0.07,
      fill: { color: h.color }, line: { color: h.color },
    });

    // Icon circle
    s.addShape(pres.shapes.OVAL, {
      x: x + (cW - 0.76) / 2, y: cY + 0.2, w: 0.76, h: 0.76,
      fill: { color: h.color }, line: { color: h.color },
    });
    s.addText(h.icon, {
      x: x + (cW - 0.76) / 2, y: cY + 0.2, w: 0.76, h: 0.76,
      fontSize: 22, align: "center", valign: "middle", margin: 0,
    });

    s.addText(h.title, {
      x: x + 0.15, y: cY + 1.1, w: cW - 0.3, h: 0.48,
      fontSize: 15, bold: true, color: C.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0,
    });

    const bullets = h.items.map((it, bi) => ({
      text: it,
      options: { bullet: true, breakLine: bi < h.items.length - 1 },
    }));
    s.addText(bullets, {
      x: x + 0.2, y: cY + 1.68, w: cW - 0.4, h: 2.4,
      fontSize: 12.5, color: "C4B5FD",
      fontFace: "Microsoft YaHei", valign: "top",
      paraSpaceAfter: 7,
    });
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 7 — 核心数据（大数字，白底）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.22, h: 5.625,
    fill: { color: C.rose }, line: { color: C.rose },
  });

  s.addText("核心数据", {
    x: 0.45, y: 0.22, w: 9, h: 0.6,
    fontSize: 30, bold: true, color: C.textDark,
    fontFace: "Microsoft YaHei", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 0.82, w: 3.0, h: 0.05,
    fill: { color: C.rose }, line: { color: C.rose },
  });
  s.addText("真实表现，数字说话", {
    x: 0.45, y: 0.96, w: 9, h: 0.36,
    fontSize: 13.5, color: C.textMute,
    fontFace: "Microsoft YaHei", margin: 0,
  });

  const stats = [
    { num: "10×", label: "效率提升", sub: "vs. 传统手工制作", color: C.purple },
    { num: "<5", unit: "分钟", label: "生成时间", sub: "完整 12 页 PPT", color: C.rose },
    { num: "100%", label: "可编辑率", sub: "原生 DrawingML 输出", color: C.cyan },
    { num: "5+", label: "输入格式", sub: "PDF · DOCX · URL · MD", color: C.amber },
  ];

  stats.forEach((st, i) => {
    const x = 0.45 + i * 2.38;
    const cW = 2.16;
    const cH = 3.72;
    const cY = 1.52;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cY, w: cW, h: cH,
      fill: { color: C.lavender }, line: { color: st.color, width: 1 },
      shadow: mk(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cY, w: cW, h: 0.07,
      fill: { color: st.color }, line: { color: st.color },
    });

    s.addText(st.num, {
      x, y: cY + 0.5, w: cW, h: 1.1,
      fontSize: 44, bold: true, color: st.color,
      fontFace: "Calibri", align: "center", margin: 0,
    });

    if (st.unit) {
      s.addText(st.unit, {
        x, y: cY + 1.6, w: cW, h: 0.36,
        fontSize: 14, color: st.color,
        fontFace: "Microsoft YaHei", align: "center", margin: 0,
      });
    }

    s.addText(st.label, {
      x, y: cY + 2.06, w: cW, h: 0.44,
      fontSize: 15, bold: true, color: C.textDark,
      fontFace: "Microsoft YaHei", align: "center", margin: 0,
    });

    s.addText(st.sub, {
      x: x + 0.08, y: cY + 2.55, w: cW - 0.16, h: 0.9,
      fontSize: 11, color: C.textMute,
      fontFace: "Microsoft YaHei", align: "center", wrap: true,
    });
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 8 — 应用场景（彩色背景卡片）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lavender };

  s.addText("应用场景", {
    x: 0.55, y: 0.22, w: 9, h: 0.6,
    fontSize: 30, bold: true, color: C.textDark,
    fontFace: "Microsoft YaHei", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 0.82, w: 3.0, h: 0.05,
    fill: { color: C.purple }, line: { color: C.purple },
  });
  s.addText("覆盖主流演示需求场景", {
    x: 0.55, y: 0.96, w: 9, h: 0.36,
    fontSize: 13.5, color: C.textMute,
    fontFace: "Microsoft YaHei", margin: 0,
  });

  const scenes = [
    {
      icon: "💼",
      title: "商务汇报",
      color: C.purple,
      bgGrad: "2D0B5C",
      tags: ["季度复盘", "项目提案", "投资路演"],
      desc: "快速将业务数据、调研报告转化为专业商务风格汇报 PPT，为决策者提供清晰的视觉化呈现，提升会议效率",
    },
    {
      icon: "🎓",
      title: "学术展示",
      color: C.rose,
      bgGrad: "5C0B2D",
      tags: ["论文答辩", "课题汇报", "学术会议"],
      desc: "将学术论文、研究报告自动转换为逻辑清晰、图文并茂的学术演示文稿，有效传递研究价值",
    },
    {
      icon: "🚀",
      title: "产品发布",
      color: C.cyan,
      bgGrad: "0B3B5C",
      tags: ["新品发布", "功能演示", "用户手册"],
      desc: "将产品文档、功能说明迅速生成高颜值发布演示，大幅缩短上市准备周期，第一时间抓住市场窗口",
    },
  ];

  scenes.forEach((sc, i) => {
    const x = 0.35 + i * 3.22;
    const cW = 2.98;
    const cY = 1.45;
    const cH = 3.95;

    // Card dark bg
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cY, w: cW, h: cH,
      fill: { color: sc.bgGrad }, line: { color: sc.color, width: 1 },
      shadow: mk(),
    });

    // Icon circle
    s.addShape(pres.shapes.OVAL, {
      x: x + (cW - 0.82) / 2, y: cY + 0.22, w: 0.82, h: 0.82,
      fill: { color: sc.color }, line: { color: sc.color },
    });
    s.addText(sc.icon, {
      x: x + (cW - 0.82) / 2, y: cY + 0.22, w: 0.82, h: 0.82,
      fontSize: 24, align: "center", valign: "middle", margin: 0,
    });

    s.addText(sc.title, {
      x: x + 0.15, y: cY + 1.18, w: cW - 0.3, h: 0.48,
      fontSize: 18, bold: true, color: C.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0,
    });

    // Tags
    const tagW = (cW - 0.48) / 3;
    sc.tags.forEach((tag, ti) => {
      const tagX = x + 0.18 + ti * (tagW + 0.06);
      s.addShape(pres.shapes.RECTANGLE, {
        x: tagX, y: cY + 1.78, w: tagW, h: 0.3,
        fill: { color: sc.color, transparency: 50 }, line: { color: sc.color },
      });
      s.addText(tag, {
        x: tagX, y: cY + 1.78, w: tagW, h: 0.3,
        fontSize: 9.5, color: C.white, bold: true,
        align: "center", valign: "middle", fontFace: "Microsoft YaHei", margin: 0,
      });
    });

    s.addText(sc.desc, {
      x: x + 0.18, y: cY + 2.22, w: cW - 0.36, h: 1.56,
      fontSize: 12.5, color: "C4B5FD",
      fontFace: "Microsoft YaHei", wrap: true, valign: "top",
    });
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 9 — CTA 结尾（双色分屏）
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();

  // Left half — purple
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 5, h: 5.625,
    fill: { color: C.purple }, line: { color: C.purple },
  });
  // Right half — plum
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5, y: 0, w: 5, h: 5.625,
    fill: { color: C.plum }, line: { color: C.plum },
  });

  // Deco circle center
  s.addShape(pres.shapes.OVAL, {
    x: 3.3, y: 1.0, w: 3.4, h: 3.4,
    fill: { color: C.white, transparency: 94 },
    line: { color: C.white, transparency: 88, width: 1 },
  });

  // Left content
  s.addText("准备好了吗？", {
    x: 0.45, y: 0.85, w: 4.2, h: 0.48,
    fontSize: 16, color: "C4B5FD",
    fontFace: "Microsoft YaHei", margin: 0,
  });
  s.addText("让 AI\n替你做 PPT", {
    x: 0.45, y: 1.38, w: 4.2, h: 2.0,
    fontSize: 42, bold: true, color: C.white,
    fontFace: "Microsoft YaHei", margin: 0, valign: "top",
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 3.45, w: 1.6, h: 0.05,
    fill: { color: C.rose }, line: { color: C.rose },
  });
  s.addText("即刻体验，彻底告别 PPT 焦虑", {
    x: 0.45, y: 3.6, w: 4.2, h: 0.44,
    fontSize: 13, color: "C4B5FD",
    fontFace: "Microsoft YaHei", margin: 0,
  });

  // Primary button
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 4.22, w: 2.2, h: 0.62,
    fill: { color: C.rose }, line: { color: C.rose },
    shadow: { type: "outer", color: "EC4899", blur: 14, offset: 3, angle: 135, opacity: 0.5 },
  });
  s.addText("立即体验 →", {
    x: 0.45, y: 4.22, w: 2.2, h: 0.62,
    fontSize: 15, bold: true, color: C.white,
    fontFace: "Microsoft YaHei", align: "center", valign: "middle", margin: 0,
  });

  // Right content
  const checks = [
    { icon: "✓", text: "多格式输入，零门槛上手" },
    { icon: "✓", text: "分钟级生成，十倍效率" },
    { icon: "✓", text: "原生可编辑，完全掌控" },
    { icon: "✓", text: "专业视觉，媲美设计师" },
  ];
  checks.forEach((c, i) => {
    s.addShape(pres.shapes.OVAL, {
      x: 5.35, y: 1.55 + i * 0.78, w: 0.38, h: 0.38,
      fill: { color: C.purple }, line: { color: C.purpleL },
    });
    s.addText(c.icon, {
      x: 5.35, y: 1.55 + i * 0.78, w: 0.38, h: 0.38,
      fontSize: 13, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.text, {
      x: 5.86, y: 1.57 + i * 0.78, w: 3.8, h: 0.36,
      fontSize: 14, color: "C4B5FD",
      fontFace: "Microsoft YaHei", margin: 0,
    });
  });

  // Bottom info
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5, y: 5.18, w: 5, h: 0.445,
    fill: { color: "0F0820" }, line: { color: "0F0820" },
  });
  s.addText("PPT Master  ·  AI-Powered Presentation Engine  ·  2026", {
    x: 5.15, y: 5.2, w: 4.7, h: 0.38,
    fontSize: 9.5, color: "7C6FA0", fontFace: "Calibri", margin: 0,
  });
}

// ── Write file ──────────────────────────────────────────────
pres.writeFile({ fileName: "AI生成PPT产品介绍.pptx" })
  .then(() => console.log("✅ 文件已生成：AI生成PPT产品介绍.pptx"))
  .catch((e) => console.error("❌ 生成失败：", e));
