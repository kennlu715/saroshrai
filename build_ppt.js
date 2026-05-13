const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaUsers, FaRocket, FaBullseye, FaBolt, FaSitemap, FaUserAstronaut,
  FaSeedling, FaHandsHelping, FaBrain, FaFire, FaCogs, FaEye,
  FaCompass, FaHeart, FaShieldAlt, FaLightbulb, FaArrowUp, FaCheckCircle
} = require("react-icons/fa");

// ====== 颜色方案: Midnight Executive ======
const NAVY = "1E2761";        // 深邃海军蓝（主色）
const NAVY_DEEP = "0F1747";   // 更深的导航色（背景层）
const ICE = "CADCFC";         // 冰蓝（辅色）
const WHITE = "FFFFFF";
const GOLD = "F5C76A";        // 暖金（强调色）
const TEXT_DARK = "1A1F3A";
const TEXT_BODY = "3D4663";
const TEXT_MUTED = "8590B0";
const BG_LIGHT = "F7F9FC";
const CARD_BG = "FFFFFF";
const DIVIDER = "E0E6F0";

// ====== 工具函数 ======
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}
async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}
const makeShadow = () => ({
  type: "outer", color: "1E2761", blur: 12, offset: 3, angle: 90, opacity: 0.10
});

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
  pres.author = "Team Sharing";
  pres.title = "组织与人才：AI时代的领导力变革";

  const W = 13.3, H = 7.5;

  // 预生成所有图标
  const icons = {
    users: await iconToBase64Png(FaUsers, "#" + NAVY, 256),
    rocket: await iconToBase64Png(FaRocket, "#" + NAVY, 256),
    bullseye: await iconToBase64Png(FaBullseye, "#" + NAVY, 256),
    bolt_gold: await iconToBase64Png(FaBolt, "#" + GOLD, 256),
    sitemap: await iconToBase64Png(FaSitemap, "#" + WHITE, 256),
    astronaut: await iconToBase64Png(FaUserAstronaut, "#" + WHITE, 256),
    seedling: await iconToBase64Png(FaSeedling, "#" + WHITE, 256),
    hands: await iconToBase64Png(FaHandsHelping, "#" + WHITE, 256),
    brain: await iconToBase64Png(FaBrain, "#" + NAVY, 256),
    fire: await iconToBase64Png(FaFire, "#" + NAVY, 256),
    cogs: await iconToBase64Png(FaCogs, "#" + NAVY, 256),
    eye: await iconToBase64Png(FaEye, "#" + GOLD, 256),
    compass: await iconToBase64Png(FaCompass, "#" + GOLD, 256),
    heart: await iconToBase64Png(FaHeart, "#" + GOLD, 256),
    shield: await iconToBase64Png(FaShieldAlt, "#" + GOLD, 256),
    lightbulb: await iconToBase64Png(FaLightbulb, "#" + GOLD, 256),
    arrowup: await iconToBase64Png(FaArrowUp, "#" + GOLD, 256),
    check: await iconToBase64Png(FaCheckCircle, "#" + GOLD, 256),
  };

  // ============================================================
  // SLIDE 1 — 理念与组织特点（深色主标题 + 三大特点卡片）
  // ============================================================
  const s1 = pres.addSlide();
  s1.background = { color: NAVY_DEEP };

  // 顶部细金线 + 章节标号
  s1.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 0.55, w: 0.5, h: 0.06, fill: { color: GOLD }, line: { color: GOLD, width: 0 }
  });
  s1.addText("01 / 组织与人才", {
    x: 1.3, y: 0.4, w: 6, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: GOLD, charSpacing: 4, bold: true, margin: 0
  });

  // 主标题（核心理念）
  s1.addText("坚持用户价值，用户为本，长期主义", {
    x: 0.7, y: 0.95, w: 12, h: 0.85,
    fontSize: 36, fontFace: "Microsoft YaHei", color: WHITE, bold: true, margin: 0
  });

  // 副标题装饰
  s1.addText("Mission · Vision · Long-Termism", {
    x: 0.7, y: 1.78, w: 12, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: ICE, italic: true, margin: 0
  });

  // 中部 quote 引导词
  s1.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.55, w: 0.08, h: 0.85, fill: { color: GOLD }, line: { color: GOLD, width: 0 }
  });
  s1.addText("我们相信怎样的组织，能跑得更远", {
    x: 0.95, y: 2.6, w: 8, h: 0.4,
    fontSize: 16, fontFace: "Microsoft YaHei", color: ICE, margin: 0
  });
  s1.addText("ORGANIZATIONAL DNA", {
    x: 0.95, y: 3.0, w: 8, h: 0.35,
    fontSize: 11, fontFace: "Arial", color: TEXT_MUTED, charSpacing: 6, margin: 0
  });

  // 三张特点卡片（横向排列）
  const cardY = 3.85;
  const cardH = 3.0;
  const cardW = 3.95;
  const gap = 0.25;
  const startX = (W - (cardW * 3 + gap * 2)) / 2;

  const cards = [
    {
      icon: icons.rocket,
      title: "小团队，做大事",
      kw: "敏捷 · 端到端",
      lines: [
        "小团队作战，专注力与决策速度并存",
        "端到端的生产模式，让反馈与迭代飞快"
      ]
    },
    {
      icon: icons.users,
      title: "能人越强，人才密度",
      kw: "Density · Compete",
      lines: [
        "能人越强，组织越能撬动复杂问题",
        "好的infra支持并行实验，形成良性竞赛"
      ]
    },
    {
      icon: icons.bullseye,
      title: "管事比管人更重要",
      kw: "Taste · Speed",
      lines: [
        "深度思考、创意、taste、动手能力",
        "关注事的进展与迭代速度，少做形式管理"
      ]
    }
  ];

  for (let i = 0; i < cards.length; i++) {
    const cx = startX + i * (cardW + gap);
    // 卡片底
    s1.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: WHITE }, line: { color: WHITE, width: 0 },
      shadow: makeShadow()
    });
    // 顶部金色装饰条
    s1.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.08,
      fill: { color: GOLD }, line: { color: GOLD, width: 0 }
    });
    // 编号
    s1.addText(`0${i + 1}`, {
      x: cx + 0.35, y: cardY + 0.25, w: 0.8, h: 0.5,
      fontSize: 30, fontFace: "Georgia", color: ICE, bold: true, italic: true, margin: 0
    });
    // 图标
    s1.addImage({
      data: cards[i].icon,
      x: cx + cardW - 0.95, y: cardY + 0.32, w: 0.6, h: 0.6
    });
    // 关键词（小标）
    s1.addText(cards[i].kw, {
      x: cx + 0.35, y: cardY + 0.95, w: cardW - 0.7, h: 0.3,
      fontSize: 10, fontFace: "Arial", color: GOLD, charSpacing: 3, bold: true, margin: 0
    });
    // 标题
    s1.addText(cards[i].title, {
      x: cx + 0.35, y: cardY + 1.25, w: cardW - 0.7, h: 0.55,
      fontSize: 20, fontFace: "Microsoft YaHei", color: NAVY, bold: true, margin: 0
    });
    // 分隔线
    s1.addShape(pres.shapes.LINE, {
      x: cx + 0.35, y: cardY + 1.92, w: 0.5, h: 0,
      line: { color: NAVY, width: 1.5 }
    });
    // 描述
    s1.addText([
      { text: cards[i].lines[0], options: { breakLine: true } },
      { text: cards[i].lines[1] }
    ], {
      x: cx + 0.35, y: cardY + 2.05, w: cardW - 0.7, h: 0.85,
      fontSize: 12, fontFace: "Microsoft YaHei", color: TEXT_BODY, paraSpaceAfter: 4,
      margin: 0, valign: "top"
    });
  }

  // 页脚
  s1.addText("Page 01 / 03", {
    x: 0.7, y: H - 0.45, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: TEXT_MUTED, charSpacing: 3, margin: 0
  });

  // ============================================================
  // SLIDE 2 — 四大变革方向（2×2 矩阵卡片）
  // ============================================================
  const s2 = pres.addSlide();
  s2.background = { color: BG_LIGHT };

  // 头部标识
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 0.55, w: 0.5, h: 0.06, fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s2.addText("02 / 四大变革方向", {
    x: 1.3, y: 0.4, w: 6, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: NAVY, charSpacing: 4, bold: true, margin: 0
  });

  // 主标题
  s2.addText("从干部、组织、人才到文化的系统重塑", {
    x: 0.7, y: 0.95, w: 12, h: 0.7,
    fontSize: 30, fontFace: "Microsoft YaHei", color: TEXT_DARK, bold: true, margin: 0
  });
  s2.addText("Four Pillars of Transformation in the AI Era", {
    x: 0.7, y: 1.65, w: 12, h: 0.35,
    fontSize: 13, fontFace: "Georgia", color: TEXT_MUTED, italic: true, margin: 0
  });

  // 2x2 卡片
  const m_topY = 2.3;
  const m_w = 6.05;
  const m_h = 2.35;
  const m_gap = 0.2;
  const m_x1 = 0.7;
  const m_x2 = m_x1 + m_w + m_gap;
  const m_y1 = m_topY;
  const m_y2 = m_topY + m_h + m_gap;

  const pillars = [
    {
      x: m_x1, y: m_y1,
      no: "I",
      icon: icons.sitemap,
      title: "干部革新",
      en: "LEADERSHIP RENEWAL",
      points: [
        "建立AI领导力，持续 raise the bar",
        "管理者亲自下场，先革自己的命"
      ]
    },
    {
      x: m_x2, y: m_y1,
      no: "II",
      icon: icons.astronaut,
      title: "组织重塑",
      en: "ORG RESHAPE",
      points: [
        "更扁平、压中间层，管理者转实干者",
        "鼓励大胆尝试，编制不增、AI增效"
      ]
    },
    {
      x: m_x1, y: m_y2,
      no: "III",
      icon: icons.seedling,
      title: "人才激活",
      en: "TALENT ACTIVATION",
      points: [
        "吸引AI原生人才，提升人才密度",
        "状态最好的人灵活配置在关键战场"
      ]
    },
    {
      x: m_x2, y: m_y2,
      no: "IV",
      icon: icons.hands,
      title: "文化升级",
      en: "CULTURE UPGRADE",
      points: [
        "坦诚透明、low-ego、团队 co-design",
        "战时状态：直面问题、不粉饰、做扎实"
      ]
    }
  ];

  for (const p of pillars) {
    // 卡片背景
    s2.addShape(pres.shapes.RECTANGLE, {
      x: p.x, y: p.y, w: m_w, h: m_h,
      fill: { color: CARD_BG }, line: { color: WHITE, width: 0 },
      shadow: makeShadow()
    });
    // 左侧深蓝条带（图标区）
    s2.addShape(pres.shapes.RECTANGLE, {
      x: p.x, y: p.y, w: 1.45, h: m_h,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 }
    });
    // 罗马编号（背景大字）
    s2.addText(p.no, {
      x: p.x + 0.05, y: p.y + 0.1, w: 1.4, h: 0.6,
      fontSize: 16, fontFace: "Georgia", color: GOLD, bold: true, italic: true,
      align: "center", margin: 0
    });
    // 图标
    s2.addImage({
      data: p.icon,
      x: p.x + (1.45 - 0.78) / 2, y: p.y + 0.85, w: 0.78, h: 0.78
    });
    // 右侧内容区
    const cx = p.x + 1.65;
    const cw = m_w - 1.85;
    s2.addText(p.en, {
      x: cx, y: p.y + 0.25, w: cw, h: 0.3,
      fontSize: 10, fontFace: "Arial", color: GOLD, charSpacing: 4, bold: true, margin: 0
    });
    s2.addText(p.title, {
      x: cx, y: p.y + 0.55, w: cw, h: 0.55,
      fontSize: 22, fontFace: "Microsoft YaHei", color: NAVY, bold: true, margin: 0
    });
    s2.addShape(pres.shapes.LINE, {
      x: cx, y: p.y + 1.15, w: 0.45, h: 0,
      line: { color: GOLD, width: 2 }
    });
    s2.addText([
      { text: p.points[0], options: { bullet: { code: "25A0" }, breakLine: true } },
      { text: p.points[1], options: { bullet: { code: "25A0" } } }
    ], {
      x: cx, y: p.y + 1.3, w: cw, h: 0.95,
      fontSize: 12.5, fontFace: "Microsoft YaHei", color: TEXT_BODY,
      paraSpaceAfter: 4, valign: "top", margin: 0
    });
  }

  // 底部 takeaway 横条
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: H - 0.95, w: W - 1.4, h: 0.55,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s2.addImage({
    data: icons.bolt_gold,
    x: 0.95, y: H - 0.85, w: 0.32, h: 0.32
  });
  s2.addText("战时状态，凝聚胜利信心 — 直面问题、正视差距、不回避、不粉饰、在实践中进步", {
    x: 1.4, y: H - 0.92, w: W - 2.2, h: 0.5,
    fontSize: 13, fontFace: "Microsoft YaHei", color: WHITE, bold: true,
    valign: "middle", margin: 0
  });

  s2.addText("Page 02 / 03", {
    x: 0.7, y: H - 0.3, w: 4, h: 0.25,
    fontSize: 10, fontFace: "Arial", color: TEXT_MUTED, charSpacing: 3, margin: 0
  });

  // ============================================================
  // SLIDE 3 — AI领导力（左核心 + 右行动准则）
  // ============================================================
  const s3 = pres.addSlide();
  s3.background = { color: WHITE };

  // 左半部分：深色背景 + AI领导力的三个能力
  s3.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 5.6, h: H,
    fill: { color: NAVY_DEEP }, line: { color: NAVY_DEEP, width: 0 }
  });
  // 左上金色细条
  s3.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 0.55, w: 0.5, h: 0.06, fill: { color: GOLD }, line: { color: GOLD, width: 0 }
  });
  s3.addText("03 / AI 领导力", {
    x: 1.3, y: 0.4, w: 4, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: GOLD, charSpacing: 4, bold: true, margin: 0
  });

  s3.addText("AI 时代的", {
    x: 0.7, y: 0.95, w: 4.5, h: 0.55,
    fontSize: 22, fontFace: "Microsoft YaHei", color: ICE, margin: 0
  });
  s3.addText("领导力三力", {
    x: 0.7, y: 1.45, w: 4.5, h: 0.85,
    fontSize: 38, fontFace: "Microsoft YaHei", color: WHITE, bold: true, margin: 0
  });
  s3.addText("Leadership in the Age of AI", {
    x: 0.7, y: 2.3, w: 4.5, h: 0.35,
    fontSize: 12, fontFace: "Georgia", color: GOLD, italic: true, charSpacing: 2, margin: 0
  });

  // 三力（垂直排布）
  const triItems = [
    { icon: icons.compass, title: "判断力", en: "JUDGMENT", desc: "在不确定中识别真问题，做出正确取舍" },
    { icon: icons.heart, title: "感召力", en: "INSPIRATION", desc: "让最想做事的人聚拢，激发团队的状态" },
    { icon: icons.shield, title: "担责力", en: "ACCOUNTABILITY", desc: "面对不确定与风险，敢于承担并交付结果" }
  ];

  let tY = 3.0;
  for (const t of triItems) {
    // 图标圆底
    s3.addShape(pres.shapes.OVAL, {
      x: 0.7, y: tY, w: 0.7, h: 0.7,
      fill: { color: NAVY }, line: { color: GOLD, width: 1.5 }
    });
    s3.addImage({ data: t.icon, x: 0.86, y: tY + 0.16, w: 0.38, h: 0.38 });
    // 文本
    s3.addText(t.title, {
      x: 1.55, y: tY + 0.02, w: 3.6, h: 0.4,
      fontSize: 20, fontFace: "Microsoft YaHei", color: WHITE, bold: true, margin: 0
    });
    s3.addText(t.en, {
      x: 1.55, y: tY + 0.42, w: 3.6, h: 0.25,
      fontSize: 9, fontFace: "Arial", color: GOLD, charSpacing: 3, bold: true, margin: 0
    });
    s3.addText(t.desc, {
      x: 1.55, y: tY + 0.67, w: 3.85, h: 0.45,
      fontSize: 11, fontFace: "Microsoft YaHei", color: ICE, margin: 0
    });
    tY += 1.35;
  }

  // 右半部分：行动准则
  s3.addText("管理者的四条行动准则", {
    x: 6.0, y: 0.7, w: 6.7, h: 0.55,
    fontSize: 22, fontFace: "Microsoft YaHei", color: TEXT_DARK, bold: true, margin: 0
  });
  s3.addText("FOUR ACTION PRINCIPLES FOR LEADERS", {
    x: 6.0, y: 1.22, w: 6.7, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: NAVY, charSpacing: 4, bold: true, margin: 0
  });
  s3.addShape(pres.shapes.LINE, {
    x: 6.0, y: 1.55, w: 0.6, h: 0,
    line: { color: GOLD, width: 2.5 }
  });

  // 四条准则
  const rules = [
    {
      icon: icons.lightbulb,
      title: "亲手干，深度使用 AI",
      desc: "不只听汇报，自己上手用、迭代、感受 AI 的能力边界与机会点"
    },
    {
      icon: icons.cogs,
      title: "重构组织与流程",
      desc: "零基思维设计层级、审批、会议、流程，让组织更轻、更快"
    },
    {
      icon: icons.brain,
      title: "做好 人 + Agent 的配置",
      desc: "AI 能做的交给 AI，把人放到判断、创造、协同的高价值环节"
    },
    {
      icon: icons.fire,
      title: "让状态最好的人上场",
      desc: "营造坦诚共创的环境，感召优秀团队，驱动现有人才转型"
    }
  ];

  let rY = 1.85;
  const rH = 1.25;
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    // 卡片底
    s3.addShape(pres.shapes.RECTANGLE, {
      x: 6.0, y: rY, w: 6.7, h: rH,
      fill: { color: BG_LIGHT }, line: { color: BG_LIGHT, width: 0 }
    });
    // 左侧金色短条
    s3.addShape(pres.shapes.RECTANGLE, {
      x: 6.0, y: rY, w: 0.08, h: rH,
      fill: { color: GOLD }, line: { color: GOLD, width: 0 }
    });
    // 编号
    s3.addText(`0${i + 1}`, {
      x: 6.25, y: rY + 0.18, w: 0.7, h: 0.45,
      fontSize: 22, fontFace: "Georgia", color: ICE, bold: true, italic: true, margin: 0
    });
    // 图标
    s3.addImage({
      data: r.icon,
      x: 11.95, y: rY + 0.32, w: 0.55, h: 0.55
    });
    // 标题
    s3.addText(r.title, {
      x: 7.05, y: rY + 0.18, w: 4.7, h: 0.45,
      fontSize: 16, fontFace: "Microsoft YaHei", color: NAVY, bold: true, margin: 0
    });
    // 描述
    s3.addText(r.desc, {
      x: 7.05, y: rY + 0.62, w: 4.7, h: 0.6,
      fontSize: 11.5, fontFace: "Microsoft YaHei", color: TEXT_BODY, margin: 0
    });
    rY += rH + 0.1;
  }

  // 底部页脚
  s3.addText("Page 03 / 03", {
    x: 6.0, y: H - 0.4, w: 4, h: 0.25,
    fontSize: 10, fontFace: "Arial", color: TEXT_MUTED, charSpacing: 3, margin: 0
  });
  s3.addText("Team Sharing · 组织与人才", {
    x: W - 4.7, y: H - 0.4, w: 4, h: 0.25,
    fontSize: 10, fontFace: "Arial", color: TEXT_MUTED, charSpacing: 3, align: "right", margin: 0
  });

  // 输出
  await pres.writeFile({ fileName: "组织与人才_AI时代领导力分享.pptx" });
  console.log("✅ PPT 生成完成: 组织与人才_AI时代领导力分享.pptx");
})();
