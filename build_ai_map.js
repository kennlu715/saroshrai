const PptxGenJS = require("pptxgenjs");
const pptx = new PptxGenJS();

// ===== 全局设置 =====
const THEME = {
  bg: "FFFFFF",
  cardBg: "FFFFFF",
  cardBorder: "E8E8E8",
  title: "1A1A1A",
  subtitle: "666666",
  desc: "333333",
  tagBg: "F0F5FF",
  tagText: "2B6CB0",
  tagBg2: "F0FFF4",
  tagText2: "276749",
  tagBg3: "FFFAF0",
  tagText3: "C05621",
  chapterNum: "A0AEC0",
  chapterTitle: "1A1A1A",
  chapterEn: "A0AEC0",
  link: "3182CE",
  sectionIcon: "4A5568",
};
const W = 13.33;
const H = 7.5;

pptx.layout = "LAYOUT_16x9";
pptx.author = "腾讯";
pptx.title = "鹅厂AI工具探索地图";

// ===== 辅助函数 =====
function addCard(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: THEME.cardBg },
    line: { color: THEME.cardBorder, width: 1 },
    radius: 0.08
  });
}

function addTag(slide, text, x, y, w, h, type = "blue") {
  const bg = type === "green" ? THEME.tagBg2 : type === "orange" ? THEME.tagBg3 : THEME.tagBg;
  const txt = type === "green" ? THEME.tagText2 : type === "orange" ? THEME.tagText3 : THEME.tagText;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: bg },
    line: { color: bg },
    radius: 0.05
  });
  slide.addText(text, {
    x, y, w, h,
    fontSize: 8,
    color: txt,
    align: "center",
    valign: "middle",
    fontFace: "微软雅黑"
  });
}

function addLink(slide, url, x, y, w, h) {
  // 显示链接文字
  slide.addText(url, {
    x, y, w, h,
    fontSize: 9,
    color: THEME.link,
    fontFace: "微软雅黑",
    hyperlink: url.startsWith("http") ? { url: url } : { url: "https://" + url }
  });
}

function addSectionHeader(slide, chapterNum, title, enTitle, y) {
  slide.addText(chapterNum, {
    x: 0.5, y, w: 12, h: 0.25,
    fontSize: 9, color: THEME.chapterNum, fontFace: "微软雅黑", letterSpacing: 2
  });
  slide.addText(title + "  " + enTitle, {
    x: 0.5, y: y + 0.25, w: 12, h: 0.4,
    fontSize: 16, bold: true, color: THEME.chapterTitle, fontFace: "微软雅黑"
  });
}

// ============================================================
// 幻灯片1：封面
// ============================================================
{
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: THEME.bg }, line: { color: THEME.bg } });

  slide.addText("鹅厂 AI 工具探索地图", {
    x: 0.5, y: 2.2, w: 12, h: 0.8,
    fontSize: 32, bold: true, color: THEME.title, fontFace: "微软雅黑"
  });
  slide.addText("从编程开发到游戏生产 — 全场景 AI 工具矩阵", {
    x: 0.5, y: 3.1, w: 12, h: 0.4,
    fontSize: 14, color: THEME.subtitle, fontFace: "微软雅黑"
  });

  const tags = [
    { text: "编程开发", color: "EBF8FF", textColor: "2B6CB0" },
    { text: "设计创意", color: "FAF5FF", textColor: "6B46C1" },
    { text: "智能办公", color: "FFFAF0", textColor: "C05621" },
    { text: "游戏生产", color: "F0FFF4", textColor: "276749" },
  ];
  let tx = 0.5;
  tags.forEach((t) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: tx, y: 3.8, w: 1.6, h: 0.4,
      fill: { color: t.color },
      line: { color: t.color },
      radius: 0.1
    });
    slide.addText(t.text, {
      x: tx, y: 3.8, w: 1.6, h: 0.4,
      fontSize: 10, bold: true, color: t.textColor, align: "center", valign: "middle", fontFace: "微软雅黑"
    });
    tx += 1.9;
  });

  slide.addText("腾讯内部 · AI 工具全景探索     |     数据来源：腾讯官方 & 内网公开资料     |     2026.05", {
    x: 0, y: 7.0, w: W, h: 0.3,
    fontSize: 8, color: THEME.subtitle, align: "center", valign: "middle", fontFace: "微软雅黑"
  });
}

// ============================================================
// 幻灯片2：编程开发 CODING & BUILD
// ============================================================
{
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: THEME.bg }, line: { color: THEME.bg } });

  addSectionHeader(slide, "CHAPTER 02", "编程开发", "CODING & BUILD", 0.3);

  const tools = [
    {
      name: "CodeBuddy",
      desc: "一站式 AI 编码助手，含 IDE、插件、CLI 多种形态",
      link: "https://codebuddy.cn",
      tags: [{ text: "首页", type: "blue" }, { text: "IDE", type: "blue" }, { text: "插件", type: "blue" }, { text: "CLI", type: "blue" }],
      articles: [{ text: "全家桶介绍", type: "green" }, { text: "IDE版", type: "green" }, { text: "插件版", type: "green" }, { text: "CLI版", type: "green" }],
      cases: [{ text: "医疗团队AI全流程落地", type: "orange" }, { text: "养出能思考的专家", type: "orange" }]
    },
    {
      name: "WorkBuddy",
      desc: "非开发者的 AI 工作站，不写代码也能用 AI 干活",
      link: "https://workbuddy.woa.com",
      tags: [{ text: "一键安装", type: "blue" }, { text: "Claw", type: "blue" }, { text: "交流群", type: "blue" }, { text: "官方文档", type: "blue" }],
      articles: [{ text: "开虾直播回看", type: "green" }],
      cases: [{ text: "案例合集 K吧", type: "orange" }]
    },
    {
      name: "With",
      desc: "对话即开发，需求即代码",
      link: "https://with.woa.com",
      tags: [{ text: "官网", type: "blue" }],
      articles: [{ text: "使用指引", type: "green" }],
      cases: []
    },
    {
      name: "Knot",
      desc: "AI 应用开发平台",
      link: "https://knot.woa.com",
      tags: [{ text: "官网", type: "blue" }],
      articles: [{ text: "直播回看", type: "green" }],
      cases: []
    },
  ];

  const cardW = 6.0, cardH = 2.6, startX = 0.5, startY = 1.0, gapX = 0.3, gapY = 0.25;
  tools.forEach((tool, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    addCard(slide, x, y, cardW, cardH);

    // 工具名
    slide.addText(tool.name, {
      x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.35,
      fontSize: 14, bold: true, color: THEME.title, fontFace: "微软雅黑"
    });
    // 描述
    slide.addText(tool.desc, {
      x: x + 0.2, y: y + 0.5, w: cardW - 0.4, h: 0.3,
      fontSize: 10, color: THEME.desc, fontFace: "微软雅黑"
    });
    // 官网链接
    addLink(slide, tool.link, x + 0.2, y + 0.85, 1.5, 0.25);

    // 标签行
    let tx = x + 0.2, ty = y + 1.15;
    tool.tags.forEach((tag) => {
      const tw = tag.text.length * 0.18 + 0.3;
      addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
      tx += tw + 0.1;
    });

    // 官方直播/文章
    if (tool.articles.length > 0) {
      slide.addText("🎬 官方直播 / 文章", {
        x: x + 0.2, y: y + 1.5, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 1.75;
      tool.articles.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }

    // 用户案例
    if (tool.cases.length > 0) {
      slide.addText("📝 用户案例", {
        x: x + 0.2, y: y + 2.05, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 2.3;
      tool.cases.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }
  });
}

// ============================================================
// 幻灯片3：设计与创意 DESIGN & CREATE
// ============================================================
{
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: THEME.bg }, line: { color: THEME.bg } });

  addSectionHeader(slide, "CHAPTER 03", "设计与创意", "DESIGN & CREATE", 0.3);

  const tools = [
    {
      name: "Ardot",
      desc: "智能设计工具，让灵感鲜活落地",
      link: "https://ardot.woa.com",
      tags: [{ text: "官网", type: "blue" }],
      articles: [],
      cases: []
    },
    {
      name: "三谦",
      desc: "让设计师专注于创造",
      link: "https://sanquan.woa.com",
      tags: [{ text: "官网", type: "blue" }],
      articles: [],
      cases: []
    },
    {
      name: "XDream",
      desc: "多模态内容创作平台，文生图/图生视频/文生3D",
      link: "https://xdream.woa.com",
      tags: [{ text: "官网", type: "blue" }, { text: "AIGC", type: "blue" }],
      articles: [{ text: "游戏资产生成", type: "green" }],
      cases: []
    },
    {
      name: "Miora",
      desc: "通用 AI 创作助手，文案/PPT/脚本/数据分析",
      link: "https://miora.woa.com",
      tags: [{ text: "官网", type: "blue" }],
      articles: [],
      cases: []
    },
  ];

  const cardW = 6.0, cardH = 2.6, startX = 0.5, startY = 1.0, gapX = 0.3, gapY = 0.25;
  tools.forEach((tool, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    addCard(slide, x, y, cardW, cardH);

    slide.addText(tool.name, {
      x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.35,
      fontSize: 14, bold: true, color: THEME.title, fontFace: "微软雅黑"
    });
    slide.addText(tool.desc, {
      x: x + 0.2, y: y + 0.5, w: cardW - 0.4, h: 0.3,
      fontSize: 10, color: THEME.desc, fontFace: "微软雅黑"
    });
    addLink(slide, tool.link, x + 0.2, y + 0.85, 1.5, 0.25);

    let tx = x + 0.2, ty = y + 1.15;
    tool.tags.forEach((tag) => {
      const tw = tag.text.length * 0.18 + 0.3;
      addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
      tx += tw + 0.1;
    });

    if (tool.articles.length > 0) {
      slide.addText("🎬 官方直播 / 文章", {
        x: x + 0.2, y: y + 1.5, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 1.75;
      tool.articles.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }

    if (tool.cases.length > 0) {
      slide.addText("📝 用户案例", {
        x: x + 0.2, y: y + 2.05, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 2.3;
      tool.cases.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }
  });
}

// ============================================================
// 幻灯片4：智能办公 SMART OFFICE
// ============================================================
{
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: THEME.bg }, line: { color: THEME.bg } });

  addSectionHeader(slide, "CHAPTER 04", "智能办公", "SMART OFFICE", 0.3);

  const tools = [
    {
      name: "混元大模型",
      desc: "腾讯全链路自研大模型，API/SDK 接入，支持私有化部署",
      link: "https://hunyuan.qq.com",
      tags: [{ text: "Turbo版", type: "blue" }, { text: "API", type: "blue" }, { text: "私有化", type: "blue" }],
      articles: [{ text: "技术白皮书", type: "green" }],
      cases: []
    },
    {
      name: "元宝",
      desc: "随身智能助手，文档解析1000万字，MAU 1300万+",
      link: "https://yuanbao.qq.com",
      tags: [{ text: "App", type: "blue" }, { text: "小程序", type: "blue" }],
      articles: [{ text: "使用技巧", type: "green" }],
      cases: [{ text: "运营提效案例", type: "orange" }]
    },
    {
      name: "OpenClaw 内网版",
      desc: "私人 AI 助理，内网知识库接入，数据不出内网",
      link: "https://openclaw.woa.com",
      tags: [{ text: "内网", type: "blue" }, { text: "安全合规", type: "blue" }],
      articles: [],
      cases: []
    },
    {
      name: "iWiki + AI",
      desc: "团队知识协作平台，AI 写作助手，知识图谱构建",
      link: "https://iwiki.woa.com",
      tags: [{ text: "写作", type: "blue" }, { text: "知识库", type: "blue" }],
      articles: [],
      cases: []
    },
  ];

  const cardW = 6.0, cardH = 2.6, startX = 0.5, startY = 1.0, gapX = 0.3, gapY = 0.25;
  tools.forEach((tool, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    addCard(slide, x, y, cardW, cardH);

    slide.addText(tool.name, {
      x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.35,
      fontSize: 14, bold: true, color: THEME.title, fontFace: "微软雅黑"
    });
    slide.addText(tool.desc, {
      x: x + 0.2, y: y + 0.5, w: cardW - 0.4, h: 0.3,
      fontSize: 10, color: THEME.desc, fontFace: "微软雅黑"
    });
    addLink(slide, tool.link, x + 0.2, y + 0.85, 1.5, 0.25);

    let tx = x + 0.2, ty = y + 1.15;
    tool.tags.forEach((tag) => {
      const tw = tag.text.length * 0.18 + 0.3;
      addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
      tx += tw + 0.1;
    });

    if (tool.articles.length > 0) {
      slide.addText("🎬 官方直播 / 文章", {
        x: x + 0.2, y: y + 1.5, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 1.75;
      tool.articles.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }

    if (tool.cases.length > 0) {
      slide.addText("📝 用户案例", {
        x: x + 0.2, y: y + 2.05, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 2.3;
      tool.cases.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }
  });
}

// ============================================================
// 幻灯片5：游戏生产 GAME PRODUCTION
// ============================================================
{
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: THEME.bg }, line: { color: THEME.bg } });

  addSectionHeader(slide, "CHAPTER 05", "游戏生产", "GAME PRODUCTION", 0.3);

  const tools = [
    {
      name: "VISVISE",
      desc: "全链路AI创作解决方案，动画/3D模型/智能NPC/数字人",
      link: "https://visvise.woa.com",
      tags: [{ text: "科隆展", type: "blue" }, { text: "3D生成", type: "blue" }],
      articles: [{ text: "技术解析", type: "green" }],
      cases: []
    },
    {
      name: "GiiNEX",
      desc: "游戏AI引擎，生成式AI+决策AI，全生命周期解决方案",
      link: "https://giinex.woa.com",
      tags: [{ text: "GDC2024", type: "blue" }, { text: "PC+移动", type: "blue" }],
      articles: [],
      cases: []
    },
    {
      name: "混元游戏",
      desc: "工业级AIGC游戏内容生产引擎，降本50-70%",
      link: "https://hunyuan.qq.com/game",
      tags: [{ text: "AIGC", type: "blue" }, { text: "降本", type: "blue" }],
      articles: [],
      cases: []
    },
    {
      name: "3D城市自动生成",
      desc: "AI自动生成游戏内3D城市环境，支持UGC关卡设计",
      link: "https://3dcity.woa.com",
      tags: [{ text: "UGC", type: "blue" }, { text: "关卡设计", type: "blue" }],
      articles: [],
      cases: []
    },
  ];

  const cardW = 6.0, cardH = 2.6, startX = 0.5, startY = 1.0, gapX = 0.3, gapY = 0.25;
  tools.forEach((tool, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    addCard(slide, x, y, cardW, cardH);

    slide.addText(tool.name, {
      x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.35,
      fontSize: 14, bold: true, color: THEME.title, fontFace: "微软雅黑"
    });
    slide.addText(tool.desc, {
      x: x + 0.2, y: y + 0.5, w: cardW - 0.4, h: 0.3,
      fontSize: 10, color: THEME.desc, fontFace: "微软雅黑"
    });
    addLink(slide, tool.link, x + 0.2, y + 0.85, 1.5, 0.25);

    let tx = x + 0.2, ty = y + 1.15;
    tool.tags.forEach((tag) => {
      const tw = tag.text.length * 0.18 + 0.3;
      addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
      tx += tw + 0.1;
    });

    if (tool.articles.length > 0) {
      slide.addText("🎬 官方直播 / 文章", {
        x: x + 0.2, y: y + 1.5, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 1.75;
      tool.articles.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }

    if (tool.cases.length > 0) {
      slide.addText("📝 用户案例", {
        x: x + 0.2, y: y + 2.05, w: 2.5, h: 0.22,
        fontSize: 8, color: THEME.subtitle, fontFace: "微软雅黑"
      });
      tx = x + 0.2; ty = y + 2.3;
      tool.cases.forEach((tag) => {
        const tw = tag.text.length * 0.18 + 0.3;
        addTag(slide, tag.text, tx, ty, tw, 0.28, tag.type);
        tx += tw + 0.1;
      });
    }
  });
}

// ============================================================
// 幻灯片6：学习与支持 + 行动建议
// ============================================================
{
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: THEME.bg }, line: { color: THEME.bg } });

  addSectionHeader(slide, "CHAPTER 06", "学习与支持", "LEARN & SUPPORT", 0.3);

  // 左列：学习工具
  const learnTools = [
    { name: "小Q同学", desc: "腾讯学堂AI助手，积分兑礼，课程推荐", link: "https://xiaoq.woa.com" },
    { name: "KM", desc: "内部经验分享平台，含AI效率星期四专题", link: "https://km.woa.com" },
    { name: "企微小助手", desc: "按号搜索内部信息，2000/7000/8000/O2000/小T", link: "https://work.weixin.qq.com" },
  ];

  learnTools.forEach((tool, i) => {
    const x = 0.5, y = 1.0 + i * 1.9, w = 5.8, h = 1.7;
    addCard(slide, x, y, w, h);
    slide.addText(tool.name, {
      x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.35,
      fontSize: 13, bold: true, color: THEME.title, fontFace: "微软雅黑"
    });
    slide.addText(tool.desc, {
      x: x + 0.2, y: y + 0.5, w: w - 0.4, h: 0.9,
      fontSize: 10, color: THEME.desc, fontFace: "微软雅黑", wrap: true
    });
  });

  // 右列：行动建议
  slide.addText("行动建议", {
    x: 6.8, y: 1.0, w: 6.0, h: 0.35,
    fontSize: 14, bold: true, color: THEME.title, fontFace: "微软雅黑"
  });

  const actions = [
    { step: "第一周", title: "选1个工具深度用", detail: "选你岗位最相关的工具，每天用它完成1个真实任务" },
    { step: "第二周", title: "串进工作流", detail: "把AI工具固定到日常工作流中" },
    { step: "第三周", title: "沉淀团队经验", detail: "在iWiki建「AI使用小技巧」页面，团队共享" },
    { step: "第四周", title: "组织AI分享", detail: "发起团队AI工具分享会，建立AI用例库" },
  ];

  actions.forEach((act, i) => {
    const x = 6.8, y = 1.4 + i * 1.35, w = 6.0, h = 1.2;
    addCard(slide, x, y, w, h);
    slide.addText(act.step, {
      x: x + 0.2, y: y + 0.12, w: 1.5, h: 0.28,
      fontSize: 10, bold: true, color: THEME.link, fontFace: "微软雅黑"
    });
    slide.addText(act.title, {
      x: x + 1.8, y: y + 0.12, w: w - 2.0, h: 0.28,
      fontSize: 11, bold: true, color: THEME.title, fontFace: "微软雅黑"
    });
    slide.addText(act.detail, {
      x: x + 0.2, y: y + 0.45, w: w - 0.4, h: 0.65,
      fontSize: 9, color: THEME.desc, fontFace: "微软雅黑", wrap: true
    });
  });

  slide.addText("核心主张：亲手干 · 深度用 AI · 把 Agent 配进每条业务流 · 建立团队 AI 用例库", {
    x: 0.5, y: 7.0, w: W - 1.0, h: 0.3,
    fontSize: 10, color: THEME.subtitle, align: "center", valign: "middle", fontFace: "微软雅黑"
  });
}

// ===== 输出文件 =====
pptx.writeFile({ fileName: "鹅厂AI工具探索地图.pptx" })
  .then(() => console.log("✅ 生成成功：鹅厂AI工具探索地图.pptx"))
  .catch(e => { console.error("❌ 生成失败:", e); process.exit(1); });
