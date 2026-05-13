import express from 'express';
import cors from 'cors';
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'fs';
import crypto from 'crypto';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3001;
const DATA_FILE = './data.json';
const UPLOAD_DIR = './uploads';
const SESSION_SECRET = crypto.randomBytes(32).toString('hex');

// 确保上传目录存在
if (!existsSync(UPLOAD_DIR)) {
  mkdirSync(UPLOAD_DIR, { recursive: true });
}

// 配置 multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({ 
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.docx', '.doc', '.pdf'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('仅支持 Word 和 PDF 格式'));
    }
  }
});

// 中间件
app.use(cors());
app.use(express.json());
app.use(express.static('../'));
app.use('/uploads', express.static(UPLOAD_DIR));

// Session 存储（内存中，生产环境建议用 Redis）
const sessions = {};

// 读取数据
function loadData() {
  try {
    if (existsSync(DATA_FILE)) {
      const content = readFileSync(DATA_FILE, 'utf-8');
      return JSON.parse(content);
    }
  } catch (e) {
    console.log('Loading default data');
  }
  return getDefaultData();
}

// 保存数据
function saveData(data) {
  data.lastUpdate = new Date().toISOString();
  writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
  return data;
}

// 默认管理员列表
function getDefaultData() {
  return {
    lastUpdate: new Date().toISOString(),
    admins: [
      { id: 'admin1', name: '管理员', email: 'admin@corp.com', role: 'super_admin', org: 'IEG/技术工程部' }
    ],
    sections: [
      {
        id: 'coding',
        name: '编程开发',
        nameEn: 'CODING & BUILD',
        color: 'blue',
        icon: '💻',
        tools: [
          {
            id: 'codebuddy',
            name: 'CodeBuddy',
            description: '一站式 AI 编码助手，含 IDE、插件、CLI 多种形态',
            url: 'https://codebuddy.cn',
            tags: ['IDE', '插件', 'CLI'],
            resources: ['全家桶介绍', 'IDE版', '插件版', 'CLI版'],
            cases: ['医疗团队AI全流程落地', '养出能思考的专家']
          },
          {
            id: 'workbuddy',
            name: 'WorkBuddy',
            description: '腾讯自研研发工作台，代码审查与团队协作',
            url: 'https://workbuddy.woa.com',
            tags: ['代码审查', '协作'],
            resources: [],
            cases: []
          },
          {
            id: 'with',
            name: 'With',
            description: '腾讯内部协作平台，团队沟通与项目管理',
            url: 'https://with.woa.com',
            tags: ['协作', '沟通'],
            resources: [],
            cases: []
          },
          {
            id: 'knot',
            name: 'Knot',
            description: '代码知识图谱工具，代码理解与导航',
            url: 'https://knot.woa.com',
            tags: ['知识图谱', '代码理解'],
            resources: [],
            cases: []
          }
        ]
      },
      {
        id: 'design',
        name: '设计与创意',
        nameEn: 'DESIGN & CREATE',
        color: 'purple',
        icon: '🎨',
        tools: [
          {
            id: 'ardot',
            name: 'Ardot',
            description: '设计资产管理与协作平台',
            url: 'https://ardot.woa.com',
            tags: ['设计资产', '协作'],
            resources: [],
            cases: []
          },
          {
            id: 'sanquan',
            name: '三谦',
            description: '创意设计工具与资源库',
            url: 'https://sanquan.woa.com',
            tags: ['创意', '设计'],
            resources: [],
            cases: []
          },
          {
            id: 'xdream',
            name: 'XDream',
            description: 'AI 驱动的创意设计平台',
            url: 'https://xdream.woa.com',
            tags: ['AI创意', '设计'],
            resources: [],
            cases: []
          },
          {
            id: 'miora',
            name: 'Miora',
            description: '设计灵感与素材平台',
            url: 'https://miora.woa.com',
            tags: ['灵感', '素材'],
            resources: [],
            cases: []
          }
        ]
      },
      {
        id: 'productivity',
        name: '智能办公',
        nameEn: 'INTELLIGENT OFFICE',
        color: 'orange',
        icon: '⚡',
        tools: [
          {
            id: 'hunyuan',
            name: '混元大模型',
            description: '腾讯自研大语言模型，智能问答与内容生成',
            url: 'https://hunyuan.qq.com',
            tags: ['大模型', 'AI助手'],
            resources: [],
            cases: []
          },
          {
            id: 'yuanbao',
            name: '元宝',
            description: '腾讯混元 AI 智能助手 APP',
            url: 'https://yuanbao.qq.com',
            tags: ['智能助手', '移动端'],
            resources: [],
            cases: []
          },
          {
            id: 'openclaw',
            name: 'OpenClaw',
            description: 'AI 能力开放平台',
            url: 'https://openclaw.woa.com',
            tags: ['AI开放', 'API'],
            resources: [],
            cases: []
          },
          {
            id: 'iwiki',
            name: 'iWiki',
            description: '企业知识管理与搜索平台',
            url: 'https://iwiki.woa.com',
            tags: ['知识管理', '搜索'],
            resources: [],
            cases: []
          }
        ]
      },
      {
        id: 'gaming',
        name: '游戏生产',
        nameEn: 'GAME PRODUCTION',
        color: 'red',
        icon: '🎮',
        tools: [
          {
            id: 'visvise',
            name: 'VISVISE',
            description: '游戏可视化设计与制作工具',
            url: 'https://visvise.woa.com',
            tags: ['可视化', '游戏制作'],
            resources: [],
            cases: []
          },
          {
            id: 'giinex',
            name: 'GiiNEX',
            description: 'AI驱动的游戏生产平台',
            url: 'https://giinex.woa.com',
            tags: ['AI游戏', '生产'],
            resources: [],
            cases: []
          },
          {
            id: 'hunyuan-game',
            name: '混元游戏',
            description: '游戏行业AI解决方案',
            url: 'https://hunyuan.qq.com/game',
            tags: ['AI游戏', '解决方案'],
            resources: [],
            cases: []
          },
          {
            id: '3dcity',
            name: '3D城市',
            description: '3D城市场景编辑器与素材库',
            url: 'https://3dcity.woa.com',
            tags: ['3D', '素材'],
            resources: [],
            cases: []
          }
        ]
      }
    ],
    experiences: [
      {
        id: 'exp1',
        title: 'CodeBuddy 提效心得',
        author: '工程师小王',
        content: '使用 CodeBuddy 后，代码审查效率提升了 50%，特别是在处理重复性高的代码时效果显著。强烈推荐搭配 IDE 插件使用，体验更佳。',
        date: '2026-05-10'
      }
    ],
    stars: [
      {
        id: 'star1',
        name: '张明',
        title: 'AI工具探索先锋',
        avatarUrl: '',
        description: '累计分享 12 篇 AI 工具使用心得，帮助 50+ 同事提升工作效率。精通 CodeBuddy、GiiNEX 等工具，擅长解决各类技术难题。',
        badges: ['工具达人', '分享之星'],
        date: '2026-05-01'
      },
      {
        id: 'star2',
        name: '李华',
        title: '智能办公先锋',
        avatarUrl: '',
        description: '在部门内推广 AI 工具落地，带领团队效率提升 40%。善于挖掘工具潜力，总结出一套实用的 AI 提效方法论。',
        badges: ['团队 leader', '最佳实践'],
        date: '2026-05-05'
      },
      {
        id: 'star3',
        name: '王芳',
        title: '创意设计达人',
        avatarUrl: '',
        description: '深度使用设计类 AI 工具，在多个项目中成功应用 AI 辅助设计，缩短设计周期 60%，获得团队一致好评。',
        badges: ['设计专家', '创新先锋'],
        date: '2026-05-08'
      }
    ]
  };
}

// 生成 session token
function generateToken() {
  return crypto.randomBytes(24).toString('hex');
}

// 验证管理员权限
function verifyAdmin(token) {
  if (!token) return null;
  const session = sessions[token];
  if (!session) return null;
  if (Date.now() - session.loginTime > 24 * 60 * 60 * 1000) {
    delete sessions[token];
    return null;
  }
  return session.admin;
}

// ========== 认证 API ==========

// 获取当前管理员信息
app.get('/api/admin/me', (req, res) => {
  const token = req.headers['x-session-token'];
  const admin = verifyAdmin(token);
  if (admin) {
    res.json({ loggedIn: true, admin });
  } else {
    res.json({ loggedIn: false });
  }
});

// 管理员登录（基于邮箱/工号验证）
app.post('/api/admin/login', (req, res) => {
  const { email, staffId, org } = req.body;
  
  if (!email && !staffId) {
    return res.status(400).json({ error: '请提供邮箱或工号' });
  }
  
  const data = loadData();
  
  // 检查是否为注册管理员
  let admin = data.admins?.find(a => 
    a.email === email || 
    (staffId && a.staffId === staffId)
  );
  
  // 如果不是预设管理员，但提供了组织架构信息，可以作为普通管理员注册
  if (!admin && email && org) {
    admin = {
      id: 'admin_' + Date.now(),
      email,
      staffId,
      org,
      role: 'admin',
      name: email.split('@')[0]
    };
    
    data.admins = data.admins || [];
    data.admins.push(admin);
    saveData(data);
  }
  
  if (!admin) {
    return res.status(403).json({ 
      error: '您没有管理权限',
      hint: '请联系管理员添加您的账号到白名单'
    });
  }
  
  // 创建 session
  const token = generateToken();
  sessions[token] = {
    admin,
    loginTime: Date.now()
  };
  
  res.json({ 
    success: true, 
    token,
    admin: {
      id: admin.id,
      name: admin.name,
      email: admin.email,
      role: admin.role,
      org: admin.org
    }
  });
});

// 管理员登出
app.post('/api/admin/logout', (req, res) => {
  const token = req.headers['x-session-token'];
  if (token && sessions[token]) {
    delete sessions[token];
  }
  res.json({ success: true });
});

// 获取管理员列表（仅 super_admin）
app.get('/api/admin/list', (req, res) => {
  const token = req.headers['x-session-token'];
  const admin = verifyAdmin(token);
  
  if (!admin || admin.role !== 'super_admin') {
    return res.status(403).json({ error: '需要超级管理员权限' });
  }
  
  const data = loadData();
  res.json(data.admins || []);
});

// 添加管理员（仅 super_admin）
app.post('/api/admin', (req, res) => {
  const token = req.headers['x-session-token'];
  const admin = verifyAdmin(token);
  
  if (!admin || admin.role !== 'super_admin') {
    return res.status(403).json({ error: '需要超级管理员权限' });
  }
  
  const data = loadData();
  const newAdmin = req.body;
  
  if (!newAdmin.email || !newAdmin.name) {
    return res.status(400).json({ error: '缺少必要字段' });
  }
  
  if (data.admins?.find(a => a.email === newAdmin.email)) {
    return res.status(400).json({ error: '管理员已存在' });
  }
  
  newAdmin.id = 'admin_' + Date.now();
  newAdmin.role = newAdmin.role || 'admin';
  data.admins = data.admins || [];
  data.admins.push(newAdmin);
  saveData(data);
  
  res.json(newAdmin);
});

// 删除管理员（仅 super_admin）
app.delete('/api/admin/:id', (req, res) => {
  const token = req.headers['x-session-token'];
  const admin = verifyAdmin(token);
  
  if (!admin || admin.role !== 'super_admin') {
    return res.status(403).json({ error: '需要超级管理员权限' });
  }
  
  const data = loadData();
  const idx = data.admins.findIndex(a => a.id === req.params.id);
  
  if (idx === -1) {
    return res.status(404).json({ error: '管理员不存在' });
  }
  
  // 防止删除自己
  if (data.admins[idx].email === admin.email) {
    return res.status(400).json({ error: '不能删除自己' });
  }
  
  data.admins.splice(idx, 1);
  saveData(data);
  res.json({ success: true });
});

// ========== 数据 API ==========

// 获取全部数据
app.get('/api/data', (req, res) => {
  const data = loadData();
  // 不返回完整的 admin 列表
  const publicData = { ...data };
  delete publicData.admins;
  res.json(publicData);
});

// ========== 板块管理 ==========

// 添加板块（需管理员权限）
app.post('/api/sections', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const section = req.body;
  
  if (!section.id || !section.name) {
    return res.status(400).json({ error: '缺少必要字段' });
  }
  
  if (data.sections.find(s => s.id === section.id)) {
    return res.status(400).json({ error: '板块ID已存在' });
  }
  
  section.tools = section.tools || [];
  data.sections.push(section);
  saveData(data);
  res.json(section);
});

// 更新板块
app.put('/api/sections/:id', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const idx = data.sections.findIndex(s => s.id === req.params.id);
  
  if (idx === -1) {
    return res.status(404).json({ error: '板块不存在' });
  }
  
  data.sections[idx] = { ...data.sections[idx], ...req.body };
  saveData(data);
  res.json(data.sections[idx]);
});

// 删除板块
app.delete('/api/sections/:id', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const idx = data.sections.findIndex(s => s.id === req.params.id);
  
  if (idx === -1) {
    return res.status(404).json({ error: '板块不存在' });
  }
  
  data.sections.splice(idx, 1);
  saveData(data);
  res.json({ success: true });
});

// ========== 工具管理 ==========

// 添加工具
app.post('/api/sections/:sectionId/tools', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const section = data.sections.find(s => s.id === req.params.sectionId);
  
  if (!section) {
    return res.status(404).json({ error: '板块不存在' });
  }
  
  const tool = { ...req.body, id: req.body.id || 'tool' + Date.now() };
  section.tools = section.tools || [];
  section.tools.push(tool);
  saveData(data);
  res.json(tool);
});

// 更新工具
app.put('/api/sections/:sectionId/tools/:toolId', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const section = data.sections.find(s => s.id === req.params.sectionId);
  
  if (!section) {
    return res.status(404).json({ error: '板块不存在' });
  }
  
  const toolIdx = section.tools.findIndex(t => t.id === req.params.toolId);
  if (toolIdx === -1) {
    return res.status(404).json({ error: '工具不存在' });
  }
  
  section.tools[toolIdx] = { ...section.tools[toolIdx], ...req.body };
  saveData(data);
  res.json(section.tools[toolIdx]);
});

// 删除工具
app.delete('/api/sections/:sectionId/tools/:toolId', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const section = data.sections.find(s => s.id === req.params.sectionId);
  
  if (!section) {
    return res.status(404).json({ error: '板块不存在' });
  }
  
  const toolIdx = section.tools.findIndex(t => t.id === req.params.toolId);
  if (toolIdx === -1) {
    return res.status(404).json({ error: '工具不存在' });
  }
  
  section.tools.splice(toolIdx, 1);
  saveData(data);
  res.json({ success: true });
});

// ========== 文件上传 ==========

// 上传文件
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: '请选择要上传的文件' });
  }
  
  const fileUrl = `/uploads/${req.file.filename}`;
  res.json({
    success: true,
    url: fileUrl,
    originalName: req.file.originalname,
    size: req.file.size
  });
});

// 上传错误处理
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: '文件大小不能超过 10MB' });
    }
    return res.status(400).json({ error: err.message });
  }
  if (err) {
    return res.status(400).json({ error: err.message });
  }
  next();
});

// ========== 经验分享管理 ==========

// 添加经验（公开，无需登录）
app.post('/api/experiences', (req, res) => {
  const data = loadData();
  const exp = { ...req.body, id: req.body.id || 'exp' + Date.now() };
  data.experiences = data.experiences || [];
  data.experiences.unshift(exp);
  saveData(data);
  res.json(exp);
});

// 更新经验
app.put('/api/experiences/:id', (req, res) => {
  const token = req.headers['x-session-token'];
  const admin = verifyAdmin(token);
  
  if (!admin) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const idx = data.experiences.findIndex(e => e.id === req.params.id);
  
  if (idx === -1) {
    return res.status(404).json({ error: '经验不存在' });
  }
  
  data.experiences[idx] = { ...data.experiences[idx], ...req.body };
  saveData(data);
  res.json(data.experiences[idx]);
});

// 删除经验
app.delete('/api/experiences/:id', (req, res) => {
  const token = req.headers['x-session-token'];
  if (!verifyAdmin(token)) {
    return res.status(401).json({ error: '请先登录' });
  }
  
  const data = loadData();
  const idx = data.experiences.findIndex(e => e.id === req.params.id);
  
  if (idx === -1) {
    return res.status(404).json({ error: '经验不存在' });
  }
  
  data.experiences.splice(idx, 1);
  saveData(data);
  res.json({ success: true });
});

// 启动服务器
app.listen(PORT, '0.0.0.0', () => {
  console.log(`🚀 AI 工具地图服务器已启动: http://localhost:${PORT}`);
  console.log(`📝 管理后台: http://localhost:${PORT}/admin.html`);
  console.log(`🗺️  工具地图: http://localhost:${PORT}/index.html`);
  console.log(`🔐 认证模式: 组织架构权限控制已启用`);
  console.log(`🌐 局域网访问: http://<你的IP>:${PORT}`);
});