const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, Header, Footer
} = require("docx");

// ===== 常量 =====
const FONT = "微软雅黑";
const GREEN = "14523F";
const PURPLE = "7C3AED";
const PAGE_W = 11906, MARGIN = 1440;
const CONTENT_W = PAGE_W - MARGIN * 2; // 9026

const border = { style: BorderStyle.SINGLE, size: 4, color: "9CA3AF" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// ===== 工具函数 =====
const p = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 360 },
  ...opts.para,
  children: [new TextRun({ text, font: FONT, size: 24, ...opts.run })]
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 280, after: 160 },
  children: [new TextRun({ text, font: FONT, size: 30, bold: true, color: GREEN })]
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 200, after: 120 },
  children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: PURPLE })]
});

const bullet = (text, bold = false) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80, line: 340 },
  children: [new TextRun({ text, font: FONT, size: 24, bold })]
});

const cell = (text, w, opts = {}) => new TableCell({
  borders, width: { size: w, type: WidthType.DXA }, margins: cellMargins,
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  children: [new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [new TextRun({ text, font: FONT, size: 22, bold: !!opts.bold })]
  })]
});

// ===== 表格 1：团队分工 =====
const teamCols = [700, 1600, 1800, 1500, 800, 2626]; // 合计 9026
const teamHeader = ["序号", "姓名", "学院", "专业", "年级", "分工（开发/设计/调研/路演）"];
const teamData = [
  ["1", "刘秋宇（队长）", "人工智能学院", "人工智能相关", "24级", "提示词工程 + 整体项目统筹：和 CodeBuddy 对接生成、迭代 Demo（核心 Vibe Coding 岗位）"],
  ["2", "队员2", "人工智能学院", "待定", "待定", "开发调试：测试网页 Demo、排查运行 bug、修改 AI 生成代码"],
  ["3", "队员3", "人工智能学院", "待定", "待定", "视觉设计：UI 风格、配图、页面美化、配色、海报素材"],
  ["4", "队员4", "人工智能学院", "待定", "待定", "内容调研：校园场景素材收集、关卡剧情撰写、题库文案打磨"],
  ["5", "队员5", "人工智能学院", "待定", "待定", "路演策划：PPT 撰写、Demo 演示脚本、答辩话术、作品介绍"]
];

const teamTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: teamCols,
  rows: [
    new TableRow({
      tableHeader: true,
      children: teamHeader.map((t, i) => cell(t, teamCols[i], { fill: "D9F2E6", bold: true, center: true }))
    }),
    ...teamData.map(row => new TableRow({
      children: row.map((t, i) => cell(t, teamCols[i], { center: i === 0 || i === 4 }))
    }))
  ]
});

// ===== 表格 2：开发计划（对应赛程） =====
const planCols = [2200, 2000, 4826];
const planData = [
  ["阶段", "时间", "我们的任务与产出"],
  ["第一阶段：组队报名", "11.9 - 11.14", "完成组队与报名表填写；产出可运行 Demo（campus-quest.html，已完成）；提交本初步方案"],
  ["第二阶段：训练营", "11月中旬", "全队参加 3 场训练营，重点学习 AI 工具使用与提示词技巧"],
  ["第三阶段：松开发", "11.28", "6-8 小时黑客松开发：扩展至 5+ 关卡、加入音效动画、AI 工具实际使用演示"],
  ["第四阶段：路演展示", "11.29", "完成作品介绍 PPT、Demo 路演脚本，现场演示并答辩"],
  ["第五阶段：成果展", "赛后 1 周", "整理项目复盘材料，配合学术厅大厅成果展示"]
];
const planTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: planCols,
  rows: [
    new TableRow({
      tableHeader: true,
      children: planData[0].map((t, i) => cell(t, planCols[i], { fill: "E4DAFB", bold: true, center: true }))
    }),
    ...planData.slice(1).map(row => new TableRow({
      children: row.map((t, i) => cell(t, planCols[i]))
    }))
  ]
});

// ===== 文档主体 =====
const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { size: 30, bold: true, font: FONT, color: GREEN },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        run: { size: 26, bold: true, font: FONT, color: PURPLE },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } }
    ]
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: 16838 }, // A4
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN, space: 1 } },
        children: [new TextRun({ text: "首届\u201CAI+X\u201D校园 Vibe Coding 创新黑客松 · 初步方案", font: FONT, size: 18, color: "6B7280" })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "第 ", font: FONT, size: 18, color: "6B7280" }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "6B7280" }),
          new TextRun({ text: " 页", font: FONT, size: 18, color: "6B7280" })]
      })] })
    },
    children: [
      // ===== 封面标题 =====
      new Paragraph({ spacing: { before: 2400, after: 240 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "初步方案", font: FONT, size: 56, bold: true, color: GREEN })] }),
      new Paragraph({ spacing: { after: 160 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "上海中侨职业技术大学首届\u201CAI+X\u201D校园 Vibe Coding 创新黑客松", font: FONT, size: 28, color: "374151" })] }),
      new Paragraph({ spacing: { after: 600 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "项目：中侨校园一日闯关 —— 校园主题文字闯关小游戏", font: FONT, size: 26, bold: true, color: PURPLE })] }),
      new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "意向赛道：AI + 校园文娱", font: FONT, size: 24 })] }),
      new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "团队：刘秋宇（队长）及队员共 5 人 · 人工智能学院", font: FONT, size: 24 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "2026 年 11 月 14 日", font: FONT, size: 24 })] }),

      // ===== 一、项目概述 =====
      new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "一、项目概述", font: FONT, size: 30, bold: true, color: GREEN })] }),
      h2("1.1 项目名称"),
      p("《中侨校园一日闯关》—— 校园主题文字闯关小游戏"),
      h2("1.2 一句话创意"),
      p("依托 Vibe Coding，打造上海中侨校园主题文字闯关小游戏，用提示词驱动 AI 快速生成可交互 Demo，带给同学趣味校园互动体验。"),
      h2("1.3 项目亮点"),
      bullet("纯前端单文件实现：一个 HTML 文件内嵌全部代码，双击即玩，零后端、零依赖，演示零风险"),
      bullet("剧情数据驱动：关卡内容与代码逻辑分离，扩展新关卡只需添加数据节点"),
      bullet("校园场景强共鸣：早八赶课、图书馆占座、社团招新等真实校园情景，贴合同学生活"),
      bullet("AI 深度参与：从关卡文案、UI 配色到代码生成全程与 AI（CodeBuddy）协作，体现 AI+X 主题")

      // ===== 二、作品形态与玩法 =====
      , new Paragraph({ heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "二、作品形态与玩法设计", font: FONT, size: 30, bold: true, color: GREEN })] }),
      h2("2.1 玩法概述"),
      p("玩家扮演一名中侨学子，从早晨到傍晚经历校园一天的三个关键场景。每个关卡呈现一段剧情文字与三个选项，玩家点击选项后：剧情跳转、得分增减；错误选择触发\u201C小失败\u201D回到本关重选，保证可通关、不卡死。"),
      h2("2.2 关卡设计"),
      bullet("序章 · 早八危机：起床迟到的抉择，考验时间管理（引出 AI 导论课程）"),
      bullet("第一关 · 课堂问答：随机点名回答 AI 应用场景，融入课程知识点"),
      bullet("第二关 · 图书馆之谜：无人占座物品的处理，考验校园公德"),
      bullet("第三关 · 社团招新夜：AI 社与篮球社的终极选择，呼应黑客松本身"),
      h2("2.3 结算机制"),
      p("通关后按总分给出四档评价：校园传说·学霸 / 优秀中侨人 / 普通中侨人 / 快乐摸鱼人，鼓励重复游玩。"),
      h2("2.4 视觉风格"),
      p("简约校园风：深绿渐变底色 + 紫色高亮描边，贴合本次比赛海报色调；卡片式布局，适配手机与电脑屏幕。")

      // ===== 三、Demo 完成情况 =====
      , new Paragraph({ heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "三、Demo 完成情况（随本方案一并提交）", font: FONT, size: 30, bold: true, color: GREEN })] }),
      bullet("文件名：campus-quest.html（单文件，约 300 行）", true),
      bullet("已完成：序章 + 三大关卡完整剧情分支、实时计分、失败重选机制、胜利结算页、重新开始按钮"),
      bullet("已验证：无报错，浏览器双击即可运行，无需任何安装"),
      bullet("技术栈：原生 HTML + CSS + JavaScript，未使用任何框架")

      // ===== 四、团队分工 =====
      , new Paragraph({ heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "四、团队分工（5 人）", font: FONT, size: 30, bold: true, color: GREEN })] }),
      teamTable,

      // ===== 五、开发计划 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "五、开发计划（对应赛事流程）", font: FONT, size: 30, bold: true, color: GREEN })] }),
      planTable,

      // ===== 六、报名信息确认 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "六、报名信息确认", font: FONT, size: 30, bold: true, color: GREEN })] }),
      bullet("意向赛道：AI + 校园文娱（冲击最佳校园应用奖）"),
      bullet("是否零基础参赛：是（优先助教结对，符合比赛福利）"),
      bullet("是否需要匹配队友：不需要（已 5 人满编）"),
      bullet("AI 工具使用：同意在参赛过程中如实说明 AI 工具使用情况，作品可现场演示", true)
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("c:/Users/moon-/CodeBuddy/20261111101827/初步方案-中侨校园一日闯关.docx", buf);
  console.log("OK");
});
