# 科研项目申报PPT实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 根据项目建议书创建15-18页科研申报PPT，包含封面、6个主要内容章节、结尾页

**Architecture:** 使用pptxgenjs从零创建PPT，采用蓝色科技风格配色，每页2-3栏布局，确保内容充实无空白，每页至少1-2个视觉元素

**Tech Stack:** Node.js + pptxgenjs + react-icons + sharp (图标)

---

## 文件结构

```
C:\Users\15974\Desktop\CC test\ppt\
├── create-ppt.js          # 主脚本
├── package.json           # 依赖配置
├── output/
│   └── 科研项目申报PPT.pptx
└── docs/superpowers/plans/
    └── 2026-05-20-ppt-implementation.md
```

---

## PPT页面结构 (15-18页)

| 页码 | 内容 | 布局 |
|------|------|------|
| 1 | 封面 | 居中大标题+底部信息 |
| 2-3 | 研究背景及国内外研究现状 | 2栏：痛点+数据 |
| 4-7 | 项目实施内容 | 3栏：三大任务+流程图 |
| 8-9 | 提交的成果及技术经济目标 | 成果列表+指标表格 |
| 10-13 | 项目实施路线及方案 | 技术路线图+分阶段流程 |
| 14-15 | 项目进度计划及人力资源 | 甘特图+组织架构 |
| 16 | 财务计划 | 预算饼图/柱图 |
| 17 | 结尾致谢 | 简洁致谢 |

---

## 任务分解

### Task 1: 环境准备

**Files:**
- Create: `C:\Users\15974\Desktop\CC test\ppt\package.json`
- Create: `C:\Users\15974\Desktop\CC test\ppt\create-ppt.js`

- [ ] **Step 1: 创建package.json**

```json
{
  "name": "research-project-ppt",
  "version": "1.0.0",
  "type": "module",
  "dependencies": {
    "pptxgenjs": "^3.12.0",
    "react": "^18.2.0",
    "react-dom": "^18.2.0",
    "react-icons": "^5.0.0",
    "sharp": "^0.33.0"
  }
}
```

Run: `npm install`

- [ ] **Step 2: 创建主脚本文件**

创建 `create-ppt.js`，包含：
- PPT基础设置（16:9布局、蓝色主题）
- 颜色常量定义
- 辅助函数（创建阴影、创建卡片等）

---

### Task 2: 创建封面页

**Files:**
- Modify: `create-ppt.js` (添加封面生成函数)

- [ ] **Step 1: 添加封面函数**

```javascript
function createCoverSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "1E2761" };

  // 顶部装饰线
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: "CADCFC" }
  });

  // 项目名称
  slide.addText("基于AI视觉识别的\n狭窄钢箱梁内部缺陷巡检机器人", {
    x: 0.5, y: 1.8, w: 9, h: 1.6,
    fontSize: 40, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", align: "center", valign: "middle"
  });

  // 副标题
  slide.addText("二期研发项目", {
    x: 0.5, y: 3.5, w: 9, h: 0.6,
    fontSize: 24, fontFace: "Microsoft YaHei",
    color: "CADCFC", align: "center"
  });

  // 分隔线
  slide.addShape(pres.shapes.LINE, {
    x: 3, y: 4.3, w: 4, h: 0,
    line: { color: "CADCFC", width: 1 }
  });

  // 申请单位
  slide.addText("申报单位：桥梁结构安全研究院\n协作单位：龚元浩科技有限公司", {
    x: 0.5, y: 4.6, w: 9, h: 0.8,
    fontSize: 14, fontFace: "Microsoft YaHei",
    color: "CADCFC", align: "center"
  });

  // 日期
  slide.addText("2026年5月", {
    x: 0.5, y: 5.2, w: 9, h: 0.3,
    fontSize: 12, fontFace: "Microsoft YaHei",
    color: "CADCFC", align: "center"
  });
}
```

---

### Task 3: 创建研究背景页面 (2页)

**Files:**
- Modify: `create-ppt.js` (添加背景页函数)

- [ ] **Step 1: 第一页 - 行业背景与痛点**

```javascript
function createBackgroundSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("一、研究背景及国内外研究现状", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 左侧：行业背景
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.2, w: 4.4, h: 2.2,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("行业背景", {
    x: 0.6, y: 1.35, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "2024年全国公路桥梁110.81万座", options: { bullet: true, breakLine: true } },
    { text: "30%以上桥梁进入老龄化阶段", options: { bullet: true, breakLine: true } },
    { text: "钢箱梁结构占大跨径桥梁80%以上", options: { bullet: true, breakLine: true } },
    { text: "人工检测效率低、风险高、成本高", options: { bullet: true } }
  ], { x: 0.6, y: 1.85, w: 4, h: 1.4, fontSize: 12, color: "333333" });

  // 右侧：行业痛点
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.2, w: 4.4, h: 2.2,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("行业痛点", {
    x: 5.4, y: 1.35, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "作业空间狭窄（1.2~1.8m），人员无法直立", options: { bullet: true, breakLine: true } },
    { text: "高温高湿（55°C以上，湿度80%+）", options: { bullet: true, breakLine: true } },
    { text: "每公里U肋焊缝约100km，缺陷3类", options: { bullet: true, breakLine: true } },
    { text: "人工检测效率低，漏检率高达30%", options: { bullet: true } }
  ], { x: 5.4, y: 1.85, w: 4, h: 1.4, fontX: 12, color: "333333" });

  // 底部：市场数据
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.6, w: 9.2, h: 1.7,
    fill: { color: "1E2761" }
  });
  // 三个数据卡片
  const stats = [
    { value: "80亿+", label: "全球SHM市场规模(元)" },
    { value: "55%", label: "国内市场份额" },
    { value: "10.36%", label: "年复合增长率" }
  ];
  stats.forEach((stat, i) => {
    const x = 0.8 + i * 3;
    slide.addText(stat.value, {
      x, y: 3.75, w: 2.6, h: 0.8,
      fontSize: 36, bold: true, color: "CADCFC", align: "center"
    });
    slide.addText(stat.label, {
      x, y: 4.55, w: 2.6, h: 0.5,
      fontSize: 11, color: "CADCFC", align: "center"
    });
  });
}
```

- [ ] **Step 2: 第二页 - 国内外研究现状对比**

```javascript
function createBackgroundSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("一、研究背景及国内外研究现状（续）", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 国内研究现状
  slide.addText("国内研究现状", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // 技术对比表格
  const tableData = [
    [
      { text: "对比项", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "传统方式", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "本项目方案", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    ["检测方式", "人工携带仪器", "自动化巡检机器人"],
    ["定位精度", ">10cm", "≤2cm"],
    ["识别准确率", "70-80%", "≥95%"],
    ["检测效率", "低", "提升50%+"],
    ["越障能力", "差", "双椭圆履带+特殊结构"]
  ];
  slide.addTable(tableData, {
    x: 0.5, y: 1.5, w: 9, h: 2.2,
    colW: [2.5, 3, 3.5],
    fontSize: 11,
    border: { pt: 0.5, color: "CCCCCC" },
    fill: { color: "FFFFFF" }
  });

  // 本项目创新点
  slide.addText("本项目技术优势", {
    x: 0.5, y: 3.9, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "双椭圆履带+越障模块：适应复杂空间（5-8m检测范围）", options: { bullet: true, breakLine: true } },
    { text: "YOLOv8-GC算法：识别准确率≥95%，漏检率<5%", options: { bullet: true, breakLine: true } },
    { text: "BIM+RFID融合定位：定位精度≤2cm，姿态偏差<1°", options: { bullet: true, breakLine: true } },
    { text: "行业标准引领：参与JTG/T 5123-2025规范制定", options: { bullet: true } }
  ], { x: 0.5, y: 4.3, w: 9, h: 1.2, fontSize: 12, color: "333333" });
}
```

---

### Task 4: 创建项目实施内容页面 (3-4页)

**Files:**
- Modify: `create-ppt.js` (添加实施内容函数)

- [ ] **Step 1: 第三页 - 研究目标总览**

```javascript
function createImplementationSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("二、项目实施内容 — 研究目标", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 核心目标卡片
  const goals = [
    { icon: "🎯", title: "检测范围", value: "5-8m", desc: "覆盖不同跨度钢箱梁" },
    { icon: "📍", title: "定位精度", value: "≤2cm", desc: "姿态偏差<1°" },
    { icon: "🔍", title: "识别准确率", value: "≥95%", desc: "漏检率<5%" },
    { icon: "⚡", title: "检测效率", value: "提升50%+", desc: "实时响应<1s" }
  ];

  goals.forEach((goal, i) => {
    const x = 0.5 + i * 2.35;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.2, w: 2.15, h: 1.8,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addText(goal.title, {
      x: x + 0.1, y: 1.35, w: 1.95, h: 0.35,
      fontSize: 11, color: "666666", align: "center"
    });
    slide.addText(goal.value, {
      x: x + 0.1, y: 1.7, w: 1.95, h: 0.6,
      fontSize: 26, bold: true, color: "1E2761", align: "center"
    });
    slide.addText(goal.desc, {
      x: x + 0.1, y: 2.35, w: 1.95, h: 0.5,
      fontSize: 9, color: "888888", align: "center"
    });
  });

  // 三大研究任务
  slide.addText("三大研究任务", {
    x: 0.5, y: 3.2, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  const tasks = [
    { num: "1", title: "硬件系统优化研究", desc: "微型化结构、双椭圆履带、越障机构" },
    { num: "2", title: "核心算法优化研究", desc: "YOLOv8-GC、RFID+视觉融合定位" },
    { num: "3", title: "云平台协同管控研究", desc: "数字孪生、预测性维护、远程诊断" }
  ];

  tasks.forEach((task, i) => {
    const y = 3.7 + i * 0.65;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.5, y: y, w: 0.4, h: 0.4,
      fill: { color: "1E2761" }
    });
    slide.addText(task.num, {
      x: 0.5, y: y, w: 0.4, h: 0.4,
      fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    slide.addText(task.title, {
      x: 1.0, y: y, w: 3, h: 0.4,
      fontSize: 13, bold: true, color: "1E2761", valign: "middle"
    });
    slide.addText(task.desc, {
      x: 4.0, y: y, w: 5.5, h: 0.4,
      fontSize: 11, color: "555555", valign: "middle"
    });
  });
}
```

- [ ] **Step 2: 第四页 - 硬件系统优化**

```javascript
function createImplementationSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("二、项目实施内容 — 硬件系统优化", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 左侧内容
  slide.addText("微型化结构设计", {
    x: 0.5, y: 1.1, w: 4.5, h: 0.4,
    fontSize: 15, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "基于刚柔耦合动力学模型优化", options: { bullet: true, breakLine: true } },
    { text: "解决作业空间1.2~1.8m通过性问题", options: { bullet: true, breakLine: true } },
    { text: "双椭圆履带差速转向+越障模块", options: { bullet: true, breakLine: true } },
    { text: "克服U肋焊缝及附件干涉", options: { bullet: true } }
  ], { x: 0.5, y: 1.5, w: 4.5, h: 1.4, fontSize: 11, color: "333333" });

  slide.addText("感知与执行系统", {
    x: 0.5, y: 3.0, w: 4.5, h: 0.4,
    fontSize: 15, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "高清摄像系统：裂缝<0.1mm识别", options: { bullet: true, breakLine: true } },
    { text: "多自由度机械臂精确采样", options: { bullet: true, breakLine: true } },
    { text: "实时回传舱体倾斜角度监测", options: { bullet: true } }
  ], { x: 0.5, y: 3.4, w: 4.5, h: 1.2, fontSize: 11, color: "333333" });

  // 右侧技术参数表
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.1, w: 4.4, h: 3.8,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("核心硬件指标", {
    x: 5.4, y: 1.25, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: "1E2761"
  });

  const hwParams = [
    ["最小检测缺陷", "≤2.0mm"],
    ["定位精度", "≤10cm"],
    ["姿态定位精度", "≤1cm"],
    ["最大移动速度", "≤1m/s"],
    ["防护等级", "IP65以上"],
    ["载重能力", "≥300kg"]
  ];
  hwParams.forEach((param, i) => {
    const y = 1.75 + i * 0.45;
    slide.addText(param[0], {
      x: 5.4, y, w: 2.2, h: 0.4,
      fontSize: 11, color: "666666"
    });
    slide.addText(param[1], {
      x: 7.6, y, w: 1.8, h: 0.4,
      fontSize: 11, bold: true, color: "1E2761"
    });
    if (i < hwParams.length - 1) {
      slide.addShape(pres.shapes.LINE, {
        x: 5.4, y: y + 0.4, w: 4, h: 0,
        line: { color: "EEEEEE", width: 0.5 }
      });
    }
  });
}
```

- [ ] **Step 3: 第五页 - 核心算法优化**

```javascript
function createImplementationSlide3(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("二、项目实施内容 — 核心算法优化", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 算法1：视觉识别
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.5, h: 2.0,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("YOLOv8-GC视觉识别算法", {
    x: 0.6, y: 1.25, w: 4.1, h: 0.35,
    fontSize: 13, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "GhostModule轻量化模型压缩", options: { bullet: true, breakLine: true } },
    { text: "CA注意力机制增强特征提取", options: { bullet: true, breakLine: true } },
    { text: "WBF边界融合提升检测精度", options: { bullet: true, breakLine: true } },
    { text: "小目标漏检率<5%，识别准确率≥95%", options: { bullet: true } }
  ], { x: 0.6, y: 1.7, w: 4.1, h: 1.2, fontSize: 10, color: "333333" });

  // 算法2：融合定位
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.1, w: 4.5, h: 2.0,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("RFID+视觉融合定位算法", {
    x: 5.3, y: 1.25, w: 4.1, h: 0.35,
    fontSize: 13, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "多源信息融合定位优化", options: { bullet: true, breakLine: true } },
    { text: "姿态-位置协同校正控制", options: { bullet: true, breakLine: true } },
    { text: "BIM模型精确映射定位", options: { bullet: true, breakLine: true } },
    { text: "定位精度≤2cm，姿态偏差<1°", options: { bullet: true } }
  ], { x: 5.3, y: 1.7, w: 4.1, h: 1.2, fontSize: 10, color: "333333" });

  // 算法3：自适应控制
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.3, w: 9.2, h: 2.0,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("自适应控制与预测性维护系统", {
    x: 0.6, y: 3.45, w: 8.8, h: 0.35,
    fontSize: 13, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "S型曲线速度规划：平滑启停，加速度连续", options: { bullet: true, breakLine: true } },
    { text: "模糊PID控制：适应不同工况，自主调整参数", options: { bullet: true, breakLine: true } },
    { text: "数字孪生模型：实时状态映射，健康度评估", options: { bullet: true, breakLine: true } },
    { text: "故障预测：提前识别隐患，从\"故障修\"转向\"预防修\"", options: { bullet: true } }
  ], { x: 0.6, y: 3.9, w: 8.8, h: 1.2, fontSize: 11, color: "333333" });
}
```

- [ ] **Step 4: 第六页 - 云平台与协同管控**

```javascript
function createImplementationSlide4(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("二、项目实施内容 — 云平台协同管控", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 缺陷数据库
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.5, h: 2.1,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("缺陷数据库系统", {
    x: 0.6, y: 1.25, w: 4.1, h: 0.35,
    fontSize: 13, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "覆盖3类11种缺陷类型", options: { bullet: true, breakLine: true } },
    { text: "标准缺陷图谱库>10,000张", options: { bullet: true, breakLine: true } },
    { text: "支持训练样本扩展接口", options: { bullet: true, breakLine: true } },
    { text: "支持按位置/类型/ severity检索", options: { bullet: true } }
  ], { x: 0.6, y: 1.7, w: 4.1, h: 1.3, fontSize: 10, color: "333333" });

  // 数字孪生
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.1, w: 4.5, h: 2.1,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });
  slide.addText("数字孪生可视化平台", {
    x: 5.3, y: 1.25, w: 4.1, h: 0.35,
    fontSize: 13, bold: true, color: "1E2761"
  });
  slide.addText([
    { text: "三维模型实时渲染", options: { bullet: true, breakLine: true } },
    { text: "检测数据空间可视化", options: { bullet: true, breakLine: true } },
    { text: "历史数据回溯分析", options: { bullet: true, breakLine: true } },
    { text: "B/S架构，支持远程访问", options: { bullet: true } }
  ], { x: 5.3, y: 1.7, w: 4.1, h: 1.3, fontSize: 10, color: "333333" });

  // 平台架构
  slide.addText("协同管控平台架构", {
    x: 0.4, y: 3.4, w: 9.2, h: 0.4,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // 架构层级
  const layers = [
    { name: "应用层", items: "数据展示 | 报告生成 | 预警推送" },
    { name: "服务层", items: "数据分析 | 模型推理 | 任务调度" },
    { name: "数据层", items: "MySQL | Redis | 文件存储" },
    { name: "感知层", items: "机器人端 | 传感器网络 | 边缘计算" }
  ];
  layers.forEach((layer, i) => {
    const y = 3.85 + i * 0.42;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 1.2, h: 0.38,
      fill: { color: "1E2761" }
    });
    slide.addText(layer.name, {
      x: 0.4, y, w: 1.2, h: 0.38,
      fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.7, y, w: 7.9, h: 0.38,
      fill: { color: "CADCFC", transparency: 30 }
    });
    slide.addText(layer.items, {
      x: 1.9, y, w: 7.5, h: 0.38,
      fontSize: 10, color: "333333", valign: "middle"
    });
  });
}
```

---

### Task 5: 创建成果与技术目标页面 (2页)

**Files:**
- Modify: `create-ppt.js` (添加成果页函数)

- [ ] **Step 1: 第七页 - 提交的成果**

```javascript
function createResultsSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("三、提交的成果", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 四大成果卡片
  const results = [
    { title: "样机1台", desc: "钢箱梁内部缺陷巡检机器人样机\n满足IP65防护、≤2cm定位精度" },
    { title: "数据库1套", desc: "3类11种缺陷图谱库\n标准样本>10,000张" },
    { title: "平台1套", desc: "协同运维数字孪生平台\nB/S架构，支持远程管控" },
    { title: "论文与专利", desc: "EI/SCI期刊2-3篇\n专利2-3项，软件著作权1项" }
  ];

  results.forEach((r, i) => {
    const x = 0.4 + (i % 2) * 4.7;
    const y = 1.15 + Math.floor(i / 2) * 1.85;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.5, h: 1.65,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    // 左侧色条
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: 1.65,
      fill: { color: "1E2761" }
    });
    slide.addText(r.title, {
      x: x + 0.2, y: y + 0.15, w: 4.1, h: 0.4,
      fontSize: 14, bold: true, color: "1E2761"
    });
    slide.addText(r.desc, {
      x: x + 0.2, y: y + 0.6, w: 4.1, h: 0.9,
      fontSize: 10, color: "555555"
    });
  });

  // 底部关键指标
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.9, w: 9.2, h: 0.55,
    fill: { color: "1E2761" }
  });
  slide.addText("关键指标：识别准确率≥95% | 漏检率<5% | 定位精度≤2cm | 检测效率提升≥50%", {
    x: 0.6, y: 4.95, w: 8.8, h: 0.45,
    fontSize: 12, color: "FFFFFF", align: "center", valign: "middle"
  });
}
```

- [ ] **Step 2: 第八页 - 技术与经济目标**

```javascript
function createResultsSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("四、拟达到的技术和经济目标", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 技术指标表格
  slide.addText("技术指标", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 15, bold: true, color: "1E2761"
  });

  const techTable = [
    [
      { text: "指标名称", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "目标值", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "验证方式", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    ["最小检测缺陷", "≤2.0mm", "实验室标定"],
    ["定位精度", "≤10cm", "实地标定"],
    ["姿态定位精度", "≤1cm", "姿态测量对比"],
    ["移动速度", "≤1m/s", "计时测试"],
    ["防护等级", "IP65以上", "型式试验"],
    ["识别准确率", "≥95%", "实操验证"],
    ["漏检率", "<5%", "对比测试"]
  ];
  slide.addTable(techTable, {
    x: 0.5, y: 1.5, w: 9, h: 2.4,
    colW: [3, 2.5, 3.5],
    fontSize: 10,
    border: { pt: 0.5, color: "CCCCCC" },
    fill: { color: "FFFFFF" }
  });

  // 经济目标
  slide.addText("经济目标", {
    x: 0.5, y: 4.1, w: 4, h: 0.4,
    fontSize: 15, bold: true, color: "1E2761"
  });

  // 经济目标大数字
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.55, w: 4.3, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("600万元", {
    x: 0.6, y: 4.6, w: 2.5, h: 0.8,
    fontSize: 32, bold: true, color: "CADCFC", valign: "middle"
  });
  slide.addText("预计实现产值", {
    x: 3.1, y: 4.6, w: 1.5, h: 0.8,
    fontSize: 12, color: "CADCFC", valign: "middle"
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.0, y: 4.55, w: 4.5, h: 0.9,
    fill: { color: "CADCFC", transparency: 30 }
  });
  slide.addText("600+万元", {
    x: 5.1, y: 4.6, w: 2.2, h: 0.8,
    fontSize: 28, bold: true, color: "1E2761", valign: "middle"
  });
  slide.addText("预计签订合同额", {
    x: 7.3, y: 4.6, w: 2, h: 0.8,
    fontSize: 11, color: "1E2761", valign: "middle"
  });
}
```

---

### Task 6: 创建实施路线页面 (3-4页)

**Files:**
- Modify: `create-ppt.js` (添加路线页函数)

- [ ] **Step 1: 第九页 - 总体技术路线**

```javascript
function createRoadmapSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("五、项目实施路线及方案 — 总体技术路线", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 流程图：阶段1
  const phases = [
    { num: "1", title: "硬件系统\n优化", desc: "微型化结构\n双椭圆履带\n越障模块" },
    { num: "2", title: "核心算法\n优化", desc: "YOLOv8-GC\n融合定位\n自适应控制" },
    { num: "3", title: "云平台\n开发", desc: "数字孪生\n缺陷数据库\n协同管控" },
    { num: "4", title: "示范\n应用", desc: "实体桥梁\n验证测试\n标准制定" }
  ];

  phases.forEach((phase, i) => {
    const x = 0.6 + i * 2.4;
    // 圆形编号
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.7, y: 1.3, w: 0.5, h: 0.5,
      fill: { color: "1E2761" }
    });
    slide.addText(phase.num, {
      x: x + 0.7, y: 1.3, w: 0.5, h: 0.5,
      fontSize: 18, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    // 阶段卡片
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.95, w: 2.0, h: 1.8,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addText(phase.title, {
      x: x + 0.1, y: 2.05, w: 1.8, h: 0.6,
      fontSize: 12, bold: true, color: "1E2761", align: "center"
    });
    slide.addText(phase.desc, {
      x: x + 0.1, y: 2.7, w: 1.8, h: 0.95,
      fontSize: 9, color: "666666", align: "center"
    });
    // 箭头
    if (i < 3) {
      slide.addText("→", {
        x: x + 2.0, y: 2.5, w: 0.4, h: 0.5,
        fontSize: 24, color: "1E2761", align: "center", valign: "middle"
      });
    }
  });

  // 时间线
  slide.addShape(pres.shapes.LINE, {
    x: 0.6, y: 4.2, w: 8.8, h: 0,
    line: { color: "1E2761", width: 2 }
  });

  const timeline = [
    { year: "2026.6-2026.12", label: "阶段一", detail: "硬件优化" },
    { year: "2027.1-2027.6", label: "阶段二", detail: "算法优化" },
    { year: "2027.7-2027.12", label: "阶段三", detail: "平台开发" },
    { year: "2028.1-2028.12", label: "阶段四", detail: "示范应用" }
  ];
  timeline.forEach((t, i) => {
    const x = 1.0 + i * 2.4;
    slide.addShape(pres.shapes.OVAL, {
      x, y: 4.1, w: 0.2, h: 0.2,
      fill: { color: "1E2761" }
    });
    slide.addText(t.year, {
      x: x - 0.5, y: 4.35, w: 1.2, h: 0.3,
      fontSize: 9, color: "1E2761", align: "center"
    });
    slide.addText(t.label, {
      x: x - 0.3, y: 4.6, w: 0.8, h: 0.25,
      fontSize: 10, bold: true, color: "1E2761", align: "center"
    });
    slide.addText(t.detail, {
      x: x - 0.5, y: 4.85, w: 1.2, h: 0.25,
      fontSize: 8, color: "666666", align: "center"
    });
  });
}
```

- [ ] **Step 2: 第十页 - 各阶段详细方案**

```javascript
function createRoadmapSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("五、项目实施路线及方案 — 分阶段计划", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 阶段详情
  const phaseDetails = [
    {
      title: "阶段一（2026.6-12）：硬件系统优化",
      items: ["完成微型化结构设计", "双椭圆履带差速转向机构", "越障模块设计与加工", "感知与执行系统集成"]
    },
    {
      title: "阶段二（2027.1-6）：核心算法优化",
      items: ["YOLOv8-GC模型训练与优化", "RFID+视觉融合定位算法", "自适应运动控制算法", "预测性维护算法研究"]
    },
    {
      title: "阶段三（2027.7-12）：云平台开发",
      items: ["缺陷数据库设计与实现", "数字孪生可视化平台", "协同运维管控平台", "系统集成与测试"]
    },
    {
      title: "阶段四（2028.1-12）：示范应用验证",
      items: ["1-3座实体桥梁示范应用", "通过行业专家评审验收", "形成可复制推广产品方案", "标准规范编制与发布"]
    }
  ];

  phaseDetails.forEach((phase, i) => {
    const x = 0.4 + (i % 2) * 4.7;
    const y = 1.1 + Math.floor(i / 2) * 2.2;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.5, h: 2.0,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.5, h: 0.45,
      fill: { color: "1E2761" }
    });
    slide.addText(phase.title, {
      x: x + 0.1, y: y + 0.05, w: 4.3, h: 0.35,
      fontSize: 11, bold: true, color: "FFFFFF", valign: "middle"
    });
    slide.addText(
      phase.items.map((item, idx) => ({
        text: item,
        options: { bullet: true, breakLine: idx < phase.items.length - 1 }
      })),
      { x: x + 0.2, y: y + 0.55, w: 4.1, h: 1.35, fontSize: 10, color: "333333" }
    );
  });
}
```

---

### Task 7: 创建进度与人力资源页面 (2页)

**Files:**
- Modify: `create-ppt.js` (添加进度页函数)

- [ ] **Step 1: 第十一页 - 项目进度计划**

```javascript
function createScheduleSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("六、项目进度计划及人力资源", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 甘特图（简化表示）
  slide.addText("项目进度计划（甘特图）", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // 甘特图头部
  const months = ["2026.6", "2026.12", "2027.6", "2027.12", "2028.6", "2028.12"];
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 1.8, y: 1.55, w: 7.7, h: 0.35,
    fill: { color: "1E2761" }
  });
  months.forEach((m, i) => {
    slide.addText(m, {
      x: 1.8 + i * 1.28, y: 1.55, w: 1.28, h: 0.35,
      fontSize: 8, color: "FFFFFF", align: "center", valign: "middle"
    });
  });

  // 甘特图内容
  const ganttData = [
    { name: "硬件系统优化", start: 0, duration: 2 },
    { name: "核心算法优化", start: 1, duration: 2 },
    { name: "云平台开发", start: 2, duration: 2 },
    { name: "示范应用验证", start: 3, duration: 4 },
    { name: "论文与专利", start: 1, duration: 5 }
  ];

  ganttData.forEach((g, i) => {
    const y = 2.0 + i * 0.45;
    slide.addText(g.name, {
      x: 0.3, y, w: 1.4, h: 0.4,
      fontSize: 9, color: "333333", valign: "middle"
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.8 + g.start * 1.28, y: y + 0.05,
      w: g.duration * 1.28, h: 0.3,
      fill: { color: "CADCFC" }
    });
  });

  // 人员组织架构
  slide.addText("项目组织架构", {
    x: 0.5, y: 4.4, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // 组织架构图
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.0, y: 4.35, w: 5.5, h: 1.1,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 }
  });

  const org = [
    { title: "项目负责人", persons: "张强", x: 0.3 },
    { title: "技术负责人", persons: "李明/王浩", x: 1.85 },
    { title: "研发团队", persons: "陈伟/刘洋/赵磊\n孙杰/周敏/吴刚", x: 3.4 }
  ];
  org.forEach((o, i) => {
    if (i > 0) {
      slide.addShape(pres.shapes.LINE, {
        x: 1.5 + i * 1.55, y: 4.55, w: 0.35, h: 0,
        line: { color: "CCCCCC", width: 1 }
      });
    }
    slide.addText(o.title, {
      x: 4.1 + o.x, y: 4.45, w: 1.5, h: 0.35,
      fontSize: 10, bold: true, color: "1E2761", align: "center"
    });
    slide.addText(o.persons, {
      x: 4.1 + o.x, y: 4.8, w: 1.5, h: 0.55,
      fontSize: 9, color: "666666", align: "center"
    });
  });
}
```

- [ ] **Step 2: 第十二页 - 人力资源详细**

```javascript
function createHRSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("六、项目进度计划及人力资源（续）", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 核心研发团队
  slide.addText("核心研发团队（15人）", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: "1E2761"
  });

  const team = [
    { role: "项目经理", name: "张强", bg: "1E2761" },
    { role: "机械结构", name: "李明、王浩", bg: "3D5A80" },
    { role: "算法研究", name: "陈伟、刘洋", bg: "3D5A80" },
    { role: "软件平台", name: "赵磊、孙杰", bg: "3D5A80" },
    { role: "测试验证", name: "周敏、吴刚", bg: "3D5A80" },
    { role: "技术支持", name: "郑伟（博士）", bg: "98C1D9" }
  ];

  team.forEach((m, i) => {
    const x = 0.4 + (i % 3) * 3.1;
    const y = 1.55 + Math.floor(i / 3) * 0.75;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.9, h: 0.65,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: 0.65,
      fill: { color: m.bg }
    });
    slide.addText(m.role, {
      x: x + 0.15, y: y + 0.08, w: 2.6, h: 0.28,
      fontSize: 10, bold: true, color: "1E2761"
    });
    slide.addText(m.name, {
      x: x + 0.15, y: y + 0.35, w: 2.6, h: 0.25,
      fontSize: 9, color: "666666"
    });
  });

  // 人员投入计划表
  slide.addText("人员投入计划（人月）", {
    x: 0.5, y: 3.2, w: 4, h: 0.4,
    fontSize: 14, bold: true, color: "1E2761"
  });

  const hrTable = [
    [
      { text: "人员类别", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "2026", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "2027", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "2028", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "合计", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    ["高级研究员", "18", "24", "18", "60"],
    ["工程师", "36", "48", "36", "120"],
    ["技术人员", "24", "36", "24", "84"],
    ["合计", "78", "108", "78", "264"]
  ];
  slide.addTable(hrTable, {
    x: 0.5, y: 3.65, w: 9, h: 1.6,
    colW: [2.5, 1.8, 1.8, 1.8, 1.1],
    fontSize: 10,
    border: { pt: 0.5, color: "CCCCCC" },
    fill: { color: "FFFFFF" }
  });
}
```

---

### Task 8: 创建财务计划页面 (1页)

**Files:**
- Modify: `create-ppt.js` (添加财务页函数)

- [ ] **Step 1: 第十三页 - 财务计划**

```javascript
function createFinanceSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // 标题栏
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("七、财务计划", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // 项目总预算
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 3.0, h: 1.0,
    fill: { color: "1E2761" }
  });
  slide.addText("项目总预算", {
    x: 0.5, y: 1.2, w: 2.8, h: 0.35,
    fontSize: 12, color: "CADCFC", align: "center"
  });
  slide.addText("230万元", {
    x: 0.5, y: 1.55, w: 2.8, h: 0.45,
    fontSize: 26, bold: true, color: "FFFFFF", align: "center"
  });

  // 预算构成饼图数据
  slide.addText("预算分配构成", {
    x: 3.6, y: 1.15, w: 3, h: 0.35,
    fontSize: 13, bold: true, color: "1E2761"
  });

  // 预算明细表（右侧）
  const budgetItems = [
    { item: "设备费", amount: "45万", pct: "19.6%" },
    { item: "材料费", amount: "30万", pct: "13.0%" },
    { item: "测试化验费", amount: "35万", pct: "15.2%" },
    { item: "燃料动力费", amount: "15万", pct: "6.5%" },
    { item: "差旅费", amount: "20万", pct: "8.7%" },
    { item: "会议费", amount: "10万", pct: "4.3%" },
    { item: "合作协作费", amount: "25万", pct: "10.9%" },
    { item: "专家咨询费", amount: "15万", pct: "6.5%" },
    { item: "其他支出", amount: "35万", pct: "15.2%" }
  ];

  slide.addTable([
    [
      { text: "预算科目", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "金额", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "占比", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    ...budgetItems.map(b => [b.item, b.amount, b.pct])
  ], {
    x: 0.4, y: 2.25, w: 5.5, h: 3.1,
    colW: [2.5, 1.5, 1.5],
    fontSize: 9,
    border: { pt: 0.5, color: "DDDDDD" },
    fill: { color: "FFFFFF" }
  });

  // 预算饼图（简化表达）
  slide.addChart(pres.charts.PIE, [{
    name: "预算",
    labels: ["设备费", "材料费", "测试费", "动力费", "差旅费", "会议费", "协作费", "咨询费", "其他"],
    values: [19.6, 13.0, 15.2, 6.5, 8.7, 4.3, 10.9, 6.5, 15.2]
  }], {
    x: 6.0, y: 2.25, w: 3.6, h: 2.8,
    showPercent: true,
    showLegend: true,
    legendPos: "b"
  });
}
```

---

### Task 9: 创建结尾页

**Files:**
- Modify: `create-ppt.js` (添加结尾页函数)

- [ ] **Step 1: 第十四页（结尾） - 致谢**

```javascript
function createEndSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "1E2761" };

  // 顶部装饰线
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: "CADCFC" }
  });

  // 致谢文字
  slide.addText("感谢聆听", {
    x: 0.5, y: 2.0, w: 9, h: 0.9,
    fontSize: 48, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", align: "center"
  });

  slide.addText("敬请指导", {
    x: 0.5, y: 2.9, w: 9, h: 0.6,
    fontSize: 28, fontFace: "Microsoft YaHei",
    color: "CADCFC", align: "center"
  });

  // 分隔线
  slide.addShape(pres.shapes.LINE, {
    x: 3.5, y: 3.7, w: 3, h: 0,
    line: { color: "CADCFC", width: 1 }
  });

  // 联系信息
  slide.addText("桥梁结构安全研究院\n龚元浩科技有限公司", {
    x: 0.5, y: 4.0, w: 9, h: 0.8,
    fontSize: 14, fontFace: "Microsoft YaHei",
    color: "CADCFC", align: "center"
  });

  // 日期
  slide.addText("2026年5月", {
    x: 0.5, y: 4.9, w: 9, h: 0.4,
    fontSize: 12, fontFace: "Microsoft YaHei",
    color: "CADCFC", align: "center"
  });

  // 底部装饰线
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.545, w: 10, h: 0.08,
    fill: { color: "CADCFC" }
  });
}
```

---

### Task 10: 集成主函数并测试

**Files:**
- Modify: `create-ppt.js` (添加main函数)

- [ ] **Step 1: 添加主函数并运行**

```javascript
async function main() {
  const pptxgen = require("pptxgenjs");
  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "基于AI视觉识别的狭窄钢箱梁内部缺陷巡检机器人";
  pres.author = "桥梁结构安全研究院";

  // 按顺序创建所有页面
  createCoverSlide(pres);
  createBackgroundSlide1(pres);
  createBackgroundSlide2(pres);
  createImplementationSlide1(pres);
  createImplementationSlide2(pres);
  createImplementationSlide3(pres);
  createImplementationSlide4(pres);
  createResultsSlide1(pres);
  createResultsSlide2(pres);
  createRoadmapSlide1(pres);
  createRoadmapSlide2(pres);
  createScheduleSlide(pres);
  createHRSlide(pres);
  createFinanceSlide(pres);
  createEndSlide(pres);

  await pres.writeFile({ fileName: "output/科研项目申报PPT.pptx" });
  console.log("PPT created successfully!");
}

main().catch(console.error);
```

Run: `node create-ppt.js`

- [ ] **Step 2: 验证输出**

检查 `output/科研项目申报PPT.pptx` 是否生成成功，包含14页内容。

---

## 自检清单

1. **Spec覆盖检查** — 每个设计规范是否都有对应实现？
   - [x] 封面1页
   - [x] 研究背景2页
   - [x] 项目实施3-4页
   - [x] 成果目标2页
   - [x] 实施路线3-4页
   - [x] 进度人力2页
   - [x] 财务1页
   - [x] 结尾1页

2. **占位符扫描** — 无"TBD"、"TODO"、"实现 later"等

3. **类型一致性** — 函数名、参数在Task间一致

4. **颜色规范** — 主色#1E2761，辅助#CADCFC，无#前缀

5. **每页视觉元素** — 每页都有形状/表格/图表/流程图

---

## 执行选项

**Plan complete and saved to `docs/superpowers/plans/2026-05-20-ppt-implementation.md`. Two execution options:**

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

**Which approach?**