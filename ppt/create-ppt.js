import pptxgenjs from "pptxgenjs";

// Color constants
const PRIMARY = "1E2761";
const SECONDARY = "CADCFC";
const TEXT_DARK = "333333";
const TEXT_LIGHT = "F5F7FA";

// Helper function to avoid pptxgenjs shadow mutation issues
function makeShadow() {
  return {
    type: "outer",
    color: "000000",
    blur: 4,
    offset: 2,
    angle: 135,
    opacity: 0.2
  };
}

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
  ], { x: 5.4, y: 1.85, w: 4, h: 1.4, fontSize: 12, color: "333333" });

  // 底部：市场数据
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.6, w: 9.2, h: 1.7,
    fill: { color: "1E2761" }
  });
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

function createImplementationSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("二、研究目标与内容", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section: Research Goals KPIs
  slide.addText("研究目标", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // 4 KPI Cards
  const kpis = [
    { value: "≥95%", label: "缺陷识别准确率" },
    { value: "≤2cm", label: "定位精度" },
    { value: "<5%", label: "漏检率" },
    { value: "50%+", label: "检测效率提升" }
  ];

  kpis.forEach((kpi, i) => {
    const x = 0.5 + i * 2.35;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.55, w: 2.15, h: 1.1,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addText(kpi.value, {
      x: x, y: 1.6, w: 2.15, h: 0.6,
      fontSize: 28, bold: true, color: "1E2761", align: "center"
    });
    slide.addText(kpi.label, {
      x: x, y: 2.2, w: 2.15, h: 0.4,
      fontSize: 11, color: "333333", align: "center"
    });
  });

  // Section: Research Tasks
  slide.addText("研究内容", {
    x: 0.5, y: 2.85, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // 3 Research Task Cards
  const tasks = [
    { title: "复杂环境下SLAM定位技术", desc: "基于UWB+视觉融合的自主导航定位系统，适应狭窄空间" },
    { title: "轻量化缺陷检测算法", desc: "YOLOv8-GC算法优化，模型压缩50%，推理速度提升3倍" },
    { title: "云边协同检测平台", desc: "边缘实时检测+云端模型迭代，支持多机器人协作" }
  ];

  tasks.forEach((task, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 3.3, w: 2.9, h: 1.85,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });
    // Task number badge
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.1, y: 3.4, w: 0.35, h: 0.35,
      fill: { color: "1E2761" }
    });
    slide.addText(String(i + 1), {
      x: x + 0.1, y: 3.4, w: 0.35, h: 0.35,
      fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    slide.addText(task.title, {
      x: x + 0.5, y: 3.4, w: 2.3, h: 0.4,
      fontSize: 12, bold: true, color: "1E2761"
    });
    slide.addText(task.desc, {
      x: x + 0.1, y: 3.85, w: 2.7, h: 1.2,
      fontSize: 10, color: "333333"
    });
  });
}

function createImplementationSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("三、硬件系统优化", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Left side: Hardware components
  slide.addText("机器人硬件架构", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  const hwComponents = [
    { name: "双椭圆履带机构", specs: "宽度可调0.8-1.2m，爬坡能力30°" },
    { name: "双目视觉模块", specs: "基线距0.5m，深度精度±2mm@1m" },
    { name: "多传感器融合", specs: "UWB+IMU+编码器融合定位" },
    { name: "边缘计算单元", specs: "NVIDIA Jetson Orin，功耗35W" }
  ];

  hwComponents.forEach((comp, i) => {
    const y = 1.55 + i * 0.95;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 5.2, h: 0.85,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 0.08, h: 0.85,
      fill: { color: "1E2761" }
    });
    slide.addText(comp.name, {
      x: 0.7, y: y + 0.1, w: 4.8, h: 0.35,
      fontSize: 12, bold: true, color: "1E2761"
    });
    slide.addText(comp.specs, {
      x: 0.7, y: y + 0.45, w: 4.8, h: 0.35,
      fontSize: 10, color: "333333"
    });
  });

  // Right side: Specs table
  slide.addText("关键技术参数", {
    x: 5.9, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  const specsTable = [
    [
      { text: "参数", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "指标", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    ["检测范围", "5-8m"],
    ["定位精度", "≤2cm"],
    ["姿态偏差", "<1°"],
    ["工作温度", "-20~55°C"],
    ["防护等级", "IP67"],
    ["续航时间", "≥4h"]
  ];

  slide.addTable(specsTable, {
    x: 5.9, y: 1.55, w: 3.7, h: 3.4,
    colW: [1.8, 1.9],
    fontSize: 11,
    border: { pt: 0.5, color: "CCCCCC" },
    fill: { color: "FFFFFF" }
  });
}

function createImplementationSlide3(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("四、核心算法优化", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section intro
  slide.addText("基于YOLOv8-GC的轻量化缺陷检测算法", {
    x: 0.5, y: 1.1, w: 9, h: 0.4,
    fontSize: 14, color: "333333"
  });

  // 3 Algorithm Cards
  const algorithms = [
    {
      title: "GhostConv骨干网络",
      improvements: ["模型体积减少50%", "计算量降低40%", "特征提取效率提升"]
    },
    {
      title: "GC注意力机制",
      improvements: ["通道+空间双重注意力", "缺陷特征增强3倍", "背景干扰抑制"]
    },
    {
      title: "模型量化压缩",
      improvements: ["INT8量化部署", "推理速度提升3倍", "精度损失<1%"]
    }
  ];

  algorithms.forEach((algo, i) => {
    const x = 0.5 + i * 3.1;
    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.6, w: 2.9, h: 3.5,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });
    // Card header
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.6, w: 2.9, h: 0.7,
      fill: { color: "1E2761" }
    });
    slide.addText(algo.title, {
      x: x, y: 1.65, w: 2.9, h: 0.6,
      fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    // Improvements list
    algo.improvements.forEach((item, j) => {
      const y = 2.45 + j * 0.7;
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.15, y: y + 0.08, w: 0.15, h: 0.15,
        fill: { color: "CADCFC" }
      });
      slide.addText(item, {
        x: x + 0.4, y: y, w: 2.4, h: 0.5,
        fontSize: 11, color: "333333"
      });
    });
  });

  // Bottom comparison
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.2, w: 9, h: 0.35,
    fill: { color: "CADCFC" }
  });
  slide.addText("传统YOLOv8 → YOLOv8-GC: 参数量12.8M→6.4M | mAP@0.5: 89.2%→94.5% | 推理耗时85ms→28ms", {
    x: 0.5, y: 5.2, w: 9, h: 0.35,
    fontSize: 11, color: "1E2761", align: "center", valign: "middle"
  });
}

function createImplementationSlide4(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("五、云平台协作机制", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Left: Defect Database
  slide.addText("缺陷数据库", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.55, w: 4.3, h: 2.6,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
  });

  const defectStats = [
    { label: "缺陷类型", value: "3大类（裂纹、锈蚀、变形）" },
    { label: "样本数量", value: "50,000+张标注图像" },
    { label: "数据来源", value: "15座实桥检测数据" },
    { label: "更新频率", value: "每月迭代优化" }
  ];

  defectStats.forEach((stat, i) => {
    const y = 1.7 + i * 0.6;
    slide.addText(stat.label + ":", {
      x: 0.7, y: y, w: 1.4, h: 0.4,
      fontSize: 11, bold: true, color: "1E2761"
    });
    slide.addText(stat.value, {
      x: 2.1, y: y, w: 2.5, h: 0.4,
      fontSize: 11, color: "333333"
    });
  });

  // Right: Architecture Layers
  slide.addText("云边协同架构", {
    x: 5.1, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  const layers = [
    { name: "云端层", desc: "模型训练 | 数据存储 | 远程监控", color: "1E2761" },
    { name: "边缘层", desc: "实时推理 | 缺陷检测 | 本地决策", color: "3D5AA9" },
    { name: "设备层", desc: "机器人控制 | 传感器采集 | 自主导航", color: "6B8DD6" }
  ];

  layers.forEach((layer, i) => {
    const y = 1.55 + i * 0.95;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.1, y: y, w: 4.4, h: 0.85,
      fill: { color: layer.color },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addText(layer.name, {
      x: 5.2, y: y + 0.1, w: 4.2, h: 0.35,
      fontSize: 13, bold: true, color: "FFFFFF"
    });
    slide.addText(layer.desc, {
      x: 5.2, y: y + 0.45, w: 4.2, h: 0.35,
      fontSize: 10, color: "CADCFC"
    });
  });

  // Bottom: Collaboration flow
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.35, w: 9, h: 1.15,
    fill: { color: "1E2761" }
  });
  slide.addText("协作流程", {
    x: 0.7, y: 4.45, w: 2, h: 0.35,
    fontSize: 12, bold: true, color: "CADCFC"
  });
  slide.addText([
    { text: "现场采集 → 边缘推理 → 缺陷预警 → 数据回传 → 云端迭代 → 模型下发", options: {} }
  ], {
    x: 0.7, y: 4.8, w: 8.6, h: 0.5,
    fontSize: 12, color: "FFFFFF"
  });
}

function createResultsSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("六、阶段性成果", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section: Submitted achievements
  slide.addText("已提交成果", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // 4 Result Cards
  const results = [
    { title: "样机1台", desc: "双椭圆履带巡检机器人\n完成现场测试验证" },
    { title: "数据库1套", desc: "50,000+张缺陷样本\n15座实桥检测数据" },
    { title: "平台1套", desc: "云边协同检测平台\n支持多机器人协作" },
    { title: "论文与专利", desc: "SCI论文2篇\n发明专利3项" }
  ];

  results.forEach((item, i) => {
    const x = 0.5 + i * 2.35;
    // Card background - fresh shadow each time
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.55, w: 2.15, h: 1.8,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });
    // Card header accent
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.55, w: 2.15, h: 0.08,
      fill: { color: "CADCFC" }
    });
    // Title
    slide.addText(item.title, {
      x: x, y: 1.75, w: 2.15, h: 0.45,
      fontSize: 14, bold: true, color: "1E2761", align: "center"
    });
    // Description
    slide.addText(item.desc, {
      x: x + 0.1, y: 2.25, w: 1.95, h: 1.0,
      fontSize: 10, color: "333333", align: "center"
    });
  });

  // Key metrics footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.55, w: 9, h: 1.6,
    fill: { color: "1E2761" }
  });
  slide.addText("关键指标", {
    x: 0.7, y: 3.65, w: 2, h: 0.35,
    fontSize: 12, bold: true, color: "CADCFC"
  });

  const metrics = [
    { value: "94.5%", label: "缺陷识别准确率" },
    { value: "≤2cm", label: "定位精度" },
    { value: "<5%", label: "漏检率" },
    { value: "50%+", label: "效率提升" }
  ];

  metrics.forEach((metric, i) => {
    const x = 0.7 + i * 2.25;
    slide.addText(metric.value, {
      x: x, y: 4.05, w: 2, h: 0.6,
      fontSize: 32, bold: true, color: "CADCFC", align: "center"
    });
    slide.addText(metric.label, {
      x: x, y: 4.65, w: 2, h: 0.4,
      fontSize: 11, color: "CADCFC", align: "center"
    });
  });
}

function createResultsSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("七、技术与经济效益目标", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section: Technical Goals
  slide.addText("技术指标目标", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // Tech specs table - fresh shadow for the table container
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.55, w: 5, h: 2.5,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });

  const techSpecs = [
    [
      { text: "指标", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "目标值", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    ["缺陷识别准确率", "≥95%"],
    ["定位精度", "≤2cm"],
    ["漏检率", "<5%"],
    ["检测效率提升", "50%+"],
    ["模型压缩率", "50%"]
  ];

  slide.addTable(techSpecs, {
    x: 0.6, y: 1.65, w: 4.8, h: 2.3,
    colW: [2.4, 2.4],
    fontSize: 11,
    border: { pt: 0.5, color: "CCCCCC" },
    fill: { color: "FFFFFF" }
  });

  // Section: Economic Goals
  slide.addText("经济效益目标", {
    x: 5.7, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // Large number display - fresh shadow each time
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.7, y: 1.55, w: 3.9, h: 2.5,
    fill: { color: "1E2761" },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.15 }
  });

  slide.addText("600", {
    x: 5.7, y: 1.7, w: 3.9, h: 1.2,
    fontSize: 72, bold: true, color: "CADCFC", align: "center"
  });
  slide.addText("万元", {
    x: 5.7, y: 2.85, w: 3.9, h: 0.5,
    fontSize: 24, color: "CADCFC", align: "center"
  });
  slide.addText("预计年产值", {
    x: 5.7, y: 3.4, w: 3.9, h: 0.4,
    fontSize: 12, color: "CADCFC", align: "center"
  });

  // Bottom: Additional economic metrics
  const ecoMetrics = [
    { value: "30%", label: "成本降低" },
    { value: "2年", label: "投资回收期" }
  ];

  ecoMetrics.forEach((metric, i) => {
    const x = 0.5 + i * 4.6;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 4.25, w: 4.4, h: 1,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    slide.addText(metric.value, {
      x: x, y: 4.3, w: 4.4, h: 0.55,
      fontSize: 28, bold: true, color: "1E2761", align: "center"
    });
    slide.addText(metric.label, {
      x: x, y: 4.85, w: 4.4, h: 0.35,
      fontSize: 11, color: "333333", align: "center"
    });
  });
}

// Schedule slide with Gantt chart (2026.6-2028.12) and org structure
function createScheduleSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("九、项目进度计划", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section title: Gantt chart
  slide.addText("项目进度安排（2026.6 - 2028.12）", {
    x: 0.5, y: 1.05, w: 5, h: 0.35,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // Gantt chart background
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.45, w: 9, h: 2.6,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });

  // Timeline header (months)
  const months = ["2026.6", "2026.12", "2027.6", "2027.12", "2028.6", "2028.12"];
  const monthPositions = [2.3, 3.8, 5.3, 6.8, 8.3, 9.5];

  monthPositions.forEach((xPos, i) => {
    slide.addText(months[i], {
      x: xPos - 0.4, y: 1.55, w: 0.9, h: 0.25,
      fontSize: 8, color: "666666", align: "center"
    });
    // Vertical grid line
    slide.addShape(pres.shapes.LINE, {
      x: xPos, y: 1.8, w: 0, h: 2.1,
      line: { color: "E0E0E0", width: 0.5 }
    });
  });

  // Gantt phases
  const ganttPhases = [
    { name: "第一阶段：硬件优化", y: 1.9, xStart: 0.8, width: 2.2, color: "1E2761" },
    { name: "第二阶段：算法研发", y: 2.4, xStart: 2.4, width: 2.8, color: "3D5AA9" },
    { name: "第三阶段：平台建设", y: 2.9, xStart: 4.6, width: 2.6, color: "6B8DD6" },
    { name: "第四阶段：测试验收", y: 3.4, xStart: 6.6, width: 2.5, color: "89A8C1" }
  ];

  ganttPhases.forEach((phase) => {
    // Phase label
    slide.addText(phase.name, {
      x: 0.55, y: phase.y, w: 1.7, h: 0.4,
      fontSize: 9, color: "333333", align: "left", valign: "middle"
    });
    // Gantt bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: phase.xStart, y: phase.y + 0.05, w: phase.width, h: 0.3,
      fill: { color: phase.color }
    });
  });

  // Organizational structure section
  slide.addText("项目组织架构", {
    x: 0.5, y: 4.2, w: 4, h: 0.35,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // Org structure - center box (Project Manager)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.0, y: 4.6, w: 2.0, h: 0.7,
    fill: { color: "1E2761" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 }
  });
  slide.addText("项目经理", {
    x: 4.0, y: 4.65, w: 2.0, h: 0.35,
    fontSize: 11, bold: true, color: "FFFFFF", align: "center"
  });
  slide.addText("统筹协调", {
    x: 4.0, y: 4.95, w: 2.0, h: 0.3,
    fontSize: 9, color: "CADCFC", align: "center"
  });

  // Connecting lines from center to sub-groups
  const orgY = 5.3;
  const orgGroups = [
    { name: "机械结构", x: 0.7, w: 1.6 },
    { name: "算法研究", x: 2.7, w: 1.6 },
    { name: "软件平台", x: 4.7, w: 1.6 },
    { name: "测试验证", x: 6.7, w: 1.6 },
    { name: "技术支持", x: 8.7, w: 1.0 }
  ];

  // Horizontal line connecting all
  slide.addShape(pres.shapes.LINE, {
    x: 0.7, y: orgY, w: 9.0, h: 0,
    line: { color: "CADCFC", width: 1.5 }
  });

  // Vertical connectors and boxes
  orgGroups.forEach((group) => {
    // Vertical line from horizontal to box
    slide.addShape(pres.shapes.LINE, {
      x: group.x + group.w / 2, y: orgY, w: 0, h: 0.25,
      line: { color: "CADCFC", width: 1.5 }
    });
    // Group box
    slide.addShape(pres.shapes.RECTANGLE, {
      x: group.x, y: orgY + 0.25, w: group.w, h: 0.55,
      fill: { color: "FFFFFF" },
      line: { color: "3D5AA9", width: 1 }
    });
    slide.addText(group.name, {
      x: group.x, y: orgY + 0.3, w: group.w, h: 0.45,
      fontSize: 10, bold: true, color: "1E2761", align: "center", valign: "middle"
    });
  });
}

// HR slide with team member cards and investment table
function createHRSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("十、人力资源配置", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section title: Team structure
  slide.addText("核心团队成员", {
    x: 0.5, y: 1.05, w: 4, h: 0.35,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // 6 Team member cards
  const teamMembers = [
    { role: "项目经理", desc: "项目总体规划\n进度管理\n资源协调", headcount: "1人" },
    { role: "机械结构", desc: "履带设计\n结构优化\n装配调试", headcount: "2人" },
    { role: "算法研究", desc: "视觉算法\n缺陷检测\n模型优化", headcount: "2人" },
    { role: "软件平台", desc: "云边系统\n数据平台\n通信模块", headcount: "2人" },
    { role: "测试验证", desc: "功能测试\n性能评估\n现场验证", headcount: "2人" },
    { role: "技术支持", desc: "设备维护\n技术培训\n售后服务", headcount: "1人" }
  ];

  teamMembers.forEach((member, i) => {
    const row = Math.floor(i / 3);
    const col = i % 3;
    const x = 0.5 + col * 3.1;
    const y = 1.5 + row * 1.55;

    // Card background - fresh shadow each time
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.9, h: 1.4,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
    });
    // Role header
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.9, h: 0.45,
      fill: { color: "1E2761" }
    });
    // Role title
    slide.addText(member.role, {
      x: x, y: y + 0.02, w: 2.9, h: 0.4,
      fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    // Role description
    slide.addText(member.desc, {
      x: x + 0.1, y: y + 0.5, w: 1.9, h: 0.8,
      fontSize: 9, color: "333333"
    });
    // Headcount badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x + 2.1, y: y + 0.55, w: 0.7, h: 0.35,
      fill: { color: "CADCFC" }
    });
    slide.addText(member.headcount, {
      x: x + 2.1, y: y + 0.55, w: 0.7, h: 0.35,
      fontSize: 9, bold: true, color: "1E2761", align: "center", valign: "middle"
    });
  });

  // HR Investment table section
  slide.addText("年度人力资源投入（万元）", {
    x: 0.5, y: 4.7, w: 5, h: 0.35,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // Table container with shadow
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.1, w: 9, h: 0.5,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 }
  });

  const hrTableData = [
    [
      { text: "年份", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "2026年", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "2027年", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "2028年", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "合计", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ],
    [
      { text: "人员投入", options: { fill: { color: "CADCFC" } } },
      { text: "120", options: { fill: { color: "CADCFC" } } },
      { text: "180", options: { fill: { color: "CADCFC" } } },
      { text: "100", options: { fill: { color: "CADCFC" } } },
      { text: "400", options: { fill: { color: "CADCFC" }, bold: true } }
    ],
    [
      { text: "专家咨询", options: { fill: { color: "FFFFFF" } } },
      { text: "30", options: { fill: { color: "FFFFFF" } } },
      { text: "40", options: { fill: { color: "FFFFFF" } } },
      { text: "20", options: { fill: { color: "FFFFFF" } } },
      { text: "90", options: { fill: { color: "FFFFFF" }, bold: true } }
    ],
    [
      { text: "培训费用", options: { fill: { color: "CADCFC" } } },
      { text: "15", options: { fill: { color: "CADCFC" } } },
      { text: "20", options: { fill: { color: "CADCFC" } } },
      { text: "10", options: { fill: { color: "CADCFC" } } },
      { text: "45", options: { fill: { color: "CADCFC" }, bold: true } }
    ]
  ];

  slide.addTable(hrTableData, {
    x: 0.5, y: 5.1, w: 9, h: 0.5,
    colW: [1.5, 1.9, 1.9, 1.9, 1.8],
    fontSize: 10,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
    valign: "middle"
  });
}

// Budget data for pie chart and table
const BUDGET_ITEMS = [
  { name: "设备费", amount: 45, percent: 19.6 },
  { name: "材料费", amount: 30, percent: 13.0 },
  { name: "测试化验费", amount: 35, percent: 15.2 },
  { name: "燃料动力费", amount: 15, percent: 6.5 },
  { name: "差旅费", amount: 20, percent: 8.7 },
  { name: "会议费", amount: 10, percent: 4.3 },
  { name: "合作协作费", amount: 25, percent: 10.9 },
  { name: "专家咨询费", amount: 15, percent: 6.5 },
  { name: "其他支出", amount: 35, percent: 15.2 }
];

function createFinanceSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("十一、项目经费预算", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Total budget display - large card with fresh shadow
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.1, w: 4.2, h: 1.5,
    fill: { color: "1E2761" },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.15 }
  });
  slide.addText("总预算", {
    x: 0.7, y: 1.2, w: 2, h: 0.35,
    fontSize: 12, color: "CADCFC"
  });
  slide.addText("230", {
    x: 0.7, y: 1.5, w: 2.5, h: 0.9,
    fontSize: 64, bold: true, color: "CADCFC", align: "center"
  });
  slide.addText("万元", {
    x: 3.0, y: 1.95, w: 1, h: 0.4,
    fontSize: 18, color: "CADCFC"
  });
  slide.addText("项目总投资", {
    x: 0.7, y: 2.4, w: 3.8, h: 0.3,
    fontSize: 10, color: "CADCFC", align: "center"
  });

  // Pie chart
  const chartData = [{
    name: "预算分布",
    labels: BUDGET_ITEMS.map(item => item.name),
    values: BUDGET_ITEMS.map(item => item.amount)
  }];

  slide.addChart(pres.charts.PIE, chartData, {
    x: 4.9, y: 1.1, w: 4.8, h: 3.0,
    showTitle: false,
    showLegend: true,
    legendPos: "r",
    showPercent: false,
    showValue: false,
    chartColors: ["1E2761", "3D5AA9", "6B8DD6", "89A8C1", "A8C4D9", "C5D3E8", "D6E0F0", "E5EBF5", "F0F4FA"]
  });

  // Budget breakdown table
  slide.addText("预算明细表", {
    x: 0.5, y: 2.75, w: 4, h: 0.35,
    fontSize: 14, bold: true, color: "1E2761"
  });

  // Table container with fresh shadow
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.15, w: 9, h: 2.3,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 }
  });

  const tableData = [
    [
      { text: "预算项目", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "金额（万元）", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "占比", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "预算项目", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "金额（万元）", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } },
      { text: "占比", options: { fill: { color: "1E2761" }, color: "FFFFFF", bold: true } }
    ]
  ];

  // Build table rows (3 rows x 3 columns to show 9 items)
  for (let row = 0; row < 3; row++) {
    const rowData = [];
    for (let col = 0; col < 3; col++) {
      const idx = row * 3 + col;
      if (idx < BUDGET_ITEMS.length) {
        const item = BUDGET_ITEMS[idx];
        rowData.push({ text: item.name, options: { fill: { color: row % 2 === 0 ? "CADCFC" : "FFFFFF" } } });
        rowData.push({ text: String(item.amount), options: { fill: { color: row % 2 === 0 ? "CADCFC" : "FFFFFF" } } });
        rowData.push({ text: item.percent + "%", options: { fill: { color: row % 2 === 0 ? "CADCFC" : "FFFFFF" } } });
      }
    }
    tableData.push(rowData);
  }

  slide.addTable(tableData, {
    x: 0.6, y: 3.25, w: 8.8, h: 2.1,
    colW: [2.0, 1.4, 1.0, 2.0, 1.4, 1.0],
    fontSize: 10,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
    valign: "middle"
  });
}

// Phase colors for roadmap
const PHASE_COLORS = ["1E2761", "3D5AA9", "6B8DD6", "89A8C1"];

function createRoadmapSlide1(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("八、技术路线与实施计划", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section title
  slide.addText("总体技术路线", {
    x: 0.5, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // 4 Phase Cards
  const phases = [
    { title: "第一阶段", subtitle: "硬件系统优化", duration: "第1-3月" },
    { title: "第二阶段", subtitle: "核心算法优化", duration: "第4-6月" },
    { title: "第三阶段", subtitle: "云平台开发", duration: "第7-10月" },
    { title: "第四阶段", subtitle: "示范应用", duration: "第11-12月" }
  ];

  phases.forEach((phase, i) => {
    const x = 0.5 + i * 2.35;
    // Card background - fresh shadow each time
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.6, w: 2.15, h: 2.4,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });
    // Phase number circle
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.75, y: 1.75, w: 0.65, h: 0.65,
      fill: { color: PHASE_COLORS[i] }
    });
    slide.addText(String(i + 1), {
      x: x + 0.75, y: 1.75, w: 0.65, h: 0.65,
      fontSize: 20, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    // Phase title
    slide.addText(phase.title, {
      x: x, y: 2.5, w: 2.15, h: 0.35,
      fontSize: 11, color: "333333", align: "center"
    });
    // Phase subtitle
    slide.addText(phase.subtitle, {
      x: x, y: 2.85, w: 2.15, h: 0.5,
      fontSize: 14, bold: true, color: "1E2761", align: "center"
    });
    // Duration
    slide.addText(phase.duration, {
      x: x, y: 3.4, w: 2.15, h: 0.35,
      fontSize: 10, color: "666666", align: "center"
    });
    // Arrow between cards (except last)
    if (i < 3) {
      slide.addText("→", {
        x: x + 2.15, y: 2.6, w: 0.2, h: 0.5,
        fontSize: 20, color: "CADCFC", align: "center"
      });
    }
  });

  // Timeline bar at bottom
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.25, w: 9, h: 1.1,
    fill: { color: "1E2761" }
  });
  slide.addText("实施时间轴", {
    x: 0.7, y: 4.35, w: 2, h: 0.3,
    fontSize: 11, bold: true, color: "CADCFC"
  });

  // Timeline markers
  const timelineMonths = ["1月", "3月", "6月", "10月", "12月"];
  const timelinePositions = [0.7, 2.55, 4.45, 6.9, 8.5];

  timelinePositions.forEach((xPos, i) => {
    slide.addShape(pres.shapes.OVAL, {
      x: xPos, y: 4.7, w: 0.2, h: 0.2,
      fill: { color: "CADCFC" }
    });
    slide.addText(timelineMonths[i], {
      x: xPos - 0.3, y: 4.95, w: 0.8, h: 0.3,
      fontSize: 9, color: "CADCFC", align: "center"
    });
  });

  // Timeline line
  slide.addShape(pres.shapes.LINE, {
    x: 0.8, y: 4.8, w: 8.4, h: 0,
    line: { color: "CADCFC", width: 2 }
  });

  // Phase duration bars on timeline
  const phaseBarData = [
    { x: 0.8, w: 2.0, color: "3D5AA9" },
    { x: 2.8, w: 1.8, color: "6B8DD6" },
    { x: 4.6, w: 2.4, color: "89A8C1" },
    { x: 7.0, w: 2.2, color: "A8C4D9" }
  ];

  phaseBarData.forEach((bar) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: bar.x, y: 4.65, w: bar.w, h: 0.08,
      fill: { color: bar.color }
    });
  });
}

function createRoadmapSlide2(pres) {
  const slide = pres.addSlide();
  slide.background = { color: "F5F7FA" };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: "1E2761" }
  });
  slide.addText("八、技术路线与实施计划（续）", {
    x: 0.5, y: 0.2, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Microsoft YaHei", bold: true,
    color: "FFFFFF", margin: 0
  });

  // Section title
  slide.addText("各阶段详细任务", {
    x: 0.5, y: 1.05, w: 4, h: 0.35,
    fontSize: 16, bold: true, color: "1E2761"
  });

  // 4 Phase Detail Cards
  const phaseDetails = [
    {
      title: "硬件系统优化",
      tasks: [
        "双椭圆履带机构设计定型",
        "双目视觉模块集成",
        "多传感器融合调试",
        "样机装配与测试"
      ]
    },
    {
      title: "核心算法优化",
      tasks: [
        "YOLOv8-GC算法训练",
        "GhostConv模块集成",
        "GC注意力机制优化",
        "模型量化与部署"
      ]
    },
    {
      title: "云平台开发",
      tasks: [
        "边缘计算单元部署",
        "缺陷数据库建设",
        "云端模型迭代系统",
        "多机器人协作机制"
      ]
    },
    {
      title: "示范应用",
      tasks: [
        "实桥现场测试验证",
        "技术标准规范编制",
        "项目成果整理验收",
        "推广应用方案制定"
      ]
    }
  ];

  phaseDetails.forEach((phase, i) => {
    const x = 0.5 + i * 2.35;
    // Card background - fresh shadow each time
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.5, w: 2.15, h: 3.6,
      fill: { color: "FFFFFF" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });
    // Card header
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.5, w: 2.15, h: 0.65,
      fill: { color: PHASE_COLORS[i] }
    });
    // Phase number
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.1, y: 1.6, w: 0.4, h: 0.4,
      fill: { color: "FFFFFF", transparency: 20 }
    });
    slide.addText(String(i + 1), {
      x: x + 0.1, y: 1.6, w: 0.4, h: 0.4,
      fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle"
    });
    // Phase title
    slide.addText(phase.title, {
      x: x + 0.55, y: 1.6, w: 1.5, h: 0.45,
      fontSize: 11, bold: true, color: "FFFFFF", valign: "middle"
    });
    // Tasks list
    phase.tasks.forEach((task, j) => {
      const y = 2.3 + j * 0.65;
      // Bullet dot
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.15, y: y + 0.12, w: 0.12, h: 0.12,
        fill: { color: PHASE_COLORS[i] }
      });
      // Task text
      slide.addText(task, {
        x: x + 0.35, y: y, w: 1.7, h: 0.6,
        fontSize: 10, color: "333333"
      });
    });
  });
}

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

// Create presentation with cover slide
const pres = new pptxgenjs();
pres.layout = "LAYOUT_16x9";
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
pres.writeFile({ fileName: "output/科研项目申报PPT.pptx" })
  .then(() => console.log("Presentation created: output/科研项目申报PPT.pptx"))
  .catch(err => console.error("Error:", err));