const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaBullseye, FaLayerGroup, FaProjectDiagram, FaMicrochip, FaFileAlt,
  FaChartBar, FaHistory, FaStar, FaTools, FaCogs, FaBolt, FaNetworkWired,
  FaWaveSquare, FaRulerHorizontal, FaClock, FaSignal, FaLightbulb, FaPlug,
  FaTerminal, FaCodeBranch, FaRobot, FaCheckCircle, FaExclamationTriangle,
  FaDatabase, FaFlask, FaSearch, FaArrowRight
} = require("react-icons/fa");

// ── 激光/光子学主题配色：深墨夜空 + 光子青 + 琥珀 ──
const C = {
  bg: "0B1526",      // 深墨蓝（主导色）
  bg2: "101E35",     // 卡片底
  card: "14233F",
  line: "1E3354",
  cyan: "2DD4BF",    // 光子青（强调）
  cyanD: "14B8A6",
  amber: "F5B041",   // 琥珀（点缀）
  text: "EAF2FB",
  mute: "8FA6C4",
  dim: "5B7194",
};
const FONT = "Microsoft YaHei";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

const bu = () => ({ code: "25B8", indent: 12 });
const shadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 90, opacity: 0.35 });

(async () => {
  const icons = {};
  const defs = {
    bullseye: FaBullseye, layer: FaLayerGroup, diagram: FaProjectDiagram, chip: FaMicrochip,
    file: FaFileAlt, chart: FaChartBar, history: FaHistory, star: FaStar, tools: FaTools,
    cogs: FaCogs, bolt: FaBolt, network: FaNetworkWired, wave: FaWaveSquare, ruler: FaRulerHorizontal,
    clock: FaClock, signal: FaSignal, bulb: FaLightbulb, plug: FaPlug, terminal: FaTerminal,
    branch: FaCodeBranch, robot: FaRobot, check: FaCheckCircle, warn: FaExclamationTriangle,
    db: FaDatabase, flask: FaFlask, search: FaSearch, arrow: FaArrowRight,
  };
  for (const [k, v] of Object.entries(defs)) {
    icons[k] = await icon(v, "#" + C.cyan);
    icons[k + "_a"] = await icon(v, "#" + C.amber);
    icons[k + "_d"] = await icon(v, "#" + C.dim);
  }

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "ZCode";
  pres.title = "PTS 激光自动测试平台 · 项目深度分析";

  // 通用：内容页头（小圆图标 + 页码）
  function header(s, iconKey, title, num) {
    s.background = { color: C.bg };
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: 0.38, w: 0.46, h: 0.46, fill: { color: C.cyanD, transparency: 78 }, line: { color: C.cyan, width: 1 } });
    s.addImage({ data: icons[iconKey], x: 0.61, y: 0.49, w: 0.24, h: 0.24 });
    s.addText(title, { x: 1.12, y: 0.34, w: 7.5, h: 0.55, fontSize: 25, bold: true, color: C.text, fontFace: FONT, margin: 0, valign: "middle" });
    s.addText(num, { x: 9.2, y: 0.42, w: 0.45, h: 0.4, fontSize: 12, color: C.dim, fontFace: "Courier New", align: "right", margin: 0 });
    s.addShape(pres.shapes.LINE, { x: 0.5, y: 1.02, w: 9.0, h: 0, line: { color: C.line, width: 0.75 } });
  }

  /* ═══ 1 · 封面 ═══ */
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    // 光束母题：一束青色光从左射向右侧靶点
    for (let i = 0; i < 5; i++) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: -0.5, y: 2.72 + i * 0.045, w: 8.6, h: 0.012,
        fill: { color: C.cyan, transparency: 55 + i * 8 }, line: { type: "none" },
      });
    }
    s.addShape(pres.shapes.OVAL, { x: 8.42, y: 2.6, w: 0.32, h: 0.32, fill: { color: C.amber }, line: { color: C.amber, width: 1.5, transparency: 40 } });
    s.addShape(pres.shapes.OVAL, { x: 8.12, y: 2.3, w: 0.92, h: 0.92, fill: { color: C.amber, transparency: 88 }, line: { type: "none" } });

    s.addText("PRECITEST SYSTEM", { x: 0.55, y: 0.72, w: 5, h: 0.4, fontSize: 14, color: C.cyan, charSpacing: 6, fontFace: "Courier New", bold: true, margin: 0 });
    s.addText([
      { text: "PTS 激光自动测试平台", options: { fontSize: 43, bold: true, color: C.text, breakLine: true } },
      { text: "项目深度分析报告", options: { fontSize: 43, bold: true, color: C.text } },
    ], { x: 0.55, y: 1.15, w: 8.2, h: 1.6, fontFace: FONT, margin: 0, lineSpacing: 58 });
    s.addText("Python / Tkinter · 多进程架构 · SCPI 仪器控制 · 自动化报告生成", {
      x: 0.55, y: 3.35, w: 8, h: 0.4, fontSize: 15, color: C.mute, fontFace: FONT, margin: 0,
    });
    s.addText([
      { text: "v5.3.6", options: { color: C.amber, bold: true } },
      { text: "   ·   16,496 行代码   ·   64 个源文件   ·   8 大测试模块   ·   92 次提交", options: { color: C.mute } },
    ], { x: 0.55, y: 4.55, w: 8.5, h: 0.4, fontSize: 13, fontFace: FONT, margin: 0 });
    s.addText("2026-08", { x: 8.6, y: 5.05, w: 0.9, h: 0.35, fontSize: 11, color: C.dim, fontFace: "Courier New", margin: 0, align: "right" });
  }

  /* ═══ 2 · 项目概览 ═══ */
  {
    const s = pres.addSlide();
    header(s, "bullseye", "项目概览 · 它是什么", "02");
    s.addText([
      { text: "PTS（PreciTest System）", options: { bold: true, color: C.cyan } },
      { text: " 是一套面向激光器测试的桌面级自动测试平台：以一台主控 GUI 调度光开关切换光路，驱动频谱仪、信号源、示波器、波长计、功率计等 8 类仪器，完成种子源激光的 RIN、线宽、时域、信噪比、单频、功率、PZT 调制与相噪共 8 项测试，并自动汇总生成 Word 报告。", options: {} },
    ], { x: 0.55, y: 1.3, w: 8.9, h: 1.05, fontSize: 15, color: C.text, fontFace: FONT, margin: 0, lineSpacing: 24 });

    const stats = [
      ["16.5k", "自有代码行数", "chart"],
      ["64", "Python 源文件", "file"],
      ["8+2", "测试模块 / 一键变体", "flask"],
      ["92", "Git 提交次数", "branch"],
    ];
    stats.forEach(([num, label, ic], i) => {
      const x = 0.55 + i * 2.28;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.62, w: 2.08, h: 2.15, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08, shadow: shadow() });
      s.addImage({ data: icons[ic], x: x + 0.79, y: 2.88, w: 0.5, h: 0.5 });
      s.addText(num, { x, y: 3.42, w: 2.08, h: 0.75, fontSize: 40, bold: true, color: C.amber, align: "center", fontFace: "Courier New", margin: 0 });
      s.addText(label, { x, y: 4.18, w: 2.08, h: 0.4, fontSize: 12.5, color: C.mute, align: "center", fontFace: FONT, margin: 0 });
    });
  }

  /* ═══ 3 · 总体架构 ═══ */
  {
    const s = pres.addSlide();
    header(s, "layer", "总体架构 · 三层解耦 + 多进程隔离", "03");
    const layers = [
      ["界面层", "BaseTestGUI · 8 个模块 GUI", "嵌套窗口 / 线程安全日志 / 可中断测试线程", "wave"],
      ["流程层", "BaseTestRunner · 模板方法", "清目录→连接→配置→循环测量→收尾→关资源", "cogs"],
      ["仪器层", "VisaInstrument · VISA 基类", "*IDN? 校验 / @py 后端优先回退 / SCPI 读写", "plug"],
    ];
    layers.forEach(([t, m, d, ic], i) => {
      const y = 1.32 + i * 1.02;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.55, y, w: 5.1, h: 0.88, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.07 });
      s.addImage({ data: icons[ic], x: 0.78, y: y + 0.27, w: 0.34, h: 0.34 });
      s.addText([
        { text: t + "  ", options: { bold: true, fontSize: 15, color: C.cyan } },
        { text: m, options: { fontSize: 12, color: C.mute } },
      ], { x: 1.3, y: y + 0.1, w: 4.2, h: 0.36, fontFace: FONT, margin: 0 });
      s.addText(d, { x: 1.3, y: y + 0.46, w: 4.25, h: 0.34, fontSize: 11, color: C.mute, fontFace: FONT, margin: 0 });
    });
    s.addText("core/（v5.3.0 重构产物，1,378 行）", { x: 0.55, y: 4.48, w: 5, h: 0.35, fontSize: 11.5, italic: true, color: C.dim, fontFace: FONT, margin: 0 });

    // 右侧：多进程模型
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.0, y: 1.32, w: 3.5, h: 3.5, fill: { color: C.bg2 }, line: { color: C.cyan, width: 1 }, rectRadius: 0.09, shadow: shadow() });
    s.addImage({ data: icons.diagram, x: 6.25, y: 1.55, w: 0.34, h: 0.34 });
    s.addText("多进程架构", { x: 6.72, y: 1.53, w: 2.6, h: 0.4, fontSize: 16, bold: true, color: C.text, fontFace: FONT, margin: 0 });
    s.addText([
      { text: "主平台 IntegratedPlatform", options: { bullet: bu(), color: C.text, breakLine: true } },
      { text: "每个测试模块 = 独立子进程", options: { bullet: bu(), color: C.text, breakLine: true } },
      { text: "msg_queue 子→主 报告状态", options: { bullet: bu(), color: C.text, breakLine: true } },
      { text: "cmd_queue 主→子 下发 START", options: { bullet: bu(), color: C.text, breakLine: true } },
      { text: "MODULE_REGISTRY 注册表", options: { bullet: bu(), color: C.text, breakLine: true } },
      { text: "单模块崩溃不拖垮平台", options: { bullet: bu(), color: C.amber } },
    ], { x: 6.28, y: 2.05, w: 3.05, h: 2.6, fontSize: 12.5, fontFace: FONT, margin: 0, paraSpaceAfter: 9, color: C.text });
  }

  /* ═══ 4 · 测试模块矩阵 ═══ */
  {
    const s = pres.addSlide();
    header(s, "flask", "测试模块矩阵 · 光路 A / B", "04");
    const mods = [
      ["RIN 相对强度噪声", "频谱仪 6 频段扫测\nDC / 放大常数修正", "signal"],
      ["线宽测量", "延时自外差法\n洛伦兹拟合求线宽", "ruler"],
      ["时域测试", "示波器测峰峰值 / 直流\n计算调制深度", "clock"],
      ["光谱信噪比 SNR", "OSA 扫谱测 SNR\n截图前重设 RefLevel", "chart"],
      ["单频测试", "DFB 温度 + 电流双扫\n边模抑制比判定", "search"],
      ["PZT 调制 / 波长", "波长计 DLL 测\n种子中心波长", "wave"],
      ["相位噪声", "pywinauto 驱动\n第三方上位机程序", "robot"],
      ["功率测试（光路B）", "USB 功率计\n输出功率与稳定性", "bolt"],
    ];
    mods.forEach(([t, d, ic], i) => {
      const x = 0.55 + (i % 4) * 2.28, y = 1.32 + Math.floor(i / 4) * 1.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.08, h: 1.78, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08, shadow: shadow() });
      s.addShape(pres.shapes.OVAL, { x: x + 0.16, y: y + 0.16, w: 0.5, h: 0.5, fill: { color: C.cyanD, transparency: 78 }, line: { color: C.cyan, width: 0.75 } });
      s.addImage({ data: icons[ic], x: x + 0.28, y: y + 0.28, w: 0.26, h: 0.26 });
      s.addText(t, { x: x + 0.16, y: y + 0.74, w: 1.8, h: 0.34, fontSize: 13, bold: true, color: C.text, fontFace: FONT, margin: 0 });
      s.addText(d, { x: x + 0.16, y: y + 1.08, w: 1.82, h: 0.62, fontSize: 9.5, color: C.mute, fontFace: FONT, margin: 0, lineSpacing: 13 });
    });
  }

  /* ═══ 5 · 一键测试编排 ═══ */
  {
    const s = pres.addSlide();
    header(s, "bolt", "一键测试 · 编排流程", "05");
    const steps = [
      ["01", "选择模块组合", "卡片勾选测试项"],
      ["02", "光开关分组", "按通道归并，串行调度"],
      ["03", "切通道→START", "切光开关后向子进程下发命令"],
      ["04", "等待完成", "轮询 msg_queue 状态消息"],
      ["05", "汇总报告", "docxtpl 生成 Word 报告"],
    ];
    steps.forEach(([n, t, d], i) => {
      const x = 0.55 + i * 1.86;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.4, w: 1.66, h: 1.72, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08 });
      s.addText(n, { x: x + 0.12, y: 1.52, w: 0.9, h: 0.42, fontSize: 22, bold: true, color: C.cyan, fontFace: "Courier New", margin: 0 });
      s.addText(t, { x: x + 0.12, y: 2.0, w: 1.45, h: 0.34, fontSize: 12.5, bold: true, color: C.text, fontFace: FONT, margin: 0 });
      s.addText(d, { x: x + 0.12, y: 2.36, w: 1.45, h: 0.66, fontSize: 9.5, color: C.mute, fontFace: FONT, margin: 0, lineSpacing: 13 });
      if (i < 4) s.addImage({ data: icons.arrow, x: x + 1.68, y: 2.12, w: 0.16, h: 0.16 });
    });
    const notes = [
      ["共用仪器串行化", "通道 3 的种子光测试在 RIN 结束后串行执行——两者共用同一台频谱仪 FSV3004"],
      ["数据跨模块复用", "信噪比程序的中心波长不再手填，自动使用 PZT 调制程序读到的种子波长值"],
      ["底噪自动测量", "通道 4 无光时自动执行底噪测量，纳入一键流程（v5.3.5）"],
    ];
    notes.forEach(([t, d], i) => {
      const y = 3.5 + i * 0.62;
      s.addImage({ data: icons.check, x: 0.6, y: y + 0.03, w: 0.22, h: 0.22 });
      s.addText([
        { text: t + "　", options: { bold: true, color: C.cyan } },
        { text: d, options: { color: C.mute } },
      ], { x: 0.95, y, w: 8.5, h: 0.5, fontSize: 12, fontFace: FONT, margin: 0 });
    });
  }

  /* ═══ 6 · 仪器控制 ═══ */
  {
    const s = pres.addSlide();
    header(s, "network", "仪器互联 · 四类自动化接口", "06");
    const rows = [
      ["terminal", "pyvisa · SCPI", "频谱仪 FSV3004 ×2、信号源 ×2、示波器、OSA、USB 光开关、USB 功率计；@py 纯 Python 后端优先，免装 NI-VISA，失败自动回退"],
      ["plug", "pyserial · RS-485", "DFB 种子激光器：115200 波特率自定义帧协议（帧头 0x50 0x00 / 帧尾 0x0D 0x0A），驱动电流与温度控制"],
      ["chip", "ctypes · DLL", "HighFinesse 波长计 wlmData.dll 完整 Python 封装（drivers/ 1,443 行），实时读取种子中心波长"],
      ["robot", "pywinauto · GUI 自动化", "跨进程点击第三方相噪上位机（Java AWT 窗口）按钮，按波长自动选择 1μm / 1.5μm 测量程序"],
    ];
    rows.forEach(([ic, t, d], i) => {
      const y = 1.32 + i * 0.98;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.56, h: 0.56, fill: { color: C.cyanD, transparency: 78 }, line: { color: C.cyan, width: 1 } });
      s.addImage({ data: icons[ic], x: 0.74, y: y + 0.24, w: 0.28, h: 0.28 });
      s.addText(t, { x: 1.4, y: y + 0.02, w: 3.0, h: 0.35, fontSize: 15, bold: true, color: C.cyan, fontFace: FONT, margin: 0 });
      s.addText(d, { x: 1.4, y: y + 0.38, w: 8.0, h: 0.52, fontSize: 11.5, color: C.mute, fontFace: FONT, margin: 0, lineSpacing: 15 });
      if (i < 3) s.addShape(pres.shapes.LINE, { x: 1.4, y: y + 0.92, w: 7.9, h: 0, line: { color: C.line, width: 0.5, dashType: "dot" } });
    });
  }

  /* ═══ 7 · 数据与报告链路 ═══ */
  {
    const s = pres.addSlide();
    header(s, "file", "数据链路 · 从测量到报告", "07");
    const flow = [
      ["仪器测量", "各模块输出 CSV\n波形 / 数值 / 截图", "signal"],
      ["结果落盘", "Rin/ WaveLength/\nSpectrumSNR/ … 目录", "db"],
      ["数据汇集", "data_collector\nassemble_report_data()", "search"],
      ["模板渲染", "docxtpl 渲染\ntemplate_default.docx", "file"],
      ["Word 报告", "输出至 C:\\PTS\\report\n含 InlineImage 图表", "check"],
    ];
    flow.forEach(([t, d, ic], i) => {
      const x = 0.55 + i * 1.86;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.55, w: 1.66, h: 2.0, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08, shadow: shadow() });
      s.addImage({ data: icons[ic], x: x + 0.66, y: 1.78, w: 0.36, h: 0.36 });
      s.addText(t, { x, y: 2.24, w: 1.66, h: 0.34, fontSize: 13, bold: true, color: C.text, align: "center", fontFace: FONT, margin: 0 });
      s.addText(d, { x: x + 0.08, y: 2.6, w: 1.5, h: 0.8, fontSize: 9.5, color: C.mute, align: "center", fontFace: FONT, margin: 0, lineSpacing: 13 });
      if (i < 4) s.addImage({ data: icons.arrow, x: x + 1.68, y: 2.44, w: 0.16, h: 0.16 });
    });
    s.addText([
      { text: "另有储备：", options: { bold: true, color: C.amber } },
      { text: " SQLite 结果库（data/pts_test_results.db）与 web 化雏形（FastAPI + WebSocket + 仪器 Agent + 前端，~1,400 行），为平台从单机走向网络化铺路。", options: { color: C.mute } },
    ], { x: 0.55, y: 4.0, w: 8.9, h: 0.8, fontSize: 12.5, fontFace: FONT, margin: 0, lineSpacing: 19 });
  }

  /* ═══ 8 · 代码规模分布（原生图表） ═══ */
  {
    const s = pres.addSlide();
    header(s, "chart", "代码规模 · 模块分布（行数）", "08");
    s.addChart(pres.charts.BAR, [{
      name: "行数",
      labels: ["path_a 种子源", "main_platform", "drivers 波长计", "core 基类", "web 雏形", "utils 组件", "path_b+report"],
      values: [6241, 1325, 1443, 1378, 1400, 900, 740],
    }], {
      x: 0.55, y: 1.25, w: 6.1, h: 3.9, barDir: "bar",
      chartColors: [C.cyan], chartArea: { fill: { color: C.bg } }, plotArea: { fill: { color: C.bg } },
      catAxisLabelColor: C.mute, catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT,
      valAxisLabelColor: C.dim, valAxisLabelFontSize: 10, valAxisLabelFontFace: "Courier New",
      valGridLine: { color: C.line, size: 0.5 }, catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.text, dataLabelFontSize: 10, dataLabelFontFace: "Courier New",
      showLegend: false, showTitle: false,
    });
    const facts = [
      ["38%", "path_a 占比——7 个测试模块与波长计封装是系统主体"],
      ["8 → 1", "v5.3.0 将 8 个模块共性下沉 core/，附迁移文档"],
      ["×2", "drivers/ 与 path_a/drivers/ 完全重复（1,443 行 ×2）"],
    ];
    facts.forEach(([n, d], i) => {
      const y = 1.45 + i * 1.2;
      s.addText(n, { x: 7.0, y, w: 2.4, h: 0.5, fontSize: 30, bold: true, color: C.amber, fontFace: "Courier New", margin: 0 });
      s.addText(d, { x: 7.0, y: y + 0.5, w: 2.45, h: 0.6, fontSize: 10.5, color: C.mute, fontFace: FONT, margin: 0, lineSpacing: 14 });
    });
  }

  /* ═══ 9 · 版本演进 ═══ */
  {
    const s = pres.addSlide();
    header(s, "history", "版本演进 · 从单机工具到平台化", "09");
    s.addShape(pres.shapes.LINE, { x: 0.9, y: 1.75, w: 8.2, h: 0, line: { color: C.cyan, width: 1.5 } });
    const vs = [
      ["v5.2 及以前", "单机 GUI 工具集", "功能逐个堆叠，模块间大量复制粘贴"],
      ["v5.3.0", "架构大重构", "抽取 core 三层基类、pyvisa-py 后端、配置 dataclass 化、网络配置工具"],
      ["v5.3.1–3.6", "测量细节打磨", "调制深度 / RefLevel / 通道档位 / 种子波长复用 / 统一前端风格"],
      ["进行中", "Web 化探索", "FastAPI + WebSocket + 仪器 Agent 管理系统，rin_service 剥离业务逻辑"],
    ];
    vs.forEach(([v, t, d], i) => {
      const x = 0.62 + i * 2.28;
      s.addShape(pres.shapes.OVAL, { x: x + 0.85, y: 1.66, w: 0.18, h: 0.18, fill: { color: i === 3 ? C.amber : C.cyan }, line: { color: C.bg, width: 2 } });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.1, w: 2.08, h: 2.3, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08, shadow: shadow() });
      s.addText(v, { x: x + 0.16, y: 2.26, w: 1.8, h: 0.32, fontSize: 13, bold: true, color: i === 3 ? C.amber : C.cyan, fontFace: "Courier New", margin: 0 });
      s.addText(t, { x: x + 0.16, y: 2.6, w: 1.8, h: 0.34, fontSize: 13.5, bold: true, color: C.text, fontFace: FONT, margin: 0 });
      s.addText(d, { x: x + 0.16, y: 2.98, w: 1.8, h: 1.3, fontSize: 10.5, color: C.mute, fontFace: FONT, margin: 0, lineSpacing: 15 });
    });
    s.addText("92 次提交 · 最新 v5.3.6（2026-08-06）", { x: 0.55, y: 4.75, w: 8.9, h: 0.35, fontSize: 12, italic: true, color: C.dim, fontFace: FONT, margin: 0, align: "center" });
  }

  /* ═══ 10 · 亮点 ═══ */
  {
    const s = pres.addSlide();
    header(s, "star", "工程亮点", "10");
    const pts = [
      ["架构清晰", "三层解耦 + 注册表式插件机制，新模块只改一处即可接入"],
      ["进程隔离", "每测试独立进程 + 双向队列通信，故障不扩散"],
      ["免环境依赖", "pyvisa-py 后端免装 NI-VISA；netsh 辅助 IP 一键配网"],
      ["中文工程化", "统一主题 / DPI 自适应 / Markdown 操作手册 / PyInstaller 打包"],
      ["四类接口全覆盖", "SCPI、串口、DLL、GUI 自动化，异构仪器统一纳管"],
      ["闭环到报告", "测量 → 落盘 → 汇集 → docxtpl 模板 → Word 报告全自动"],
    ];
    pts.forEach(([t, d], i) => {
      const x = 0.55 + (i % 2) * 4.6, y = 1.35 + Math.floor(i / 2) * 1.25;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.35, h: 1.08, fill: { color: C.card }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08 });
      s.addImage({ data: icons.check, x: x + 0.18, y: y + 0.2, w: 0.3, h: 0.3 });
      s.addText(t, { x: x + 0.62, y: y + 0.12, w: 3.6, h: 0.36, fontSize: 14, bold: true, color: C.cyan, fontFace: FONT, margin: 0 });
      s.addText(d, { x: x + 0.62, y: y + 0.5, w: 3.62, h: 0.5, fontSize: 11, color: C.mute, fontFace: FONT, margin: 0, lineSpacing: 15 });
    });
  }

  /* ═══ 11 · 改进建议 ═══ */
  {
    const s = pres.addSlide();
    header(s, "tools", "改进建议 · 下一步", "11");
    const rows = [
      ["路径硬编码", "数据目录固定 C:\\PTS\\zhongzi，建议相对化 / 环境变量，降低迁移成本"],
      ["脆弱的 GUI 自动化", "相噪模块按固定像素坐标点击 Java 窗口，对分辨率敏感，建议改控件级定位"],
      ["重复代码", "drivers/ 与 path_a/drivers/ 完全重复，合并可减 1,443 行"],
      ["缺工程配置", "无 requirements.txt / 自动化测试 / CI；abandoned、build/dist 应移出仓库"],
      ["逻辑与界面耦合", "测试逻辑内嵌 Tk GUI，子进程开销大且 web 端难复用；推广 rin_service 剥离模式"],
      ["消息协议", "magic-string 状态协议升级为结构化消息，便于 WebSocket 双端复用"],
    ];
    rows.forEach(([t, d], i) => {
      const y = 1.3 + i * 0.64;
      s.addImage({ data: icons.warn_a, x: 0.6, y: y + 0.04, w: 0.24, h: 0.24 });
      s.addText([
        { text: t + "　", options: { bold: true, color: C.text } },
        { text: d, options: { color: C.mute } },
      ], { x: 0.98, y, w: 8.4, h: 0.5, fontSize: 12.5, fontFace: FONT, margin: 0 });
    });
  }

  /* ═══ 12 · 结尾 ═══ */
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    for (let i = 0; i < 5; i++) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 1.4, y: 2.2 + i * 0.045, w: 7.2, h: 0.012,
        fill: { color: C.cyan, transparency: 60 + i * 8 }, line: { type: "none" },
      });
    }
    s.addShape(pres.shapes.OVAL, { x: 4.84, y: 2.08, w: 0.32, h: 0.32, fill: { color: C.amber } });
    s.addText("从单机工具，到平台化测试系统", { x: 0.5, y: 2.85, w: 9, h: 0.7, fontSize: 28, bold: true, color: C.text, align: "center", fontFace: FONT, margin: 0 });
    s.addText("PTS v5.3.6 · 深度分析报告 · 2026-08", { x: 0.5, y: 3.6, w: 9, h: 0.4, fontSize: 13, color: C.mute, align: "center", fontFace: FONT, margin: 0 });
  }

  await pres.writeFile({ fileName: "D:/Coding/Project/PTS/zhongzi/PTS项目深度分析.pptx" });
  console.log("done");
})();
