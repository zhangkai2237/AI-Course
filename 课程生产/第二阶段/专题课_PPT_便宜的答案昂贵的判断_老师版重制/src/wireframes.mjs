import { C } from "./theme.mjs";
import { P, shape, text, badge, line } from "./components.mjs";

function bubble(slide, pptx, { x, y, w, h, who, msg, side = "left", tone = "blue" }) {
  const fill = tone === "blue" ? C.blueSoft : C.greenSoft;
  const accent = tone === "blue" ? C.blue : C.green;
  text(slide, who, P(side === "left" ? x : x + w - 92, y - 26, 92, 20), { size: 9, color: accent, bold: true, align: side === "left" ? "left" : "right" });
  shape(slide, pptx, "roundRect", P(x, y, w, h), fill, { color: accent, width: 1 }, { radius: 0.08 });
  text(slide, msg, P(x + 16, y + 10, w - 32, h - 20), { size: 11.5, breakLine: true });
}

export function conversationWireframe(slide, pptx) {
  shape(slide, pptx, "roundRect", P(110, 190, 1060, 420), C.white, { color: C.line, width: 1.2 }, { radius: 0.08 });
  shape(slide, pptx, "rect", P(110, 190, 1060, 50), C.ink);
  text(slide, "售后助手 · 对话原型", P(134, 202, 450, 27), { size: 12, color: C.white, bold: true });
  badge(slide, pptx, "PUBLIC SIMULATION", 925, 201, C.yellow, C.ink, 210);
  bubble(slide, pptx, { x: 145, y: 285, w: 520, h: 64, who: "用户", msg: "我想退刚买的耳机，能帮我看一下吗？", side: "left", tone: "blue" });
  bubble(slide, pptx, { x: 480, y: 385, w: 625, h: 85, who: "助手", msg: "可以。我先展示候选订单，请你确认具体订单和商品。确认前不会提交任何售后申请。", side: "right", tone: "green" });
  shape(slide, pptx, "roundRect", P(145, 520, 820, 56), C.paper, { color: C.line, width: 1 }, { radius: 0.06 });
  text(slide, "输入消息…", P(166, 533, 650, 30), { size: 11, color: C.muted });
  shape(slide, pptx, "roundRect", P(985, 520, 120, 56), C.blue, { color: C.blue, width: 1 }, { radius: 0.06 });
  text(slide, "发送", P(985, 533, 120, 30), { size: 11, color: C.white, bold: true, align: "center" });
}

export function orderConfirmationWireframe(slide, pptx) {
  shape(slide, pptx, "roundRect", P(90, 190, 1100, 430), C.white, { color: C.line, width: 1.2 }, { radius: 0.08 });
  text(slide, "请选择要处理的订单", P(124, 215, 640, 36), { size: 18, bold: true });
  badge(slide, pptx, "需用户确认", 980, 216, "9A6A00", C.yellowSoft, 160);
  const orders = [
    ["候选订单 A", "无线耳机 · 订单 #A1028", "¥399 · 昨日签收", "待确认"],
    ["候选订单 B", "耳机保护套 · 订单 #B7731", "¥59 · 7天前签收", "非本次商品"],
  ];
  orders.forEach((o, i) => {
    const y = 282 + i * 123;
    shape(slide, pptx, "roundRect", P(124, y, 1032, 102), i === 0 ? C.blueSoft : C.paper, { color: i === 0 ? C.blue : C.line, width: 1.2 }, { radius: 0.06 });
    shape(slide, pptx, "ellipse", P(148, y + 32, 28, 28), C.white, { color: i === 0 ? C.blue : C.line, width: 2 });
    if (i === 0) shape(slide, pptx, "ellipse", P(155, y + 39, 14, 14), C.blue, { color: C.blue, width: 1 });
    text(slide, o[0], P(198, y + 14, 170, 24), { size: 9, color: C.blue, bold: true });
    text(slide, o[1], P(198, y + 40, 480, 32), { size: 14, bold: true });
    text(slide, o[2], P(700, y + 40, 270, 30), { size: 11, color: C.muted });
    badge(slide, pptx, o[3], 982, y + 34, i === 0 ? "9A6A00" : C.muted, i === 0 ? C.yellowSoft : C.paper2, 140);
  });
  text(slide, "必须显示：商品、时间、金额、状态；只给订单号不足以确认对象。", P(126, 548, 770, 35), { size: 11, color: C.red, bold: true });
  shape(slide, pptx, "roundRect", P(950, 535, 206, 58), C.blue, { color: C.blue, width: 1 }, { radius: 0.06 });
  text(slide, "确认订单 A", P(950, 548, 206, 30), { size: 12, color: C.white, bold: true, align: "center" });
}

export function handoffWireframe(slide, pptx) {
  shape(slide, pptx, "roundRect", P(75, 190, 1130, 430), C.white, { color: C.line, width: 1.2 }, { radius: 0.08 });
  shape(slide, pptx, "rect", P(75, 190, 1130, 52), C.ink);
  text(slide, "人工接管工作台 · 原型", P(102, 202, 540, 28), { size: 12, color: C.white, bold: true });
  badge(slide, pptx, "公共模拟案例", 972, 202, C.red, C.redSoft, 195);
  const panels = [
    { x: 100, w: 310, label: "会话摘要", body: "用户希望退无线耳机；已确认订单 A；原因：左耳无声；未提交退款。", tone: C.blue },
    { x: 430, w: 310, label: "已确认 / 待补", body: "已确认：订单、商品、问题\n待补：图片、联系方式\n冲突：无", tone: C.green },
    { x: 760, w: 420, label: "转交原因与下一步", body: "原因：超出自动预受理范围\n建议：人工核对证据并决定后续\n人工可推翻机器建议", tone: C.red },
  ];
  panels.forEach(p => {
    shape(slide, pptx, "roundRect", P(p.x, 278, p.w, 245), C.paper, { color: C.line, width: 1 }, { radius: 0.06 });
    shape(slide, pptx, "rect", P(p.x, 278, p.w, 7), p.tone);
    text(slide, p.label, P(p.x + 18, 305, p.w - 36, 30), { size: 14, color: p.tone, bold: true });
    text(slide, p.body, P(p.x + 18, 350, p.w - 36, 118), { size: 11.5, valign: "top", breakLine: true });
    shape(slide, pptx, "roundRect", P(p.x + 18, 476, p.w - 36, 30), C.white, { color: C.line, width: 0.8 }, { radius: 0.04 });
    text(slide, "查看记录", P(p.x + 18, 480, p.w - 36, 20), { size: 9.5, color: p.tone, bold: true, align: "center" });
  });
  line(slide, pptx, 410, 400, 426, 400, C.blue, 1.5, true);
  line(slide, pptx, 740, 400, 756, 400, C.red, 1.5, true);
  text(slide, "状态：已受理，未批准退款", P(98, 552, 650, 36), { size: 12, color: C.red, bold: true });
  text(slide, "机器建议 ≠ 人工结论", P(825, 552, 350, 36), { size: 12, color: C.muted, bold: true, align: "right" });
}
