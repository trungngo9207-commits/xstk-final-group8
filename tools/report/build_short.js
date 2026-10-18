// Short B1–B2 section (≈1.5 pages) for the 8-page team report. Usage: node build_short.js <en|vi> <out.docx>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, BorderStyle, TableLayoutType,
} = require("docx");

const LANG = process.argv[2];
const OUT = process.argv[3];
const L = (en, vi) => (LANG === "vi" ? vi : en);
const locale = L("en-US", "vi-VN");

const cells = JSON.parse(fs.readFileSync(path.join(__dirname, "cells.json"), "utf8"));
const V = (ref, sheet = "Part B") => cells[`${sheet}!${ref}`].v;

// ---------- numbers in the document's language (vi: 29,58 · 1.067,26)
const N = (v, d) => Number(v).toLocaleString(locale, { minimumFractionDigits: d, maximumFractionDigits: d });
const $ = (v, d = 2) => "$" + N(v, d);
const sci = (v, d = 2) => {
  const [m, e] = Number(v).toExponential(d).split("e");
  return `${N(Number(m), d)}×10^{${e.replace("-", "−").replace("+", "")}}`;
};
const SEP = L(", ", "; "); // function-argument separator as it reads in each locale

// ---------- styling
const FONT = "Arial", ACCENT = "1F4E79", GREY = "595959";
const PAGE_W = 11906, MARGIN = 1300, CONTENT_W = PAGE_W - 2 * MARGIN;
const BASE = 21; // 10.5 pt

function runs(text, base = {}) {
  const out = [];
  let bold = false, i = 0, buf = "";
  const flush = () => { if (buf) out.push(new TextRun({ text: buf, bold, font: FONT, ...base })); buf = ""; };
  while (i < text.length) {
    if (text.startsWith("**", i)) { flush(); bold = !bold; i += 2; continue; }
    if ((text[i] === "_" || text[i] === "^") && text[i + 1] === "{") {
      const close = text.indexOf("}", i);
      flush();
      out.push(new TextRun({ text: text.slice(i + 2, close), bold, font: FONT, ...base,
        ...(text[i] === "_" ? { subScript: true } : { superScript: true }) }));
      i = close + 1; continue;
    }
    buf += text[i]; i++;
  }
  flush();
  return out;
}

const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), spacing: { after: o.after ?? 100, line: 264 },
  alignment: o.align, indent: o.indent });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(t, { bold: true }) });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(t, { bold: true }) });
const step = (label, t) => new Paragraph({ children: runs(`**${label}** ${t}`), spacing: { after: 40, line: 252 },
  indent: { left: 280, hanging: 280 } });

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, header, rows, numCols = []) {
  const mk = (t, w, j, head) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders, margins: { top: 50, bottom: 50, left: 90, right: 90 },
    shading: head ? { fill: "D9E2F3", type: ShadingType.CLEAR, color: "auto" } : (j === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR, color: "auto" } : undefined),
    children: [new Paragraph({ children: runs(head ? `**${t}**` : String(t), { size: 19 }), spacing: { after: 0, line: 240 },
      alignment: !head && numCols.includes(j) ? AlignmentType.RIGHT : (head && j > 0 ? AlignmentType.CENTER : AlignmentType.LEFT),
      keepNext: true })],
  });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, j) => mk(h, widths[j], j, true)) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, j) => mk(c, widths[j], j, false)) }))],
  });
}
const gap = () => new Paragraph({ children: [], spacing: { after: 40 } });

// ---------- values
const n = V("B4"), mean = V("B5"), sd = V("B6"), se = V("B7"), tcrit = V("B8"), lo = V("B9"), hi = V("B10");
const t1 = V("B12"), p1 = V("B13");
const n1 = V("B21"), n2 = V("C21"), m1 = V("B22"), m2 = V("C22"), v1 = V("B23"), v2 = V("C23");
const diff = V("B26"), seW = V("B27"), tW = V("B28"), dfW = V("B29"), pW = V("B30"), pWt = V("B31");
const sp2 = V("B35"), seP = V("B36"), tP = V("B37"), pP = V("B39");
const r2 = V("B42"), sey = V("B44"), rdf = V("B47");
const median = V("B19", "Part A"), skew = V("B24", "Part A");
const ratio = Math.round(v1 / v2);

const c = [];
c.push(P(L("All numbers are computed with formulas in Group_8.xlsx, sheet Part B (cell references in brackets).",
  "Mọi số liệu được tính bằng công thức trong Group_8.xlsx, sheet Part B (ô tham chiếu ghi trong ngoặc)."),
  { run: { italics: true, color: GREY, size: 19 }, after: 60 }));

// ===== B1
c.push(H1(L("B1 · Continuous Y, no X  (Price = β_{0} + ε)", "B1 · Y liên tục, không có X  (Price = β_{0} + ε)")));
c.push(table([1000, 1450, 1550, 1700, 1500, 2106],
  L(["n (B4)", "Mean x̄ (B5)", "Std. dev. s (B6)", "SE = s/√n (B7)", "t critical (B8)", "95% CI (B9:B10)"],
    ["n (B4)", "Trung bình x̄ (B5)", "Độ lệch chuẩn s (B6)", "Sai số chuẩn s/√n (B7)", "t tới hạn (B8)", "Khoảng tin cậy 95% (B9:B10)"]),
  [[String(n), $(mean), $(sd), $(se, 3), N(tcrit, 3), `[${$(lo)}; ${$(hi)}]`]], [1, 2, 3, 4, 5]));
c.push(gap());
c.push(P(L(
  `The average bottle costs **${$(mean)}**. Prices vary widely (s = ${$(sd)}), but with ${n} wines the mean is estimated precisely ` +
  `(SE = ${$(se)}), giving the 95% CI **[${$(lo)}; ${$(hi)}]** = x̄ ± T.INV.2T(0.05${SEP}805)·SE, confirmed by CONFIDENCE.T (B15). ` +
  `95% of intervals built this way contain the true mean μ; the interval says nothing about individual bottles (the median is only ${$(median, 0)}).`,
  `Giá trung bình một chai là **${$(mean)}**. Giá giữa các chai chênh lệch rất lớn (s = ${$(sd)}), nhưng với ${n} chai, giá trung bình ` +
  `được ước lượng khá chính xác (SE = ${$(se)}), cho khoảng tin cậy 95% là **[${$(lo)}; ${$(hi)}]** = x̄ ± T.INV.2T(0,05${SEP}805)·SE, ` +
  `kiểm tra lại bằng CONFIDENCE.T (B15). Cách hiểu đúng: 95% các khoảng được xây dựng theo cách này chứa giá trung bình thật μ; ` +
  `khoảng này không mô tả giá của từng chai (trung vị chỉ là ${$(median, 0)}).`)));
c.push(P(L("**Test of H_{0}: μ = 25**", "**Kiểm định H_{0}: μ = 25**"), { after: 40 }));
c.push(step("(1)", L("Hypotheses: H_{0}: μ = 25; H_{1}: μ ≠ 25; α = 0.05.", "Giả thuyết: H_{0}: μ = 25; H_{1}: μ ≠ 25; α = 0,05.")));
c.push(step("(2)", L(`Test statistic: one-sample t = (x̄ − 25)/SE = (${N(mean, 2)} − 25)/${N(se, 3)} = **${N(t1, 2)}**, df = ${n - 1} (B12).`,
  `Thống kê kiểm định: t một mẫu = (x̄ − 25)/SE = (${N(mean, 2)} − 25)/${N(se, 3)} = **${N(t1, 2)}**, bậc tự do ${n - 1} (B12).`)));
c.push(step("(3)", L(`p-value = T.DIST.2T(${N(t1, 2)}${SEP}${n - 1}) = **${N(p1, 5)}** (B13) < 0.05.`,
  `Giá trị p = T.DIST.2T(${N(t1, 2)}${SEP}${n - 1}) = **${N(p1, 5)}** (B13) < 0,05.`)));
c.push(step("(4)", L("Decision: **reject H_{0}** (B14). The mean price is significantly above $25, consistent with $25 lying outside the 95% CI.",
  "Kết luận: **bác bỏ H_{0}** (B14). Giá trung bình cao hơn $25 một cách có ý nghĩa thống kê, khớp với việc $25 nằm ngoài khoảng tin cậy 95%.")));
c.push(gap());
c.push(P(L(
  "**Why the mean estimates β_{0}:** with no X the model predicts the same value β_{0} for every wine; OLS chooses β_{0} to minimise " +
  `Σ(Price − β_{0})^{2}, and setting the derivative −2Σ(Price − β_{0}) to zero gives β_{0} = x̄, so b_{0} = ${$(mean)}.`,
  "**Vì sao trung bình là ước lượng của β_{0}:** mô hình không có X nên dự báo cùng một giá trị β_{0} cho mọi chai; OLS chọn β_{0} để tổng " +
  `bình phương sai số Σ(Price − β_{0})^{2} nhỏ nhất, và cho đạo hàm −2Σ(Price − β_{0}) = 0 ta được β_{0} = x̄, nên b_{0} = ${$(mean)}.`)));

// ===== B2
c.push(H1(L("B2 · Continuous Y, one dummy X  (Price = β_{0} + β_{1}·Napa + ε)", "B2 · Y liên tục, một biến giả X  (Price = β_{0} + β_{1}·Napa + ε)")));
c.push(table([2306, 2300, 2300, 2400],
  L(["Group", "n (B21:C21)", "Mean price (B22:C22)", "Variance (B23:C23)"], ["Nhóm", "Số chai n (B21:C21)", "Giá trung bình (B22:C22)", "Phương sai (B23:C23)"]),
  [[L("Napa (Napa = 1)", "Napa (Napa = 1)"), String(n1), $(m1), N(v1, 2)],
   [L("Other (Napa = 0)", "Vùng khác (Napa = 0)"), String(n2), $(m2), N(v2, 2)]], [1, 2, 3]));

c.push(H2(L("B2.1 · Welch t-test (unequal variances)", "B2.1 · Kiểm định t Welch (phương sai không bằng nhau)")));
c.push(step("(1)", L("Hypotheses: H_{0}: μ_{Napa} = μ_{Other}; H_{1}: μ_{Napa} ≠ μ_{Other}; α = 0.05.",
  "Giả thuyết: H_{0}: μ_{Napa} = μ_{Khác}; H_{1}: μ_{Napa} ≠ μ_{Khác}; α = 0,05.")));
c.push(step("(2)", L(
  `Test statistic: Welch t = (${N(m1, 2)} − ${N(m2, 2)}) / √(${N(v1, 2)}/${n1} + ${N(v2, 2)}/${n2}) = ${N(diff, 2)}/${N(seW, 3)} = **${N(tW, 2)}**, df = ${N(dfW, 1)} (B26:B29).`,
  `Thống kê kiểm định: t Welch = (${N(m1, 2)} − ${N(m2, 2)}) / √(${N(v1, 2)}/${n1} + ${N(v2, 2)}/${n2}) = ${N(diff, 2)}/${N(seW, 3)} = **${N(tW, 2)}**, bậc tự do ${N(dfW, 1)} (B26:B29).`)));
c.push(step("(3)", L(`p-value = **${sci(pW)}** (B30, T.DIST.2T); Excel's T.TEST type 3 gives ${sci(pWt)} (B31); p < 0.05.`,
  `Giá trị p = **${sci(pW)}** (B30, T.DIST.2T); hàm T.TEST loại 3 cho ${sci(pWt)} (B31); p < 0,05.`)));
c.push(step("(4)", L("Decision: **reject H_{0}** (B32).", "Kết luận: **bác bỏ H_{0}** (B32).")));
c.push(gap());
c.push(P(L(
  `Napa wines cost on average **${$(diff)} more per bottle** than other wines, far beyond sampling noise. Welch's test is the right ` +
  `choice because the Napa variance is about ${ratio} times larger. This is a raw gap: score, age, volume and variety are not held ` +
  "constant, so it is not yet the premium of the Napa name (see B5 and Part C).",
  `Trung bình rượu Napa đắt hơn rượu vùng khác **${$(diff)} mỗi chai**, chênh lệch quá lớn để do ngẫu nhiên khi chọn mẫu. Kiểm định Welch ` +
  `là phù hợp vì phương sai của nhóm Napa lớn gấp khoảng ${ratio} lần. Đây mới là chênh lệch thô: chưa kiểm soát điểm chất lượng, ` +
  "tuổi, sản lượng và giống nho, nên chưa phải là mức giá cộng thêm (premium) của riêng nhãn Napa (xem B5 và Phần C).")));

c.push(H2(L("B2.2 · Regression of Price on the Napa dummy", "B2.2 · Hồi quy Price theo biến giả Napa")));
c.push(P(L(`LINEST (K42:L46): **Predicted Price = ${N(V("B53"), 2)} + ${N(V("B54"), 2)}·Napa** (R^{2} = ${N(r2, 3)}).`,
  `Hàm LINEST (K42:L46): **Giá dự báo = ${N(V("B53"), 2)} + ${N(V("B54"), 2)}·Napa** (R^{2} = ${N(r2, 3)}).`), { after: 60 }));
const coef = (r, name) => [name, N(V("B" + r), 2), N(V("C" + r), 3), N(V("D" + r), 2), sci(V("E" + r))];
c.push(table([2306, 1750, 1750, 1500, 2000],
  L(["Variable (A53:A54)", "Coefficient", "Std. error", "t Stat", "p-value"], ["Biến (A53:A54)", "Hệ số", "Sai số chuẩn", "t", "Giá trị p"]),
  [coef(53, L("Intercept", "Hệ số chặn")), coef(54, "Napa valley")], [1, 2, 3, 4]));
c.push(gap());
c.push(P(L(
  `The intercept equals the mean price of non-Napa wines (${$(m2)}, C22) and the slope equals ${N(m1, 2)} − ${N(m2, 2)} = ${N(diff, 2)}, ` +
  "the difference between the two group means (B26); check cells B56 and B57 are both 0. The model predicts b_{0} when Napa = 0 and " +
  "b_{0} + b_{1} when Napa = 1, and OLS sets each prediction to its group mean. The Napa dummy alone explains only " +
  `${N(r2 * 100, 1)}% of the variation in price.`,
  `Hệ số chặn bằng giá trung bình của nhóm không thuộc Napa (${$(m2)}, C22), còn hệ số góc bằng ${N(m1, 2)} − ${N(m2, 2)} = ${N(diff, 2)}, ` +
  "tức chênh lệch giữa hai trung bình nhóm (B26); hai ô kiểm tra B56 và B57 đều bằng 0. Lý do: khi Napa = 0 mô hình dự báo b_{0}, khi " +
  "Napa = 1 dự báo b_{0} + b_{1}, và OLS đặt mỗi dự báo bằng trung bình của nhóm đó. Riêng biến Napa chỉ giải thích được " +
  `${N(r2 * 100, 1)}% biến động của giá.`)));

c.push(H2(L("B2.3 · Which t-test does the regression reproduce?", "B2.3 · Hồi quy tái tạo chính xác kiểm định t nào?")));
c.push(table([2706, 2200, 2200, 2200],
  L(["", "Welch t-test", "Regression slope", "Pooled t-test"], ["", "t Welch", "Hệ số góc hồi quy", "t gộp"]),
  [[L("Standard error", "Sai số chuẩn"), `${N(seW, 3)} (B27)`, `${N(V("C54"), 3)} (C54)`, `${N(seP, 3)} (B36)`],
   [L("t statistic", "Thống kê t"), `${N(tW, 2)} (B28)`, `${N(V("D54"), 2)} (D54)`, `${N(tP, 2)} (B37)`],
   [L("Degrees of freedom", "Bậc tự do"), `${N(dfW, 1)} (B29)`, `${rdf} (B47)`, `${V("B38")} (B38)`],
   [L("p-value", "Giá trị p"), `${sci(pW)} (B30)`, `${sci(V("E54"))} (E54)`, `${sci(pP)} (B39)`]], [1, 2, 3]));
c.push(gap());
c.push(P(L(
  "The slope's t equals the **pooled (equal-variance) two-sample t-test** (T.TEST type 2; check cell B58 = 0), not Welch's. OLS assumes " +
  "one common error variance for every wine (homoskedasticity) and estimates it from all residuals: SSE/(n − 2) = " +
  `${N(sey, 2)}^{2} = ${N(sp2, 2)}, exactly the pooled variance s_{p}^{2} (B35), so its standard error (${N(seP, 3)}) is the pooled one. ` +
  `Since the Napa variance is about ${ratio} times larger, Welch's standard error (${N(seW, 3)}) is the more reliable; both tests reject H_{0}.`,
  "Thống kê t của hệ số góc trùng với **kiểm định t gộp** (giả định phương sai bằng nhau; T.TEST loại 2; ô kiểm tra B58 = 0), không phải t Welch. " +
  "Nguyên nhân: OLS giả định mọi quan sát có cùng một phương sai sai số (phương sai đồng nhất) và ước lượng nó từ toàn bộ phần dư: " +
  `SSE/(n − 2) = ${N(sey, 2)}^{2} = ${N(sp2, 2)}, đúng bằng phương sai gộp s_{p}^{2} (B35), nên sai số chuẩn của hệ số góc (${N(seP, 3)}) ` +
  `chính là sai số chuẩn gộp. Vì phương sai nhóm Napa lớn gấp khoảng ${ratio} lần, sai số chuẩn Welch (${N(seW, 3)}) đáng tin cậy hơn; ` +
  "tuy vậy cả hai kiểm định đều bác bỏ H_{0}.")));

c.push(H2(L("B2.4 · Safer rank-based test", "B2.4 · Kiểm định dựa trên thứ hạng")));
c.push(P(L(
  `Price is strongly right-skewed (skewness ${N(skew, 2)}; mean ${$(mean)} > median ${$(median, 0)}; histogram in A5), so a few very ` +
  "expensive bottles can drive a comparison of means. The safer check is the **Wilcoxon rank-sum (Mann–Whitney U) test**, which compares " +
  "the ranks of prices instead of their values, so outliers cannot dominate the result.",
  `Giá lệch phải mạnh (hệ số bất đối xứng ${N(skew, 2)}; trung bình ${$(mean)} > trung vị ${$(median, 0)}; xem histogram ở A5), nên ` +
  "một vài chai rất đắt có thể chi phối phép so sánh trung bình. Kiểm định an toàn hơn là **kiểm định tổng hạng Wilcoxon (Mann–Whitney U)**: " +
  "so sánh thứ hạng của giá thay vì giá trị, nên các giá trị ngoại lai không chi phối được kết quả.")));

const doc = new Document({
  creator: "Group 8",
  title: L("Group 8 – B1–B2 (short)", "Nhóm 8 – B1–B2 (bản rút gọn)"),
  styles: {
    default: { document: { run: { font: FONT, size: BASE } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: ACCENT }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: FONT, color: "222222" }, paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  sections: [{ properties: { page: { size: { width: PAGE_W, height: 16838 },
    margin: { top: 1100, bottom: 1100, left: MARGIN, right: MARGIN } } }, children: c }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length); });
