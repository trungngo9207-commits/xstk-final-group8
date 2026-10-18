// Builds the B1–B2 section of the Group 8 report as a .docx (A4, Google-Docs-friendly).
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, Footer, PageNumber, TableLayoutType,
} = require("docx");

const cells = JSON.parse(fs.readFileSync(path.join(__dirname, "cells.json"), "utf8"));
const V = (ref, sheet = "Part B") => {
  const c = cells[`${sheet}!${ref}`];
  if (!c) throw new Error("missing cell " + ref);
  return c.v;
};
const F = (ref, sheet = "Part B") => String(cells[`${sheet}!${ref}`].f);

// ---------- number formatting
const fx = (v, d) => Number(v).toFixed(d);
const money0 = (v) => "$" + Number(v).toLocaleString("en-US", { maximumFractionDigits: 0 });
const money = (v) => "$" + Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const num = (v, d) => Number(v).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const sci = (v, d = 2) => {
  const [m, e] = Number(v).toExponential(d).split("e");
  return `${m}×10^{${e.replace("-", "−").replace("+", "")}}`;
};

// ---------- styling constants
const FONT = "Arial";
const MONO = "Courier New";
const ACCENT = "1F4E79";
const GREY = "595959";
const PAGE_W = 11906, MARGIN = 1300;
const CONTENT_W = PAGE_W - 2 * MARGIN; // 9306

// ---------- rich text: **bold**, _{sub}, ^{sup}, `code`
function runs(text, base = {}) {
  const out = [];
  let bold = false, i = 0, buf = "";
  const flush = (extra = {}) => {
    if (buf) out.push(new TextRun({ text: buf, bold, font: FONT, ...base, ...extra }));
    buf = "";
  };
  while (i < text.length) {
    if (text.startsWith("**", i)) { flush(); bold = !bold; i += 2; continue; }
    if ((text[i] === "_" || text[i] === "^") && text[i + 1] === "{") {
      const close = text.indexOf("}", i);
      const inner = text.slice(i + 2, close);
      flush();
      out.push(new TextRun({ text: inner, bold, font: FONT, ...base,
        ...(text[i] === "_" ? { subScript: true } : { superScript: true }) }));
      i = close + 1; continue;
    }
    if (text[i] === "`") {
      const close = text.indexOf("`", i + 1);
      flush();
      out.push(new TextRun({ text: text.slice(i + 1, close), font: MONO, bold, ...base,
        size: base.size ? base.size - 2 : 19 }));
      i = close + 1; continue;
    }
    buf += text[i]; i++;
  }
  flush();
  return out;
}

const P = (text, opts = {}) => new Paragraph({ children: runs(text, opts.run || {}), spacing: { after: 120, line: 288 },
  alignment: opts.align, ...(opts.para || {}) });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(t) });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(t) });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: runs(t) });
const bullet = (t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: runs(t),
  spacing: { after: 80, line: 276 } });
const caption = (t) => new Paragraph({ children: runs(t, { bold: true, size: 20, color: ACCENT }),
  spacing: { before: 160, after: 80 }, keepNext: true });
const note = (t) => new Paragraph({ children: runs(t, { size: 18, color: GREY, italics: true }),
  spacing: { before: 60, after: 160 } });

// ---------- tables
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(content, width, opts = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((t) => new Paragraph({
    children: runs(String(t), { size: opts.size || 19, ...(opts.run || {}) }),
    alignment: opts.align || AlignmentType.LEFT, spacing: { after: 0, line: 252 }, keepNext: opts.keepNext,
  }));
  return new TableCell({
    children: paras, width: { size: width, type: WidthType.DXA }, borders,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    columnSpan: opts.span,
  });
}

function table(widths, header, rows, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  if (header) {
    trs.push(new TableRow({ tableHeader: true, children: header.map((h, j) =>
      cell(`**${h}**`, widths[j], { fill: "D9E2F3", align: j > 0 && opts.numHeader ? AlignmentType.CENTER : undefined })) }));
  }
  rows.forEach((r, ri) => {
    const keep = !opts.allowBreak && ri < rows.length - 1;
    trs.push(new TableRow({ cantSplit: true, children: r.map((c, j) => {
      const isNum = opts.numCols && opts.numCols.includes(j);
      const mono = opts.monoCols && opts.monoCols.includes(j);
      return cell(c, widths[j], {
        keepNext: keep,
        align: isNum ? AlignmentType.RIGHT : AlignmentType.LEFT,
        fill: opts.firstColFill && j === 0 ? "F2F2F2" : undefined,
        size: mono ? 16 : undefined,
        run: mono ? { font: MONO } : undefined,
      });
    }) }));
  });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: trs,
    layout: TableLayoutType.FIXED });
}

function callout(lines) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "9DC3E6" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "9DC3E6" },
        left: { style: BorderStyle.SINGLE, size: 24, color: ACCENT }, right: { style: BorderStyle.SINGLE, size: 4, color: "9DC3E6" } },
      shading: { fill: "EAF1FB", type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: lines.map((t, k) => k === 0
        ? new Paragraph({ children: runs(t, { bold: true, color: ACCENT }), spacing: { after: 80 } })
        : new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: runs(t), spacing: { after: 60, line: 264 } })),
    })] })],
  });
}

const spacer = () => new Paragraph({ children: [], spacing: { after: 60 } });

// ---------- values
const n = V("B4"), mean = V("B5"), sd = V("B6"), se = V("B7"), tcrit = V("B8"), lo = V("B9"), hi = V("B10");
const t1 = V("B12"), p1 = V("B13"), half = V("B15");
const n1 = V("B21"), n2 = V("C21"), m1 = V("B22"), m2 = V("C22"), v1 = V("B23"), v2 = V("C23");
const diff = V("B26"), seW = V("B27"), tW = V("B28"), dfW = V("B29"), pW = V("B30"), pWt = V("B31");
const sp2 = V("B35"), seP = V("B36"), tP = V("B37"), dfP = V("B38"), pP = V("B39");
const r2 = V("B42"), r2a = V("B43"), sey = V("B44"), Fst = V("B45"), nobs = V("B46"), rdf = V("B47"), ssr = V("B48"), sse = V("B49"), sigF = V("B50");
const median = V("B19", "Part A"), skew = V("B24", "Part A");
const ratio = v1 / v2;

// ---------- content
const children = [];

children.push(new Paragraph({ children: runs("Group 8 · Final Project · Wine Prices and Regional Price Premia",
  { color: ACCENT, bold: true, size: 20 }), spacing: { after: 60 } }));
children.push(new Paragraph({ children: runs("Part B1–B2: from one mean to a regression on the Napa dummy", { bold: true, size: 36 }),
  spacing: { after: 120 } }));
children.push(P("Section prepared by **Trung** (B1 and B2). Every statistic below is computed with formulas in **Group_8.xlsx**: sheet " +
  "**Part B** (cells A3:H59 and K41:L46) and the helper columns **U–V** on sheet **Data**. Cell references are given in brackets so " +
  "each number can be traced.", { run: { color: GREY, size: 20 } }));

children.push(H1("Overview"));
children.push(P("B1 and B2 are the first two rungs of the course's model ladder. B1 fits the simplest model, " +
  "**Price = β_{0} + ε**, which has no explanatory variable. B2 adds one dummy, **Price = β_{0} + β_{1}·Napa + ε**, where Napa = 1 " +
  `for the ${n1} Napa Valley wines and 0 for the other ${n2}. Together they give a first, raw answer to the association's question: ` +
  "how much more does a Napa bottle cost on average, before holding quality, age, volume or grape variety constant?"));
children.push(callout([
  "Key results",
  `The average bottle costs **${money(mean)}** (95% CI ${money(lo)} – ${money(hi)}), significantly more than $25 (t = ${fx(t1, 2)}, p = ${fx(p1, 5)}).`,
  `Napa wines average **${money(m1)}** against **${money(m2)}** for all other wines: a raw gap of **${money(diff)} per bottle**, highly ` +
    `significant (Welch t = ${fx(tW, 2)}, p ≈ ${sci(pW)}).`,
  `Regressing Price on the Napa dummy gives exactly the same gap (slope = ${fx(diff, 2)}), but its t statistic (${fx(tP, 2)}) is that of the ` +
    "pooled, equal-variance t-test, not Welch's.",
  "Price is strongly right-skewed, so the Wilcoxon rank-sum (Mann–Whitney U) test is the safer check of the comparison.",
]));

// ===================== B1
children.push(H1("B1 · Continuous Y, no X: Price = β_{0} + ε"));
children.push(H3("What I did"));
children.push(P("On sheet Part B (B4:B15) I computed the sample size, mean, standard deviation and standard error of Price with " +
  "`COUNT`, `AVERAGE`, `STDEV.S` and s/√n; built the 95% confidence interval with the t critical value from `T.INV.2T`; and ran the " +
  "one-sample t-test of H_{0}: μ = 25 with `T.DIST.2T`. `CONFIDENCE.T` is used as an independent check of the interval's half-width. " +
  "All ranges point to the Price column, `Data!$B$2:$B$807`."));

children.push(caption("Table B1.1 · Mean of Price and its 95% confidence interval"));
children.push(table([2700, 1300, 800, 4506], ["Statistic", "Value", "Cell", "Excel formula"], [
  ["Number of wines, n", String(n), "B4", "`=COUNT(Data!$B$2:$B$807)`"],
  ["Mean price, x̄", money(mean), "B5", "`=AVERAGE(Data!$B$2:$B$807)`"],
  ["Standard deviation, s", money(sd), "B6", "`=STDEV.S(Data!$B$2:$B$807)`"],
  ["Standard error, s/√n", money(se), "B7", "`=B6/SQRT(B4)`"],
  [`t critical (α = 0.05, df = ${n - 1})`, fx(tcrit, 3), "B8", "`=T.INV.2T(0.05,B4-1)`"],
  ["95% CI, lower limit", money(lo), "B9", "`=B5-B8*B7`"],
  ["95% CI, upper limit", money(hi), "B10", "`=B5+B8*B7`"],
  ["Half-width check", money(half), "B15", "`=CONFIDENCE.T(0.05,B6,B4)`"],
], { numCols: [1] }));
children.push(spacer());
children.push(P(`**Meaning.** The average bottle in the sample costs ${money(mean)}, but individual prices vary widely: the standard ` +
  `deviation (${money(sd)}) is larger than the mean itself. Because the mean is based on ${n} wines, its standard error is only ` +
  `${money(se)}, so the 95% confidence interval for the population mean price is narrow: **${money(lo)} to ${money(hi)}**. Read ` +
  "correctly, 95% of intervals built this way would contain the true mean; it does **not** say that 95% of wines cost between " +
  `$27 and $32 (the median bottle costs only ${money0(median)}, Part A!B19). CONFIDENCE.T returns the same half-width ` +
  `(${fx(half, 2)} = ${fx(tcrit, 3)} × ${fx(se, 3)}), which confirms the interval.`));

children.push(caption("Table B1.2 · One-sample t-test of H_{0}: μ = 25"));
children.push(table([1500, 6906, 900], ["Step", "Content", "Cell"], [
  ["1. Hypotheses", "H_{0}: μ = 25 (the true mean price is $25) against H_{1}: μ ≠ 25 (two-sided). Significance level α = 0.05.", "B11"],
  ["2. Test statistic", `One-sample t = (x̄ − μ_{0}) / (s/√n) = (${fx(mean, 2)} − 25) / ${fx(se, 3)} = **${fx(t1, 2)}**, with df = n − 1 = ${n - 1}. ` +
    "Excel: `=(B5-B11)/B7`", "B12"],
  ["3. p-value", `p = T.DIST.2T(${fx(t1, 2)}, ${n - 1}) = **${fx(p1, 5)}** < α = 0.05. Equivalently |t| = ${fx(t1, 2)} > t critical ` +
    `${fx(tcrit, 3)}. Excel: \`=T.DIST.2T(ABS(B12),B4-1)\``, "B13"],
  ["4. Decision", "**Reject H_{0}** (`=IF(B13<0.05,\"Reject H0\",\"Do not reject H0\")`). The mean retail price of these California " +
    "wines is significantly higher than $25 per bottle.", "B14"],
], { firstColFill: true }));
children.push(spacer());
children.push(P(`**Meaning.** The sample mean lies ${fx(t1, 2)} standard errors above $25. If the true mean really were $25, a t ` +
  "statistic this far from zero would occur in only about 15 of every 100,000 samples, so the data are not consistent with H_{0}. " +
  `The test agrees with the confidence interval: $25 lies outside [${money(lo)}; ${money(hi)}], and a two-sided test at the 5% level ` +
  "rejects exactly when the hypothesised value lies outside the 95% interval. With n = 806 the t-test is reliable even though Price " +
  "is skewed, because the central limit theorem makes the sampling distribution of x̄ approximately normal."));

children.push(H3("Why this mean is the estimate of β_{0}"));
children.push(P("In **Price = β_{0} + ε** there is no X, so the model predicts the same value β_{0} for every wine, and ε = Price − β_{0} " +
  "is each wine's error. Ordinary least squares chooses the estimate b_{0} that minimises the sum of squared errors " +
  "SSE(b) = Σ(Price_{i} − b)^{2}. Setting the derivative to zero gives −2·Σ(Price_{i} − b) = 0, so b = ΣPrice_{i} / n = x̄. " +
  `The sample mean is therefore the least-squares estimate of β_{0} (b_{0} = x̄ = ${money(mean)}), its standard error ` +
  `(${money(se)}) is the standard error of b_{0}, and the one-sample t-test above is simply the t-test of this intercept.`));
children.push(P(`**One-sentence answer:** with no X, least squares picks the single number that minimises Σ(Price − β_{0})^{2}, and that ` +
  `number is the sample mean, so b_{0} = x̄ = ${money(mean)}.`));

// ===================== B2
children.push(H1("B2 · Continuous Y, one dummy X: Price = β_{0} + β_{1}·Napa + ε"));
children.push(H3("What I did"));
children.push(P("On sheet Data I added two helper columns, filled down to row 807: **U** “Price if Napa” `=IF(F2=1,B2,\"\")` and " +
  "**V** “Price if not Napa” `=IF(F2=0,B2,\"\")`. Each wine's price appears in exactly one of them, and because `COUNT`, `AVERAGE` " +
  "and `VAR.S` ignore the empty text cells, the two columns give each group's statistics directly (Part B, B21:C23). I then ran " +
  "Welch's t-test with formulas and with `T.TEST` (B25:B32), regressed Price on the Napa dummy with `LINEST` (K42:L46, laid out " +
  "as regression output in A41:H54), and computed the pooled t-test (B34:B39) for comparison. Check cells B56:B58 confirm the " +
  "equalities discussed below."));

children.push(caption("Table B2.1 · Price by group"));
children.push(table([3000, 1800, 1800, 2706], ["Statistic", "Napa (Napa = 1)", "Other (Napa = 0)", "Cells · formula"], [
  ["Number of wines, n", String(n1), String(n2), "B21, C21 · `COUNT`"],
  ["Mean price", money(m1), money(m2), "B22, C22 · `AVERAGE`"],
  ["Variance of price ($^{2})", num(v1, 2), num(v2, 2), "B23, C23 · `VAR.S`"],
], { numCols: [1, 2] }));
children.push(spacer());
children.push(P(`Napa wines are on average almost twice as expensive, and their prices are far more dispersed: the Napa variance is ` +
  `about ${Math.round(ratio)} times that of the other wines (${num(v1, 2)} vs ${num(v2, 2)}). Because the two variances are so ` +
  "different, pooling them is not justified (Lecture 10), which is why the project asks for Welch's test."));

children.push(H2("B2.1 · Welch two-sample t-test (unequal variances)"));
children.push(caption("Table B2.2 · Welch t-test: Napa vs. all other wines"));
children.push(table([1500, 6906, 900], ["Step", "Content", "Cell"], [
  ["1. Hypotheses", "H_{0}: μ_{Napa} = μ_{Other} (μ_{Napa} − μ_{Other} = 0) against H_{1}: μ_{Napa} ≠ μ_{Other} (two-sided); α = 0.05.", "A25"],
  ["2. Test statistic", `Welch t = (x̄_{1} − x̄_{2}) / √(s_{1}^{2}/n_{1} + s_{2}^{2}/n_{2}) = ${fx(diff, 2)} / ${fx(seW, 3)} = **${fx(tW, 2)}**, ` +
    `with Welch–Satterthwaite df = ${fx(dfW, 1)}. Excel: difference \`=B22-C22\`, SE \`=SQRT(B23/B21+C23/C21)\`, t \`=B26/B27\`, df in B29.`,
    "B26–B29"],
  ["3. p-value", `p = T.DIST.2T(${fx(tW, 2)}, ${fx(dfW, 1)}) = **${sci(pW)}** (B30). Excel's built-in Welch test ` +
    `\`=T.TEST(Data!$U$2:$U$807,Data!$V$2:$V$807,2,3)\` (2 tails, type 3 = two-sample unequal variance) gives ${sci(pWt)} (B31). ` +
    "Both are far below α = 0.05.", "B30–B31"],
  ["4. Decision", "**Reject H_{0}** (B32). The mean price of Napa Valley wines differs significantly from that of other California wines.", "B32"],
], { firstColFill: true }));
children.push(spacer());
children.push(P(`**Meaning.** On average a Napa bottle costs **${money(diff)} more** than a non-Napa bottle, and a gap this large is ` +
  "practically impossible to obtain by sampling chance if the two population means were equal. It is, however, a **raw** difference: " +
  "it compares Napa wines with all others without holding score, age, production volume or grape variety constant, so it is not yet " +
  "the price premium of the Napa name that the association asked about (B5 and Part C address that). The small gap between B30 and " +
  `B31 arises because T.DIST.2T truncates the degrees of freedom to ${Math.floor(dfW)} while T.TEST uses ${fx(dfW, 1)}; both lead ` +
  "to the same decision."));

children.push(H2("B2.2 · Regression of Price on the Napa dummy"));
children.push(P("The regression was estimated with the array formula `{=LINEST(Data!$B$2:$B$807,Data!$F$2:$F$807,TRUE,TRUE)}` in " +
  "K42:L46 (confirmed with Ctrl+Shift+Enter). Cells A41:H54 present the LINEST output in the same layout as the ToolPak's " +
  "Regression output: t = coefficient / SE, p = T.DIST.2T(|t|, 804) and CI = coefficient ± T.INV.2T(0.05, 804) × SE."));
children.push(caption("Table B2.3 · Regression statistics (B42:B50)"));
children.push(table([3053, 1600, 3053, 1600], ["Statistic", "Value", "Statistic", "Value"], [
  ["R Square", fx(r2, 4), "Observations, n", String(nobs)],
  ["Adjusted R Square", fx(r2a, 4), "Residual df", String(rdf)],
  ["Standard error of regression", fx(sey, 2), "SS Regression (SSR)", num(ssr, 2)],
  ["F statistic", fx(Fst, 2), "SS Residual (SSE)", num(sse, 2)],
], { numCols: [1, 3] }));
children.push(note(`Significance F (B50) = F.DIST.RT(${fx(Fst, 2)}, 1, ${rdf}) = ${sci(sigF)}: the same p-value as the Napa ` +
  "coefficient's t-test, because with one X the F statistic equals t^{2} (8.71^{2} = 75.87)."));
children.push(caption("Table B2.4 · Coefficients (A52:H54)"));
const coefRow = (r, name) => [name, fx(V("B" + r), 4), fx(V("C" + r), 4), fx(V("D" + r), 2), sci(V("E" + r)),
  fx(V("F" + r), 2), fx(V("G" + r), 2)];
children.push(table([1606, 1250, 1250, 1100, 1500, 1300, 1300],
  ["Variable", "Coefficient", "Std. Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"],
  [coefRow(53, "Intercept"), coefRow(54, "Napa valley")], { numCols: [1, 2, 3, 4, 5, 6] }));
children.push(spacer());
children.push(P(`Fitted model: **Predicted Price = ${fx(V("B53"), 2)} + ${fx(V("B54"), 2)} × Napa.**`, { align: AlignmentType.CENTER }));

children.push(caption("Table B2.5 · The slope equals the difference between the group means"));
children.push(table([2200, 1700, 3406, 2000], ["Regression", "Value", "Group statistics", "Check cell"], [
  ["Intercept b_{0} (B53)", fx(V("B53"), 4), `Mean price of other wines (C22) = ${fx(m2, 4)}`, `B57 = B53 − C22 = ${V("B57")}`],
  ["Slope b_{1} (B54)", fx(V("B54"), 4), `Mean Napa − mean other (B26) = ${fx(m1, 4)} − ${fx(m2, 4)} = ${fx(diff, 4)}`, `B56 = B54 − B26 = ${V("B56")}`],
], { numCols: [1] }));
children.push(spacer());
children.push(P("**Meaning.** For a non-Napa wine (Napa = 0) the model predicts b_{0}; for a Napa wine (Napa = 1) it predicts " +
  "b_{0} + b_{1}. Least squares makes each group's prediction equal to that group's mean, exactly as the mean minimised the squared " +
  `errors in B1. So b_{0} = mean of other wines = ${money(m2)}, and b_{0} + b_{1} = mean of Napa wines = ${money(m1)}, which forces ` +
  `b_{1} = ${money(diff)}: the difference in means. Both check cells are exactly 0. R^{2} = ${fx(r2, 3)} shows that the Napa label ` +
  "alone explains only about 9% of the variation in prices; most of the variation comes from other characteristics."));

children.push(H2("B2.3 · Which two-sample t-test does the regression reproduce?"));
children.push(caption("Table B2.6 · Same difference, different standard errors"));
children.push(table([2406, 2300, 2300, 2300], ["", "Welch t-test", "Regression slope (LINEST)", "Pooled t-test"], [
  ["Difference in means ($)", `${fx(diff, 2)} (B26)`, `${fx(V("B54"), 2)} (B54)`, `${fx(diff, 2)} (B26)`],
  ["Standard error", `${fx(seW, 3)} (B27)`, `${fx(V("C54"), 3)} (C54)`, `${fx(seP, 3)} (B36)`],
  ["t statistic", `${fx(tW, 2)} (B28)`, `${fx(V("D54"), 2)} (D54)`, `${fx(tP, 2)} (B37)`],
  ["Degrees of freedom", `${fx(dfW, 1)} (B29)`, `${rdf} (B47)`, `${dfP} (B38)`],
  ["p-value (two-sided)", `${sci(pW)} (B30)`, `${sci(V("E54"))} (E54)`, `${sci(pP)} (B39)`],
], { numCols: [1, 2, 3], numHeader: true, firstColFill: true }));
children.push(spacer());
children.push(P("The regression reproduces exactly the **pooled (equal-variance) two-sample t-test**: Excel's `T.TEST(…,2,2)` or the " +
  "ToolPak's “t-Test: Two-Sample Assuming Equal Variances”. The pooled variance s_{p}^{2} = [(n_{1} − 1)s_{1}^{2} + (n_{2} − 1)s_{2}^{2}] " +
  `/ (n_{1} + n_{2} − 2) = ${num(sp2, 2)} (B35) gives SE = s_{p}·√(1/n_{1} + 1/n_{2}) = ${fx(seP, 3)} (B36), identical to the slope's ` +
  `standard error, so the t statistics (${fx(tP, 2)}), degrees of freedom (${rdf}) and p-values coincide; check cell B58 = D54 − B37 = ${V("B58")}.`));
children.push(P("**Why.** OLS assumes that every wine's error has the same variance σ^{2} (homoskedasticity) and estimates that single " +
  `variance from all residuals together: SSE / (n − 2) = ${num(sse, 2)} / ${rdf} = ${num(sp2, 2)}, which is exactly s_{p}^{2}; ` +
  `equivalently, the standard error of the regression (${fx(sey, 2)}, B44) squared is ${num(sey * sey, 2)}. Pooling the two groups' ` +
  "variances is precisely the equal-variance assumption of the pooled t-test. Welch's test instead lets each group keep its own " +
  `variance. Here the Napa variance is about ${Math.round(ratio)} times larger, so the pooled standard error understates the noise ` +
  `and Welch's (${fx(seW, 3)}) is the more trustworthy one. In this case both tests reject H_{0} by a wide margin, so the ` +
  "conclusion does not change."));

children.push(H2("B2.4 · A safer, rank-based check"));
children.push(P(`The histogram of Price (A5) is strongly right-skewed: skewness ${fx(skew, 2)} (Part A!B24), mean ${money(mean)} above ` +
  `the median ${money0(median)} (Part A!B19), and a few bottles priced up to ${money0(V("B23", "Part A"))} (Part A!B23). Because t-tests compare means, a handful ` +
  "of extreme prices can pull the result. The safer rank-based test from class is the **Wilcoxon rank-sum test**, also called the " +
  "**Mann–Whitney U test**. It pools all 806 prices, replaces each by its rank (1 = cheapest, 806 = most expensive) and tests whether " +
  "Napa wines tend to have higher ranks than the others. An extreme price simply receives the top rank, so outliers cannot dominate " +
  "the comparison. As the project allows, the test is named but not run."));

children.push(H1("What B1–B2 mean for the Napa question"));
[
  `A typical bottle in the data costs about ${money(mean)} on average (95% CI ${money(lo)} – ${money(hi)}), with very large ` +
    "dispersion and a long right tail.",
  `Napa Valley wines sell for **${money(diff)} more per bottle** on average than all other wines, and the difference is highly ` +
    "significant under both Welch's and the pooled t-test.",
  "This is a raw gap, not yet a premium for the Napa name: Napa wines may also differ in score, age, volume and variety. B5 and " +
    "Part C hold these characteristics constant to isolate the label's own effect.",
  "Statistically, B1 and B2 show that familiar tests are special cases of regression: the one-sample mean is the intercept of " +
    "Price = β_{0} + ε, and the two-group comparison is the slope on a dummy, with the regression reproducing the pooled t-test.",
].forEach((t) => children.push(bullet(t)));

// ===================== Appendix
children.push(H1("Appendix · Cell map of B1–B2 in Group_8.xlsx"));
children.push(P("Every number in this section can be traced to the cells below (sheet Part B unless stated).", { run: { color: GREY, size: 20 } }));
const label = (ref) => String(cells[`Part B!A${ref.replace(/\D/g, "")}`]?.f ?? "");
const val = (v) => (typeof v === "number" ? (Math.abs(v) < 1e-4 && v !== 0 ? sci(v) : Math.abs(v) < 1e-3 && v !== 0 ? fx(v, 6) : Number.isInteger(v) ? String(v) : fx(v, 4)) : String(v ?? ""));
const mapRows = [];
const addRow = (ref, lbl) => mapRows.push([ref, lbl ?? label(ref), "`" + F(ref) + "`", val(V(ref))]);
mapRows.push(["Data!U2", "Price if Napa (filled to row 807)", "`=IF(F2=1,B2,\"\")`", "—"]);
mapRows.push(["Data!V2", "Price if not Napa (filled to row 807)", "`=IF(F2=0,B2,\"\")`", "—"]);
["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B12", "B13", "B14", "B15"].forEach((r) => addRow(r));
[["B21", "n, Napa"], ["C21", "n, Other"], ["B22", "Mean price, Napa"], ["C22", "Mean price, Other"], ["B23", "Variance, Napa"],
  ["C23", "Variance, Other"]].forEach(([r, l]) => addRow(r, l));
["B26", "B27", "B28", "B29", "B30", "B31", "B32", "B35", "B36", "B37", "B38", "B39"].forEach((r) => addRow(r));
mapRows.push(["K42:L46", "LINEST output (5 × 2 array)", "`" + F("K42") + "`", `slope ${fx(V("K42"), 4)}, intercept ${fx(V("L42"), 4)}`]);
[["B53", "Intercept coefficient"], ["B54", "Napa coefficient"], ["D54", "t Stat, Napa"], ["E54", "P-value, Napa"]].forEach(([r, l]) => addRow(r, l));
["B56", "B57", "B58"].forEach((r) => addRow(r));
children.push(table([1050, 2900, 3956, 1400], ["Cell", "Content", "Formula", "Value"], mapRows, { monoCols: [2], numCols: [3], allowBreak: true }));

// ---------- document
const doc = new Document({
  creator: "Group 8",
  title: "Group 8 – Part B1–B2",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: FONT, color: ACCENT }, paragraph: { spacing: { before: 320, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, font: FONT, color: "222222" }, paragraph: { spacing: { before: 260, after: 100 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: FONT, color: GREY }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1200, bottom: 1200, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "Group 8 · Part B1–B2 · page ", font: FONT, size: 16, color: GREY }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children,
  }],
});

const out = process.argv[2];
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out, buf.length, "bytes"); });
