# B1–B2 · Q&A practice (individual questions, 25 points)

**How to practise:** cover the answer, say yours out loud in 30–60 seconds, then compare. The lecturer grades whether you *understand* the numbers, not whether you remember them. Always answer as: **what the number is → what it means for wine prices → why the method is right**.

Difficulty: ★ basic · ★★ standard · ★★★ curveball.
Cells refer to the **shared online workbook, sheet Part B, as of 18 Oct 2026** (B1 A1:B14 · B2.1 A16:C28 · B2.2 A30:E43 · B2.3 A45:C53 · B2.4 A55:A58 + F45:G58). Numbers marked † were checked outside the workbook (Python) and are **not** in the Excel file. Say so if you use them.

---

## 0. Your 30-second summaries

**B1.** "With no X, the model Price = β₀ + ε predicts one number for every wine, and least squares makes that number the sample mean, $29.58. Its standard error is only $1.20 because n = 806, so the 95% CI is $27.22–$31.94. Testing μ = 25 gives t = 3.81, p = 0.00015, so the average price is significantly above $25, consistent with 25 lying outside the CI."

**B2.** "Napa wines average $43.06 against $22.13 for the rest: a raw gap of $20.93. Welch's test (t = 6.72, p ≈ 9×10⁻¹¹) rejects equal means; we use Welch because the Napa variance is about 16 times larger. Regressing Price on the Napa dummy gives exactly that gap as the slope, but its t (8.71) is the pooled t-test's, because OLS assumes one common error variance. Price is very skewed, so the safer check is Mann–Whitney U, which also rejects strongly (z = 12.34). This gap is not yet the Napa *premium*: Napa wines also score higher and are more often Cabernet, which B5 controls for."

---

## 1. Explain your own cells ★

**Q1. Walk me through B1. Where does each number come from?**
B3 `=COUNT(Data!$B$2:$B$807)` = 806 wines; B4 `AVERAGE` = 29.58; B5 `STDEV.S` = 34.15; B6 = s/√n = 1.203; B7 `T.INV.2T(0.05, n−1)` = 1.963; B8:B9 = x̄ ∓ t*·SE = [27.22; 31.94]; B11 t = (x̄ − 25)/SE = 3.81; B12 `T.DIST.2T(|t|, 805)` = 0.000151; B13 `IF` gives the decision; B14 `CONFIDENCE.T` re-checks the half-width (2.36 = 1.963 × 1.203).

**Q2. How did you get each group's statistics without sorting the data?**
Conditional formulas on the Napa dummy (Data column F): `COUNTIF(F,1)` = 287, `AVERAGEIF(F,1,B)` = 43.06, and the variance as `SUMPRODUCT((F=1)*(B−mean)^2)/(n−1)` = 2689.99. The same with 0 for the other 519 wines.

**Q3. What do the "check" cells show?**
E41 = slope − (mean Napa − mean Other) = 0; E42 = intercept − mean Other = 0 (B2.2). C52 = slope t − pooled t = 0; C53 = SE(regression)² − pooled variance = 0 (B2.3). They prove the equalities the questions ask about, using the workbook's own numbers.

**Q4. Why is `ABS()` inside `T.DIST.2T`?**
`T.DIST.2T` only accepts a non-negative t. With `ABS` the formula still works if t is negative (e.g. if the groups were coded the other way round).

---

## 2. B1 concepts ★–★★

**Q5. What is the difference between the standard deviation (34.15) and the standard error (1.20)?** ★
The SD measures how much **individual wines** differ from the mean. The SE = s/√n measures how much the **sample mean** would vary from sample to sample. Averaging 806 wines cancels most of the noise, so the SE is about √806 ≈ 28 times smaller.

**Q6. Why STDEV.S and not STDEV.P?** ★
The 806 wines are a sample of the population. STDEV.S divides by n − 1, which corrects for having estimated the mean from the same data (Bessel's correction, giving an unbiased variance). STDEV.P divides by n and is for a whole population.

**Q7. Why a t critical value (1.963) and not z = 1.96?** ★
σ is unknown and we estimate it with s, which adds uncertainty, so the statistic follows a t distribution with n − 1 = 805 df. With 805 df, t is almost normal, which is why 1.963 ≈ 1.96.

**Q8. Why n − 1 degrees of freedom?** ★★
One parameter (the mean) is estimated before computing s. Once x̄ is fixed, only n − 1 deviations are free to vary.

**Q9. Interpret the 95% CI [27.22; 31.94] correctly.** ★★
If we drew many samples and built an interval this way each time, 95% of those intervals would contain the true mean price μ. It is **not** a range holding 95% of wines; most bottles are outside it (the median is $22). Nor is it "a 95% probability that μ is in this particular interval", since μ is fixed.

**Q10. What does p = 0.00015 mean?** ★★
If the true mean really were $25, a sample mean at least 3.81 SEs away from 25 would occur in about 0.015% of samples. It is **not** the probability that H₀ is true.

**Q11. Why do the CI and the test agree?** ★★
Both use the same estimate, SE and t*. A two-sided test at α = 0.05 rejects exactly when the hypothesised value lies outside the 95% CI. 25 is outside [27.22; 31.94], so we reject.

**Q12. What are Type I and Type II errors here?** ★★
Type I: rejecting μ = 25 when the true mean really is 25; its probability is α = 0.05. Type II: failing to reject when the true mean is not 25. With n = 806 and t = 3.81, a Type II error is not the concern here.

**Q13. Why is the mean the estimate of β₀ in Price = β₀ + ε?** ★★
With no X the model predicts the same value b for every wine. OLS chooses b to minimise Σ(Priceᵢ − b)². The derivative −2Σ(Priceᵢ − b) = 0 gives b = x̄. So b₀ = $29.58, its SE is the SE of the mean, and the one-sample t-test is the t-test of the intercept.

**Q14. Price is very skewed (skewness 7.87). Is the t-test still valid?** ★★
Largely yes. The t-test is about the **sample mean**, and by the central limit theorem the mean of 806 observations is approximately normal even if individual prices are not. With a small sample (say 15 wines) the skew would be a real problem.

**Q15. Would 25 be rejected at α = 0.01?** ★★★
Yes. p = 0.00015 < 0.01. Equivalently the 99% CI, x̄ ± 2.582 × 1.203 = [26.47; 32.69]†, still excludes 25.

**Q16. What if H₁ were "μ > 25" (one-sided)?** ★★★
Same t = 3.81, but only the right tail counts: p = 0.000075† (half the two-sided p). The decision does not change. The project asked for the two-sided ≠ test.

**Q17. What happens to the SE if we had 4 times as many wines?** ★★★
SE = s/√n, so it halves (to ≈ 0.60†) and the CI becomes about half as wide. Larger n buys precision, but it does not fix bias in how the sample was collected.

---

## 3. B2.1 · Welch two-sample t-test ★–★★

**Q18. State the hypotheses for B2.1.** ★
H₀: μ_Napa = μ_Other (difference 0) vs H₁: μ_Napa ≠ μ_Other, α = 0.05, two-sided.

**Q19. Why Welch and not the pooled t-test?** ★★
The two variances are very different: 2689.99 vs 171.31, a ratio of about 15.7 (SDs $51.87 vs $13.09†). Pooling assumes equal variances. Welch keeps each group's own variance in SE = √(s₁²/n₁ + s₂²/n₂). Lecture 10: when in doubt, use Welch.

**Q20. How could you formally test that the variances differ?** ★★★
With the F-test for two variances from Lecture 10: F = s²_large/s²_small = 15.70 with (286, 518) df. p ≈ 10⁻¹⁵⁰†, so equal variances are clearly rejected. In Excel, `F.TEST(range1, range2)`. Note: the F-test is itself sensitive to non-normality.

**Q21. Why is Welch's df 306.3, not 804?** ★★★
The Satterthwaite formula weights each group by its share of the SE. Almost all of the SE comes from the Napa term (2690/287 ≈ 9.4 vs 171/519 ≈ 0.3), so the effective df is close to Napa's own n − 1 = 286, not n₁ + n₂ − 2.

**Q22. Why do B26 (8.886E-11) and B27 (8.874E-11) differ slightly?** ★★★
`T.DIST.2T` truncates the df to the whole number 306, while `T.TEST(…,2,3)` uses the exact 306.3. The difference is in the third digit, and the decision is the same.

**Q23. What do the arguments of `T.TEST(range1, range2, 2, 3)` mean?** ★
2 = two-tailed. Type 1 = paired, 2 = two-sample equal variances (pooled), 3 = two-sample unequal variances (Welch).

**Q24. Interpret the result in wine terms.** ★
On average a Napa bottle costs $20.93 more than a non-Napa bottle, and that gap is far too large to be sampling noise (p ≈ 9×10⁻¹¹).

**Q25. Does this prove the "Napa Valley" label raises prices?** ★★ *(very likely question)*
No. This is a raw comparison of observational data, not a randomised experiment. Napa wines differ in other ways: higher average score (88.5 vs 85.8), older (3.06 vs 2.68 years), smaller production (3.7k vs 7.0k cases), and more often Cabernet (48% vs 26%). Part of the $20.93 is quality and variety, not the label. B5 holds these constant. In the team's log model (B5.2) the Napa effect falls to about +54% vs Generic California, other things equal. That is the "premium" the client asked about.

**Q26. Is a $20.93 gap important, or just significant?** ★★★
Both. It is almost as large as the average non-Napa bottle ($22.13). A 95% CI for the difference is roughly $14.80–$27.06† (Welch). In Excel: `=diff ± T.INV.2T(0.05, df)·SE`. Lecture 9: always read the size in real units before the p-value.

---

## 4. B2.2 · Regression on the Napa dummy ★–★★

**Q27. How do you read the LINEST output block?** ★
Row 1: slope b₁ (20.93) | intercept b₀ (22.13). Row 2: their SEs (2.403 | 1.434). Row 3: R² (0.086) | SE of regression (32.67). Row 4: F (75.87) | residual df (804). Row 5: SSR (80,972.66) | SSE (858,075.60).

**Q28. Why is the slope exactly the difference in means?** ★★
For a non-Napa wine (D = 0) the model predicts b₀; for a Napa wine (D = 1) it predicts b₀ + b₁. OLS makes each group's prediction equal to that group's mean (the same logic as B1). So b₀ = $22.13 = mean of others and b₀ + b₁ = $43.06 = mean of Napa, which gives b₁ = $20.93.

**Q29. What does R² = 0.086 mean?** ★
The Napa dummy alone explains about 8.6% of the variation in prices (SSR/SST). Most of the variation comes from other characteristics: score, variety, age, volume.

**Q30. What is the intercept's t-test (t = 15.43, p ≈ 3×10⁻⁴⁷) testing? Is it useful?** ★★★
It tests H₀: β₀ = 0, i.e. that the mean price of non-Napa wines is $0. That is meaningless, since no bottle is free. Only the slope's test answers the question.

**Q31. How are F = 75.87 and the slope's t = 8.71 related?** ★★★
With a single X, F = t². 8.71² = 75.87, and both have the same p-value (1.7×10⁻¹⁷). From two X's onward they answer different questions (Lecture 11).

**Q32. What if the dummy were coded the other way (1 = not Napa)?** ★★★
Intercept = $43.06 (Napa mean), slope = −$20.93, same |t|, same p, same R². The coding only changes which group is the reference.

**Q33. What is the "standard error of regression" 32.67?** ★★
The typical size of a residual: √(SSE/(n − 2)) = √1067.26. Even knowing whether a wine is from Napa, predictions are typically off by about $33. That fits the low R².

---

## 5. B2.3 · Why the regression t equals the pooled t ★★–★★★

**Q34. Which two-sample test does the regression reproduce, and why?** ★★ *(core question)*
The **pooled (equal-variance) t-test**. OLS assumes every observation's error has the same variance σ² (homoskedasticity) and estimates that single variance from all residuals: SSE/(n − 2) = 32.67² = 1067.26, exactly the pooled variance s_p². So its slope SE (2.403) is the pooled SE, t = 8.71 and df = 804 = n₁ + n₂ − 2. Check cells C52 and C53 = 0.

**Q35. Which result do you trust more: t = 8.71 or t = 6.72?** ★★
Welch's 6.72. Lecture 10's worst case for pooling is exactly ours: the **larger** group (519 others) has the **smaller** variance. Pooling then understates the noise from the small, volatile Napa group, so the pooled SE (2.40) is too small and t is inflated. Here both tests reject H₀ by a wide margin, so the conclusion is the same.

**Q36. When would pooled and Welch give nearly the same answer?** ★★
When the two variances are similar, or the group sizes are equal. Then little is lost by always using Welch.

**Q37. Could you make the regression give Welch-like standard errors?** ★★★
Yes, in principle, with heteroskedasticity-robust ("White") standard errors, which don't assume a common variance. Excel's LINEST/ToolPak don't offer them; statistical software does. The point estimate (20.93) would not change, only its SE.

---

## 6. B2.4 · Rank-based test ★–★★★

**Q38. Which rank-based test, and why is it safer?** ★
The Wilcoxon rank-sum test (= Mann–Whitney U). It replaces each price by its rank (1 = cheapest, 806 = most expensive), so a $440 bottle counts only as "the top rank" and a few extreme prices cannot dominate. It needs no normality assumption.

**Q39. What does Mann–Whitney test, exactly?** ★★
H₀: Napa and other wines' prices come from the same distribution (a random Napa wine is equally likely to be pricier or cheaper than a random other wine). It compares ranks, not means.

**Q40. You computed it anyway. What did you get?** ★★
Part B F45:G58: rank sum of Napa wines R₁ = 154,829.5; U = R₁ − n₁(n₁+1)/2 = 113,501.5 vs 74,476.5 expected under H₀. With a tie-corrected SD of 3,162.35, z = 12.34 and p ≈ 5.5×10⁻³⁵, so we reject H₀. It agrees with both t-tests.

**Q41. Give an intuitive reading of U.** ★★★
U/(n₁n₂) = 113,501.5/148,953 ≈ 0.76†. Pick one Napa wine and one other wine at random: about 76% of the time the Napa wine is more expensive (ties count half).

**Q42. Why the tie correction?** ★★★
Many wines share the same price ($8, $16, …), and tied values get averaged ranks. That reduces the variance of U. The correction Σ(t³ − t) = 869,232 (G54) adjusts the SD so the z is not overstated.

**Q43. What is the rank-based equivalent for B3 (six regions)?** ★★★
The Kruskal–Wallis test, the rank version of one-way ANOVA (Course Review, "one X with 3+ groups").

**Q44. Do the extreme bottles drive the Napa gap?** ★★★
Partly. All 6 wines over $150 are from Napa. Dropping them, the Napa mean falls to $36.24 and the gap to $14.11†, but it is still highly significant (Welch t = 10.34†). The rank test, which ignores how extreme they are, also rejects strongly. So the gap is not just a few outliers.

**Q45. Why not just use ln(Price) for B2?** ★★★
B2 asks about the dollar price. On the log scale the means differ by 0.539†, i.e. Napa prices are about e^0.539 − 1 ≈ 71% higher† (in geometric means), and logs make the distribution far more symmetric (skewness 0.91). That's the approach of B3–B5.

---

## 7. Excel mechanics ★

**Q46. What do the `$` signs do in `Data!$B$2:$B$807`?** They lock the range so it does not shift when the formula is copied. Nothing to do with dollars.

**Q47. Why is LINEST entered as an array?** It returns a 5 × 2 block. In Excel 365/Online it spills automatically; in older Excel or LibreOffice you select the block and confirm with Ctrl+Shift+Enter.

**Q48. How does `T.TEST(IF(F=1,B), IF(F=0,B), 2, 3)` work without helper columns?** Each `IF` returns the prices of one group and FALSE elsewhere; T.TEST ignores the FALSE values. It returns the same p as a ToolPak t-test on two separate columns.

**Q49. The notes said to use the ToolPak. Why formulas?** The brief allows "formulas or the Data Analysis ToolPak". Formulas stay visible and recalculate if the data change, and the results are identical to the ToolPak's "t-Test: Two-Sample Assuming Unequal Variances" and "Regression".

**Q50. What does "8.89E-11" mean?** 8.89 × 10⁻¹¹ = 0.0000000000889. Excel switches to E-notation for very small numbers.

---

## 8. If you get stuck

- Go back to the model ladder: *"B1 is Y = β₀ + ε, B2 adds one dummy. Every test here is a special case of that regression."*
- Separate **size** from **significance**: "$20.93 per bottle, and very unlikely to be chance."
- Name the assumption: CLT (B1), unequal variances (Welch), common variance (OLS = pooled), no normality needed (Mann–Whitney).
- It is fine to say "I'd check that in the workbook". Then point to the cell.

---

## 9. Numbers cheat sheet

| | Value |
|---|---|
| n · mean · median · SD · skewness | 806 · 29.58 · 22 · 34.15 · 7.87 |
| SE · t* · 95% CI | 1.203 · 1.963 · [27.22; 31.94] |
| t (μ₀ = 25) · p | 3.81 · 0.000151 |
| Napa: n · mean · var · SD | 287 · 43.06 · 2689.99 · 51.87† |
| Other: n · mean · var · SD | 519 · 22.13 · 171.31 · 13.09† |
| Welch: diff · SE · t · df · p | 20.93 · 3.115 · 6.72 · 306.3 · 8.9×10⁻¹¹ |
| Regression: b₀ · b₁ · SE(b₁) · t · p · R² | 22.13 · 20.93 · 2.403 · 8.71 · 1.7×10⁻¹⁷ · 0.086 |
| Pooled: s_p² · SE · t · df | 1067.26 · 2.403 · 8.71 · 804 |
| Mann–Whitney: U · E[U] · z · p | 113,501.5 · 74,476.5 · 12.34 · 5.5×10⁻³⁵ |

## 10. Thuật ngữ / Glossary (EN → VI)

| English | Tiếng Việt |
|---|---|
| mean · median · standard deviation · variance | trung bình · trung vị · độ lệch chuẩn · phương sai |
| standard error | sai số chuẩn |
| confidence interval | khoảng tin cậy |
| null / alternative hypothesis | giả thuyết không (H₀) / giả thuyết đối (H₁) |
| test statistic · p-value · significance level | thống kê kiểm định · giá trị p · mức ý nghĩa |
| reject / fail to reject H₀ | bác bỏ / chưa đủ cơ sở bác bỏ H₀ |
| degrees of freedom | bậc tự do |
| Type I / Type II error | sai lầm loại I / loại II |
| dummy variable · intercept · slope | biến giả · hệ số chặn · hệ số góc |
| pooled variance · equal / unequal variances | phương sai gộp · phương sai bằng / không bằng nhau |
| homoskedasticity | phương sai sai số không đổi (đồng nhất) |
| residual · sum of squared errors | phần dư · tổng bình phương sai số |
| right-skewed · outlier | lệch phải · giá trị ngoại lai |
| rank-sum test | kiểm định tổng hạng |
| central limit theorem | định lý giới hạn trung tâm |
| confounding / omitted variable | biến gây nhiễu / biến bị bỏ sót |
