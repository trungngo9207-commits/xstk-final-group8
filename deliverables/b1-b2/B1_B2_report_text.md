# Group 8 · B1 & B2 — text for the report

Every number below is computed in sheet **Part B** of `Group_8.xlsx` (cell in brackets).

---

## B1 · Continuous Y, no X

| Statistic | Value | Cell / formula |
|---|---|---|
| n | 806 | B4 `=COUNT(Data!$B$2:$B$807)` |
| Mean Price (x̄) | $29.58 | B5 `=AVERAGE(...)` |
| Standard deviation (s) | $34.15 | B6 `=STDEV.S(...)` |
| Standard error (s/√n) | $1.20 | B7 `=B6/SQRT(B4)` |
| t critical (α = 0.05, df = 805) | 1.963 | B8 `=T.INV.2T(0.05,B4-1)` |
| 95% CI for μ | [$27.22 ; $31.94] | B9 `=B5-B8*B7`, B10 `=B5+B8*B7` |

**Hypothesis test**
1. H₀: μ = 25  vs  H₁: μ ≠ 25 (two-sided, α = 0.05).
2. Test statistic: one-sample t = (x̄ − 25) / (s/√n) = (29.58 − 25) / 1.20 = **3.81** (B12, `=(B5-B11)/B7`).
3. p-value = `T.DIST.2T(3.81, 805)` = **0.00015** (B13) < 0.05 (equivalently |t| = 3.81 > 1.963).
4. Decision: **reject H₀**. The average price of these California wines is significantly higher than $25.

**Interpretation.** The average bottle in the sample costs $29.58. We are 95% confident that the true mean price lies between $27.22 and $31.94 (95% of intervals built this way contain μ; it does *not* mean 95% of wines cost this much). $25 lies outside the interval, which agrees with the test.

**Why the mean is β₀.** In Price = β₀ + ε there is no X, so the model predicts the same value β₀ for every wine; OLS chooses β₀ to minimise Σ(Price − β₀)², and that sum is smallest at the sample mean, so the estimate of β₀ is x̄ = $29.58.

---

## B2 · Continuous Y, one dummy X (Napa valley)

| | Napa (=1) | Other (=0) | Cells |
|---|---|---|---|
| n | 287 | 519 | B21, C21 |
| Mean Price | $43.06 | $22.13 | B22, C22 |
| Variance | 2,689.99 | 171.31 | B23, C23 |

### B2.1 Welch t-test
1. H₀: μ(Napa) = μ(Other)  vs  H₁: μ(Napa) ≠ μ(Other).
2. Welch t = (43.06 − 22.13) / √(2689.99/287 + 171.31/519) = 20.93 / 3.11 = **6.72** (B28), Welch df ≈ 306.3 (B29).
3. p-value = **8.9 × 10⁻¹¹** (B30 `=T.DIST.2T(ABS(B28),B29)`; check B31 `=T.TEST(Data!U2:U807, Data!V2:V807, 2, 3)`) < 0.05.
4. Decision: **reject H₀**. On average Napa wines cost about **$20.93 more per bottle** than other wines. This is a raw difference: it does not yet hold quality, variety or age constant.

### B2.2 Regression on the Napa dummy
LINEST (K42:L46): **Price = 22.13 + 20.93 · Napa** (R² = 0.086). The intercept (22.13) is the mean price of non-Napa wines (C22), and the slope (20.93) equals 43.06 − 22.13, the difference between the two group means (B26). The check cells B56 (= slope − difference) and B57 (= intercept − mean Other) are both 0. Reason: for Napa = 0 the model predicts β₀ and for Napa = 1 it predicts β₀ + β₁; least squares sets each prediction to its group mean.

### B2.3 Which t-test does the regression reproduce?
The slope's t (8.71, D54) is not Welch's t (6.72). The regression reproduces exactly the **pooled (equal-variance) two-sample t-test**: pooled SE = 2.403 (B36) = SE of the slope; pooled t = 8.71 (B37); df = 804 (B38) = residual df; p = 1.7 × 10⁻¹⁷ (B39, `T.TEST` type 2) = the slope's p-value; check B58 = 0. The reason is that OLS assumes one common error variance for all wines and estimates it from all residuals together: SSE/(n − 2) = 858,075.6 / 804 = 1,067.26 = pooled variance s_p² (B35) = 32.67² (Standard Error, B44). Welch's test keeps a separate variance for each group. Since the Napa variance is about 16 times larger, Welch's SE (3.11) is the more reliable one here. Both tests reject H₀.

### B2.4 Safer rank-based check
Price is strongly right-skewed (A5 histogram; skewness 7.87, 'Part A'!B24; mean $29.58 > median $22, 'Part A'!B19). A safer check is the **Wilcoxon rank-sum test (Mann–Whitney U test)**. It replaces prices by their ranks, so a few very expensive bottles (up to $440) cannot dominate the comparison. (Not required to run.)

---

## Q&A practice (individual questions, 25 points)

- **STDEV.S vs STDEV.P?** The 806 wines are a sample; `.S` divides by n − 1.
- **Why 1.963, not 1.96?** σ is unknown, so we use s and the t distribution with n − 1 = 805 df. With n this large, t ≈ z.
- **What does p = 0.00015 mean?** If μ really were 25, a t this extreme would occur in about 0.015% of samples. It is *not* the probability that H₀ is true.
- **Does B2 prove the Napa name raises prices?** No, it is a raw comparison. Napa wines may also score higher, be more often Cabernet, or age longer. B5 controls for these (links to B5.4 and B6.3).
- **Price is not normal. Are the t-tests valid?** Largely yes. With n = 806 the CLT makes sample means approximately normal (links to A6). Mann–Whitney is a robustness check.
- **Why does the regression give the pooled t?** OLS assumes one error variance for every wine (homoskedasticity), so its SE is the pooled SE.
- **What does R² = 0.086 mean?** The Napa dummy alone explains about 8.6% of the variation in Price.
- **Why is Welch's df (306) so much lower than 804?** Almost all of Welch's SE comes from the Napa term (s²/n = 2690/287 ≈ 9.4 vs 171/519 ≈ 0.3), so its df ends up close to Napa's own n − 1 = 286 rather than 804.
