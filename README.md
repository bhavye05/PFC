# MS-Excel

# Probability for Computing — Excel Cheat Sheet
**DSC-06 | Semester 2 | Ramanujan College, University of Delhi**

---

## Practical 1 — Binomial Distribution

**Chart Type:** Clustered Column Chart

**Formula:**
```excel
=BINOM.DIST(k, n, p, FALSE)
```

| Argument | Meaning |
|----------|---------|
| `k` | Number of successes |
| `n` | Total trials |
| `p` | Probability of success per trial |
| `FALSE` | PMF (point probability); use `TRUE` for CDF |

**Key Stats:** Mean = `n×p` | Variance = `n×p×(1−p)`

**Setup:** Enter k (0 to n) in col A → probabilities in col B/C/D for different p values → Select range → Insert → Column → Clustered Column

---

## Practical 2 — Multinomial Distribution

**Chart Type:** Clustered Column Chart

**Formulas:**
```excel
=MULTINOMIAL(X1, X2, X3)
=MULTINOMIAL(X1,X2,X3) * PRODUCT(p1^X1, p2^X2, p3^X3)
```

| Argument | Meaning |
|----------|---------|
| `X1, X2, X3` | Count of each outcome category |
| `p1, p2, p3` | Probability of each category (must sum to 1) |

> First formula gives the **multinomial coefficient** n!/(x1!·x2!·…·xk!).  
> Multiply by the probability term to get the full joint probability.

**Key Stats:** Cov(Xi, Xj) = `−n × pi × pj` (for i ≠ j)

---

## Practical 3 — Poisson Distribution

**Chart Type:** Line with Markers

**Formula:**
```excel
=POISSON.DIST(k, lambda, FALSE)
```

| Argument | Meaning |
|----------|---------|
| `k` | Number of events |
| `lambda` | Mean rate (λ = np) |
| `FALSE` | PMF; use `TRUE` for CDF |

**Key Stats:** Mean = λ | Variance = λ | SD = √λ

**Setup:** Enter k (0–10+) in col A → probabilities in col B → Select → Insert → Line → Line with Markers

---

## Practical 4 — Geometric Distribution

**Chart Type:** Line with Markers

**Formula (manual — no built-in Excel function):**
```excel
= (1 - p)^k * p
```

| Argument | Meaning |
|----------|---------|
| `k` | Number of failures before 1st success |
| `p` | Probability of success on each trial |

**Key Stats:** Mean = `(1−p)/p` | Variance = `(1−p)/p²`

> Excel has no `GEOM.DIST` function — you must type the formula manually for each k value.

---

## Practical 5 — Uniform Distribution

**Chart Type:** Clustered Column Chart

**Formula:**
```excel
= (x2 - x1) / (b - a)
```

| Argument | Meaning |
|----------|---------|
| `a` | Minimum of distribution |
| `b` | Maximum of distribution |
| `x1` | Lower bound of sub-interval |
| `x2` | Upper bound of sub-interval |

**Key Stats:** Mean = `(a+b)/2` | Variance = `(b−a)²/12`

> Probability is directly proportional to sub-interval width — constant PDF height = `1/(b−a)`.

---

## Practical 6 — Exponential Distribution

**Chart Type:** Scatter with Smooth Lines

**Formula:**
```excel
=EXPON.DIST(x, lambda, FALSE)
```

| Argument | Meaning |
|----------|---------|
| `x` | Value (time/distance) |
| `lambda` | Rate parameter λ |
| `FALSE` | PDF; use `TRUE` for CDF |

**Key Stats:** Mean = `1/λ` | Variance = `1/λ²`

---

## Practical 7 — Normal Distribution

**Chart Type:** Scatter with Smooth Lines

**Formulas:**
```excel
=AVERAGE(cell_range)         → Mean
=STDEV(cell_range)           → Standard Deviation
=NORM.DIST(x, mean, sd, FALSE)  → PDF value at x
```

| Argument | Meaning |
|----------|---------|
| `x` | Value to evaluate |
| `mean` | Mean (μ) |
| `sd` | Standard deviation (σ) |
| `FALSE` | PDF; use `TRUE` for CDF |

> PDF peaks at x = μ and is symmetric on both sides (bell curve).

---

## Practical 8 — Cumulative Distribution Functions (CDF)

**Chart Type:** Line with Markers

**Formulas:**
```excel
=NORM.DIST(x, mean, sd, TRUE)       → Normal CDF
=EXPON.DIST(x, lambda, TRUE)        → Exponential CDF
```

> The key difference from Practicals 6 & 7: use `TRUE` (cumulative = TRUE) instead of `FALSE`.

| Distribution | Parameters used |
|---|---|
| Normal CDF | μ = 6000, σ = 2800 |
| Exponential CDF | λ = 0.0025 |

**Setup:** Two separate Line with Markers charts (or one chart with a secondary Y-axis). Both curves are S-shaped (sigmoid), rising steeply near the mean and flattening toward 1.0.

---

## Practical 9 — Distance Between Two Distributions

**Chart Type:** Clustered Column Chart

**Formula (Euclidean Distance):**
```excel
=SQRT(SUMXMY2(array_X, array_Y))
```

| Argument | Meaning |
|----------|---------|
| `array_X` | Probability values from Distribution 1 |
| `array_Y` | Probability values from Distribution 2 |

**Manual calculation:**

| Step | Formula |
|------|---------|
| Binomial probabilities | `=BINOM.DIST(k, n, p, FALSE)` |
| Poisson probabilities | `=POISSON.DIST(k, lambda, FALSE)` |
| Squared differences | `=(B2-C2)^2` for each row |
| Euclidean distance | `=SQRT(SUM(differences_column))` |

> Example: Binomial(n=8, p=0.6) vs Poisson(λ=4.8) → Euclidean Distance ≈ 0.1501

---

## Practical 10 — Application: Binomial Distribution

**Chart Type:** Clustered Column Chart

**Formula:**
```excel
=BINOM.DIST(k, n, p, FALSE)
```

**Case Study — Smartphone Screen Defect Inspection:**

| Parameter | Value |
|-----------|-------|
| n | 6 |
| p | 0.08 |
| Observed k | 2 |
| Formula | `=BINOM.DIST(2, 6, 0.08, FALSE)` |
| Result | ≈ 0.0688 (6.88%) |

**Setup:** Highlight the bar at the observed k value in a different colour (e.g., red) to distinguish it from the rest.

---

## Practical 11 — Application: Poisson Distribution

**Chart Type:** Line with Markers

**Formula:**
```excel
=POISSON.DIST(x, lambda, FALSE)
```

**Case Study — Hospital Emergency Room Trauma Arrivals:**

| Parameter | Value |
|-----------|-------|
| λ | 3.5 |
| Target x | 4 |
| Formula | `=POISSON.DIST(4, 3.5, FALSE)` |
| Result | ≈ 0.1888 (18.88%) |

**Setup:** Highlight the data point at x = 4 with a contrasting marker colour (red).

---

## Practical 12 — Application: Normal Distribution

**Chart Type:** Scatter with Smooth Lines

**Formula:**
```excel
=NORM.DIST(x, mean, sd, FALSE)
```

**Case Study — Standardised Test Score Analysis:**

| Parameter | Value |
|-----------|-------|
| μ | 520 |
| σ | 85 |
| Formula | `=NORM.DIST(x, 520, 85, FALSE)` |

**Setup:** Add a vertical reference line or annotation at x = μ = 520 to mark the peak. Peak density = 0.004694 at x = 520.

---

## Practical 13 — Covariance & Scatter Plot

**Chart Type:** Scatter (XY) with Straight Lines and Markers + Trendline

**Formulas:**
```excel
=COVARIANCE.P(array1, array2)     → Population covariance
=COVARIANCE.S(array1, array2)     → Sample covariance
=PEARSON(array1, array2)          → Pearson r
=CORREL(array1, array2)           → Same as PEARSON
```

**To add a trendline in Excel:** Right-click data series → Add Trendline → Linear → tick "Display R² value"

**Key formulas:**

| Measure | Formula |
|---------|---------|
| Covariance | `Σ[(Xi − X̄)(Yi − Ȳ)] / n` |
| Correlation r | `Σ[(xi−x̄)(yi−ȳ)] / √[Σ(xi−x̄)² × Σ(yi−ȳ)²]` |

---

## Practical 14 — Karl Pearson's Correlation Coefficient

**Chart Type:** Scatter (XY) — three separate charts or one combined chart

**Formula:**
```excel
=PEARSON(array1, array2)
```

| Correlation Type | r Value Range | Interpretation |
|-----------------|---------------|----------------|
| Strong Positive | +0.9 to +1.0 | Both variables increase together |
| Strong Negative | −0.9 to −1.0 | One increases as other decreases |
| Near Zero | −0.1 to +0.1 | No linear relationship |

**Example Results:**

| Pair | r |
|------|---|
| X vs Y1 | −0.9988 (strong negative) |
| X vs Y2 | +0.9972 (strong positive) |
| X vs Y3 | +0.1068 (near zero) |

**Setup:** Create three scatter charts, one per Y variable. Add a trendline to each and display the r value in the legend.

---

## Practical 15 — Bivariate Frequency Distribution

**Chart Type:** Stacked Column Chart

**Formulas:**
```excel
=PEARSON(array1, array2)    → Correlation along each axis
```

**Setup:**
1. Enter the contingency table (rows = Marks groups, columns = Age groups)
2. Select frequency data → Insert → Column → Stacked Column
3. Axis labels: X = Age Group, Y = Frequency
4. Use a colour gradient (light to dark) across Marks ranges

**Interpretation:** Stacked columns show joint frequency distribution. Positive r indicates higher age groups are associated with higher mark ranges.

---

## Practical 16 — Random Numbers from Discrete Distributions

**Chart Type:** Clustered Column Chart (frequency histogram)

**Formulas:**
```excel
=BINOM.INV(1, P, RAND())          → Bernoulli (0 or 1)
=BINOM.INV(n, P, RAND())          → Binomial (count of successes)
```

> There is no built-in `POISSON.INV` — use `=NORM.INV(RAND(), λ, SQRT(λ))` as an approximation for large λ.

**To build a frequency histogram:**
```excel
=COUNTIF(data_range, k)    → Count how many times k appeared
```

**Setup:**
1. Generate 50+ values using `=BINOM.INV(n, p, RAND())` in a column
2. Use `COUNTIF` to count each k value
3. Plot count vs k → Insert → Column → Clustered Column

> Each recalculation produces new random values (press F9 to recalculate).

---

## Practical 17 — Random Numbers from Continuous Distributions

**Chart Type:** Histogram (via Data Analysis ToolPak)

**Formulas:**
```excel
=NORMINV(RAND(), mean, sd)        → Normal distribution
=a + RAND() * (b - a)             → Uniform distribution U(a, b)
```

| Function | Meaning |
|----------|---------|
| `RAND()` | Generates uniform random number ∈ [0, 1] |
| `NORMINV` | Maps the random probability to a Normal value |

**To build a histogram using Data Analysis ToolPak:**
1. Enable: File → Options → Add-ins → Analysis ToolPak → Go → tick → OK
2. Generate 100 values in col A using `=NORMINV(RAND(), mean, sd)`
3. Define bin edges in col B (e.g., 7, 13, 19, 25, 31, 37, 43)
4. Data → Data Analysis → Histogram → set Input Range & Bin Range → OK

---

## Practical 18 — Entropy from a Dataset

**Chart Type:** Clustered Column Chart

**Formulas:**
```excel
=LOG(p, 2)                        → log₂(p) — entropy in bits
=LN(p)                            → logₑ(p) — entropy in nats
=LOG10(p)                         → log₁₀(p) — entropy in hartleys
=-p * LOG(p, 2)                   → Entropy contribution H(x) for one outcome
=SUMPRODUCT(-prob_range, LOG(prob_range, 2))   → Total entropy H(X)
```

**Manual calculation table structure:**

| Column | Content | Formula |
|--------|---------|---------|
| A | Outcome label | — |
| B | P(x) | Enter manually (must sum to 1) |
| C | log₂ P(x) | `=LOG(B2, 2)` |
| D | H(x) = −p·log₂p | `=-B2*LOG(B2, 2)` |
| — | **Total H(X)** | `=SUM(D column)` |

**Entropy reference:**

| Scenario | H(X) |
|----------|------|
| Uniform over 2 outcomes | 1.0 bit |
| Uniform over 4 outcomes | 2.0 bits (maximum) |
| One certain outcome | 0.0 bits |

> Higher entropy = more uncertainty. Unequal probabilities always reduce entropy below the uniform maximum.

---

## Quick Reference — Excel Functions Summary

| Practical | Key Excel Formula |
|-----------|------------------|
| 1 — Binomial | `=BINOM.DIST(k, n, p, FALSE)` |
| 2 — Multinomial | `=MULTINOMIAL(X1, X2, X3)` |
| 3 — Poisson | `=POISSON.DIST(k, λ, FALSE)` |
| 4 — Geometric | `=(1-p)^k * p` (manual) |
| 5 — Uniform | `=(x2-x1)/(b-a)` (manual) |
| 6 — Exponential PDF | `=EXPON.DIST(x, λ, FALSE)` |
| 7 — Normal PDF | `=NORM.DIST(x, μ, σ, FALSE)` |
| 8 — CDF (Normal) | `=NORM.DIST(x, μ, σ, TRUE)` |
| 8 — CDF (Exponential) | `=EXPON.DIST(x, λ, TRUE)` |
| 9 — Euclidean Distance | `=SQRT(SUMXMY2(array1, array2))` |
| 10–12 — Applications | Same as Practicals 1, 3, 7 |
| 13 — Covariance | `=COVARIANCE.P(array1, array2)` |
| 13–14 — Correlation | `=PEARSON(array1, array2)` |
| 15 — Bivariate Freq. | `=PEARSON(array1, array2)` |
| 16 — Random Discrete | `=BINOM.INV(n, p, RAND())` |
| 17 — Random Continuous | `=NORMINV(RAND(), μ, σ)` |
| 18 — Entropy | `=SUMPRODUCT(-p_range, LOG(p_range, 2))` |

---

## Quick Reference — Chart Types

| Chart Type | Used In |
|-----------|---------|
| Clustered Column | Practicals 1, 2, 5, 9, 10, 16, 18 |
| Line with Markers | Practicals 3, 4, 8, 11 |
| Scatter with Smooth Lines | Practicals 6, 7, 12 |
| Scatter (XY) + Trendline | Practicals 13, 14 |
| Stacked Column | Practical 15 |
| Histogram (ToolPak) | Practical 17 |

---

*Probability for Computing · DSC-06 · Session 2025–26*
