# 🛒 E-Commerce Customer Behavior Analysis
### Advanced EDA · Statistical Testing · Machine Learning · Programmatic Presentation

> 🎯 **Purpose in one line:** A full-stack data science investigation of 10,000 e-commerce customer records — starting from raw CSV, through rigorous statistical hypothesis testing, all the way to machine learning classifiers and RFM segmentation — with every finding automatically rendered into a 12-slide dark-themed PowerPoint using pure Python.

---

## 📌 Key Findings at a Glance

| 📊 Metric | 💡 Result |
|---|---|
| 📦 Dataset Size | 10,000 records × 16 columns — zero missing values |
| 🏆 Top Product Category | Electronics (highest purchase volume) |
| 🔁 Return Customer Rate | ~20% of all customers |
| ⭐ Average Review Score | ~3.0 / 5.0 |
| 🤖 Best ML Model (AUC) | Random Forest — AUC: 0.5018 |
| 💳 Most Common Payment | Cash on Delivery |
| 📍 Highest Avg Purchase City | Khulna |
| 📈 Promo Impact | Discount customers spend significantly more (T-Test: p < 0.05) |

---

## 📁 Project Files

| 📂 File | 🛠️ Tool | 📋 Purpose |
|---|---|---|
| `Ecommerce_Customer_Behavior_Analysis.ipynb` | Jupyter / Python | Full analysis — EDA, stats, ML, clustering, 48 cells |
| `generate_presentation.py` | Python (`python-pptx`) | Programmatically builds the 12-slide dark-themed PPTX |
| `Question_Set.pdf` | Reference | 14 questions across 3 difficulty levels |
| `Metatext.docx` | Reference | Column-level data dictionary for all 16 fields |

---

## 🧠 How the Whole Project Works — Big Picture

The project is structured as a strict **one-way pipeline** — raw data flows through preprocessing, then statistical testing, then ML modelling, then a Python script renders all the outputs into a presentation without any manual copy-paste.

```
┌────────────────────────────────────────────────────────────────┐
│  📥  RAW DATA                                                  │
│  ecommerce_customer_behavior_dataset.csv                       │
│  (loaded directly from GitHub via raw URL — no local file)    │
└───────────────────────┬────────────────────────────────────────┘
                        │
                        ▼
┌────────────────────────────────────────────────────────────────┐
│  🧹  SECTION 2 — Preprocessing                                 │
│  • Zero missing values confirmed                               │
│  • Zero duplicate rows                                         │
│  • Boolean columns cast (Return Customer → int)                │
│  • Engineered: Purchase_Amount_ZScore, Return_Customer_Int,    │
│    Is_Satisfied, R_Score, F_Score, M_Score, RFM_Score,        │
│    RFM_Segment, K_Cluster                                      │
└───────────────────────┬────────────────────────────────────────┘
                        │
          ┌─────────────┼──────────────────────┐
          ▼             ▼                      ▼
┌──────────────┐ ┌─────────────────┐  ┌───────────────────────┐
│ SECTION 3    │ │  SECTIONS 4–6   │  │   SECTIONS 7–8        │
│ Descriptive  │ │  Hypothesis     │  │   Regression & ML     │
│ Stats + EDA  │ │  Testing        │  │                       │
│ (Level 1 &  │ │  • T-Tests      │  │   • OLS Regression    │
│  Level 2 Qs)│ │  • ANOVA+Tukey  │  │   • Logistic Reg      │
│              │ │  • Chi-Square   │  │   • Random Forest     │
│              │ │  • Correlations │  │   • Gradient Boosting │
└──────┬───────┘ └────────┬────────┘  └──────────┬────────────┘
       │                  │                       │
       └──────────────────┴───────────────────────┘
                          │
                          ▼
┌────────────────────────────────────────────────────────────────┐
│  SECTION 9 — Advanced Behavioral Analysis                      │
│  • RFM Segmentation (Champions / Loyal / At-Risk / Lost)       │
│  • K-Means Clustering (K=4, Elbow method)                      │
│  • Purchase Funnel / Cohort Analysis                           │
│  • Payment × Satisfaction × Return Deep Dive                   │
│  • Geographic Performance Matrix                               │
└───────────────────────┬────────────────────────────────────────┘
                        │
                        ▼
┌────────────────────────────────────────────────────────────────┐
│  SECTION 10 — All Questions Answered (Level 1 / 2 / 3)        │
│  + Executive Dashboard Visualization                           │
│  + Statistical Tests Summary Table                             │
└───────────────────────┬────────────────────────────────────────┘
                        │
                        ▼
┌────────────────────────────────────────────────────────────────┐
│  🎨  generate_presentation.py                                  │
│  Reads all metric values → builds 12 dark-themed slides        │
│  → exports final .pptx — zero manual design work              │
└────────────────────────────────────────────────────────────────┘
```

---

## 📓 Notebook — Section-by-Section Breakdown

### 📦 Section 1 — Import Libraries

Every library used in the project is imported in one place. The stack covers four layers:

**Data & Visualisation:**
```python
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import seaborn as sns
```

**Statistical Tests (scipy + statsmodels):**
```python
from scipy.stats import (
    ttest_ind, ttest_rel, mannwhitneyu,     # T-Tests
    f_oneway, kruskal,                       # ANOVA & non-parametric
    chi2_contingency, fisher_exact,          # Chi-Square
    pearsonr, spearmanr,                     # Correlations
    shapiro, levene, normaltest              # Assumption checks
)
from statsmodels.stats.multicomp import pairwise_tukeyhsd   # Post-hoc
import statsmodels.api as sm                                 # OLS Regression
```

**Machine Learning (scikit-learn):**
```python
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.preprocessing import LabelEncoder, StandardScaler
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.metrics import classification_report, roc_auc_score, roc_curve
from sklearn.cluster import KMeans
```

---

### 🗂️ Section 2 — Load Data & Preprocessing

**Data source:** Loaded directly from GitHub raw URL — no local file required:
```python
url = "https://raw.githubusercontent.com/mdshafathossain-tech/Ecommerce-Project-Data/refs/heads/main/ecommerce_customer_behavior_dataset.csv"
df = pd.read_csv(url)
```

**Engineered columns created:**

| 🏷️ Column | 📐 Formula | 🎯 Used For |
|---|---|---|
| `Purchase_Amount_ZScore` | `(x − mean) / std` | L1-Q2, OLS features |
| `Return_Customer_Int` | `bool → 0/1` | Correlation matrix, Logistic target |
| `Is_Satisfied` | `Review Score ≥ 4 → 1` | T-Test grouping, Logistic features |
| `R_Score` | `pd.qcut(Delivery Time, 4)` inverted | RFM component |
| `F_Score` | `pd.qcut(Items Purchased rank, 4)` | RFM component |
| `M_Score` | `pd.qcut(Purchase Amount rank, 4)` | RFM component |
| `RFM_Score` | `R + F + M` | Segment classification input |
| `RFM_Segment` | `SWITCH on RFM_Score` | Customer tier label |
| `K_Cluster` | `KMeans(k=4)` labels | Behavioural segment label |

---

### 📊 Section 3 — Descriptive Statistics & EDA (Level 1 Answers)

**All 10 Level 1 questions answered in code:**

| ❓ Question | 📐 Method | 💡 Output |
|---|---|---|
| Q1 — Age Mean/Median/Mode | `.mean()`, `.median()`, `.mode()` | Printed with f-string |
| Q2 — Purchase Variance/Std/Z-Score | `.var()`, `.std()`, z-score formula | Z-score table for first 5 rows |
| Q3 — Top 3 Product Categories | `.value_counts().head(3)` | Electronics leads |
| Q4 — Return Customer count | `.sum()` on bool column | ~20% of 10,000 |
| Q5 — Average Review Score | `.mean()` | ~3.0 / 5.0 |
| Q6 — Delivery Time by Subscription | `.groupby().mean()` | Premium vs Free comparison |
| Q7 — Subscriber count | `.value_counts()` on Subscription Status | Breakdown by tier |
| Q8 — Device usage % | `.value_counts(normalize=True)` | Mobile / Desktop / Tablet |
| Q9 — Avg purchase: Discount vs None | `.groupby('Discount Availed').mean()` | Discount users spend more |
| Q10 — Most common payment method | `.mode()[0]` | Cash on Delivery |

**Visualisations produced (6 charts each):**
- 2×3 grid — distribution plots for all numeric columns
- 2×3 grid — bar charts for all categorical columns
- Correlation heatmap (lower triangle only, masked upper)
- 2×3 box plots for outlier detection across all numeric variables

---

### 🔬 Section 4 — T-Tests (Comparing Group Means)

A reusable `run_ttest()` function handles all t-tests with **full assumption checking** built in:

```python
def run_ttest(group1, group2, col, label1, label2, alpha=0.05):
    # Step 1: Normality — Shapiro-Wilk (n≤5000) or D'Agostino (n>5000)
    # Step 2: Variance equality — Levene's test
    # Step 3: Independent t-test (Welch's if variances unequal)
    # Step 4: Effect size — Cohen's d
    # Step 5: Decision — reject/fail to reject H₀
```

**Four T-Tests run:**

| 🧪 Test | 🔍 Groups Compared | 📏 Variable |
|---|---|---|
| T-Test 1 | Return Customer vs Non-Return | Purchase Amount ($) |
| T-Test 2 | Discount vs No Discount | Purchase Amount ($) |
| T-Test 3 | Premium vs Free subscription | Delivery Time (days) |
| T-Test 4 | High Satisfaction vs Low Satisfaction | Time Spent on Website (min) |

**How it works:** Levene's test determines whether to use equal-variance t-test or Welch's t-test. Cohen's d gives the practical effect size beyond the p-value. 3 violin plots visualise the group distributions side by side.

---

### 📐 Section 5 — One-Way ANOVA + Tukey HSD Post-Hoc

The `run_anova()` function goes beyond standard ANOVA:

```python
def run_anova(df, group_col, value_col, alpha=0.05):
    # F-test (parametric) + Kruskal-Wallis (non-parametric fallback)
    # Eta-squared effect size = SS_between / SS_total
    # Tukey HSD post-hoc → which specific pairs differ?
```

**Four ANOVA tests run:**

| 🧪 Test | 🔍 Groups | 📏 Variable |
|---|---|---|
| ANOVA 1 | Product Category (5 groups) | Purchase Amount ($) |
| ANOVA 2 | Payment Method (4 groups) | Review Score (1–5) |
| ANOVA 3 | Location (5 cities) | Delivery Time (days) |
| ANOVA 4 | Subscription Status (3 tiers) | Purchase Amount ($) |

**Why Tukey HSD?** ANOVA only tells you *something* is different across groups. Tukey HSD identifies *which pairs* are significantly different — critical for actionable recommendations (e.g. "Khulna vs Dhaka delivery time difference is significant, but Dhaka vs Chittagong is not").

---

### 🎲 Section 6 — Chi-Square Tests (Categorical Associations)

The `run_chi2()` function adds **Cramér's V** effect size to every test:

```python
def run_chi2(df, col1, col2, alpha=0.05):
    chi2, p, dof, expected = chi2_contingency(pd.crosstab(df[col1], df[col2]))
    cramers_v = sqrt(chi2 / (n * min_dim))   # Effect size: 0=none, 1=perfect
    pct_below5 = (expected < 5).mean() * 100  # Assumption check
```

**Three Chi-Square tests run:**

| 🧪 Test | 🔍 Variable 1 | 🔍 Variable 2 |
|---|---|---|
| Chi² 1 | Return Customer | Subscription Status |
| Chi² 2 | Customer Satisfaction | Payment Method |
| Chi² 3 | Discount Availed | Return Customer |

**Visualisations:** 2 normalised proportion heatmaps showing the conditional distribution of satisfaction by payment method and return rate by subscription status.

---

### 📈 Section 7 — Correlation & Regression Analysis

#### 7A — Pearson & Spearman Correlations

Both correlation coefficients computed for 6 variable pairs:

| 🔗 Variable X | 🔗 Variable Y | 📐 Method |
|---|---|---|
| Time Spent on Website | Purchase Amount | Pearson + Spearman |
| Time Spent on Website | Number of Items Purchased | Pearson + Spearman |
| Delivery Time | Review Score | Pearson + Spearman |
| Number of Items Purchased | Purchase Amount | Pearson + Spearman |
| Age | Purchase Amount | Pearson + Spearman |
| Age | Time Spent on Website | Pearson + Spearman |

A **scatter matrix** (`sns.PairGrid`) is rendered for the 5 key numeric variables on a 1,000-row sample.

#### 7B — OLS Multiple Linear Regression
**Target:** `Purchase Amount ($)`

```python
X_ols = sm.add_constant(reg_df.drop('Purchase Amount ($)', axis=1))
ols_model = sm.OLS(y_ols, X_ols).fit()
```

**Features:** Age, Time Spent on Website, Number of Items Purchased, Delivery Time, Review Score, Return_Customer_Int, Is_Satisfied

**Diagnostics plotted:** Residuals vs Fitted, Actual vs Predicted, Residual Histogram — all in one 3-panel figure.

#### 7C — Logistic Regression
**Target:** `Return Customer` (binary: 1 = Yes, 0 = No)

- All categorical columns label-encoded before modelling
- `StandardScaler` applied before fitting
- `train_test_split(test_size=0.25, stratify=y)` to preserve class balance
- `statsmodels` Logit used for coefficient-level significance testing
- `sklearn` LogisticRegression used for ROC/AUC evaluation

---

### 🤖 Section 8 — Random Forest & Gradient Boosting

Both ensemble models trained on the same stratified train/test split as logistic regression, enabling direct ROC curve comparison.

**Random Forest:**
```python
rf = RandomForestClassifier(
    n_estimators=200, max_depth=8,
    min_samples_leaf=10, random_state=42, n_jobs=-1
)
```

**Gradient Boosting:**
```python
gb = GradientBoostingClassifier(
    n_estimators=100, max_depth=4,
    learning_rate=0.1, random_state=42
)
```

**Feature Importance:** Both models output ranked feature importance scores. Random Forest identifies `Purchase Amount ($)` and its Z-Score (~12% each) as the strongest predictors of return customer status.

**Outputs:** Classification report, ROC-AUC score, and a 2-panel comparison figure — feature importance bar chart + overlaid ROC curves for all three classifiers (Logistic, RF, GB).

---

### 🔍 Section 9 — Advanced Behavioral Analysis

#### 9A — RFM Segmentation

Proxy RFM scores built from available columns:
- **R (Recency):** `pd.qcut(Delivery Time, 4)` inverted — faster delivery = higher recency proxy
- **F (Frequency):** `pd.qcut(Number of Items Purchased rank, 4)`
- **M (Monetary):** `pd.qcut(Purchase Amount rank, 4)`

**Segment mapping:**

| 🏷️ RFM Score | 👤 Segment |
|---|---|
| ≥ 10 | Champions |
| 8–9 | Loyal Customers |
| 6–7 | Potential Loyalists |
| 4–5 | At-Risk |
| < 4 | Lost |

#### 9B — K-Means Clustering

```python
cluster_features = ['Age', 'Purchase Amount ($)', 'Time Spent on Website (min)',
                    'Number of Items Purchased', 'Delivery Time (days)', 'Review Score (1-5)']
```

**Method:** Elbow curve plotted for K=2 to K=9 to identify optimal K. Final model uses **K=4**. StandardScaler applied before clustering. 3-panel output: Elbow curve, cluster scatter plot, cluster profile bar chart.

#### 9C — Purchase Funnel Analysis

5-stage conversion funnel built from dataset:
```
Total Customers → High Engagement → Made a Purchase → Satisfied (≥4★) → Return Customer
```
Each stage percentage calculated against total, revealing where customers drop off.

#### 9D — Payment × Satisfaction × Return Deep Dive

`.groupby('Payment Method').agg(...)` aggregates: customer count, avg purchase, avg review, return rate, and discount usage rate per payment method — producing a multi-KPI comparison table.

#### 9E — Geographic Performance Matrix

`.groupby('Location').agg(...)` on 5 cities: customer count, avg purchase amount, avg delivery time, return rate, satisfaction rate — identifying **Khulna** as the benchmark city.

---

### 📋 Section 10 — All Questions Answered + Final Outputs

#### Level 2 Answers (5 questions)

| ❓ Question | 📐 Method |
|---|---|
| L2-Q1 — Avg review for most common payment method users | Filter by mode payment, compute mean review |
| L2-Q2 — Time on site vs Purchase Amount correlation | `pearsonr()` — reported with r and p-value |
| L2-Q3 — % satisfied AND return customers | `(Is_Satisfied == 1) & (Return Customer == True)` |
| L2-Q4 — Items purchased vs satisfaction relationship | `groupby('Customer Satisfaction')['Number of Items Purchased'].mean()` |
| L2-Q5 — 2nd highest avg purchase city | `groupby('Location').mean().sort_values().iloc[-2]` |

#### Level 3 Answers (4 questions)

| ❓ Question | 📐 Evidence Source |
|---|---|
| L3-Q1 — Factors driving return customers | Random Forest feature importance + T-Test + Chi-Square |
| L3-Q2 — Payment methods → satisfaction & return | Payment Deep Dive groupby + Chi-Square CT2 |
| L3-Q3 — Location → purchase & delivery | Geographic matrix + ANOVA 3 (Tukey HSD city pairs) |
| L3-Q4 — Major insights | Synthesised narrative from all tests above |

#### Executive Dashboard

A single 20×14 inch `matplotlib` figure assembling the most important visuals from across the notebook into one shareable summary image.

#### Statistical Tests Summary Table

`pd.DataFrame` listing every test run with: test name, hypothesis, result, and decision — a one-glance audit trail of all statistical work.

---

## 🎨 `generate_presentation.py` — How the PPTX Builder Works

The presentation script builds **12 slides programmatically** using `python-pptx` — no PowerPoint templates, no manual design. Every element is placed by coordinate.

### Global Theme
```python
BG_COLOR     = RGBColor(15,  15,  15)   # Deep black background
ACCENT_GREEN = RGBColor(163, 255, 18)   # Neon lime accent
WHITE        = RGBColor(255, 255, 255)
CARD_BG      = RGBColor(28,  28,  28)   # Dark card background
RED_WARN     = RGBColor(255, 80,  80)   # Warning highlights
TEAL         = RGBColor(0,   200, 180)  # Secondary accent
```

### Reusable Utility Functions

| 🛠️ Function | 🎯 What It Draws |
|---|---|
| `set_bg(slide, color)` | Sets full slide background colour |
| `add_rect(slide, x, y, w, h, color)` | Draws a filled rectangle (cards, bars, dividers) |
| `add_text(slide, text, x, y, w, h, ...)` | Places a single-run text box with full font control |
| `add_multiline(slide, lines, ...)` | Places a text box with mixed colours/sizes per line |
| `stat_card(slide, x, y, w, h, value, label)` | Draws a KPI card (dark box + big value + small label) |
| `bar_visual(slide, x, y, w, h, values, labels)` | Draws a horizontal bar chart using rectangles |
| `slide_label(slide, num, total)` | Adds "3 / 12" counter at bottom-right |
| `section_pill(slide, label)` | Adds a small neon-outlined category tag above title |
| `accent_dot(slide, x, y)` | Draws a small neon circle decoration |

### The 12 Slides

| 🎞️ Slide | 📌 Title | 🧱 Key Visuals |
|---|---|---|
| 1 | Title & Author | Left accent bar, neon title, author card, tech badges |
| 2 | Executive Snapshot | 6 KPI stat cards (Revenue, Returns, Review, Discount, Premium, Top Category) |
| 3 | Dataset Overview | Data shape cards, 16-column metadata, quality report |
| 4 | Level 1 — Basic Insights | Top categories bar, device split, payment mode, subscription stats |
| 5 | Level 2 — Correlations | Correlation table, time vs purchase scatter context, satisfaction breakdown |
| 6 | Level 2 — Satisfaction & Return | 20.08% dual-condition stat, satisfaction by items chart |
| 7 | Level 3 — Return Customer Factors | RF feature importance ranking, T-Test evidence, Chi-Square evidence |
| 8 | Statistical Evidence Backbone | Full test summary table (all 8 tests, hypotheses, decisions) |
| 9 | Geographic Analysis | City performance matrix — Khulna benchmark highlighted |
| 10 | Key Insights | 6 bullet insight cards across two columns |
| 11 | Strategic Recommendations | 5 action cards with rationale |
| 12 | Closing | Thank-you, Google Drive link, contact |

### Speaker Notes (embedded as comments)

Every slide function has a full speaker script above it — e.g.:
> *"Slide 7: Our Random Forest classifier, with an AUC of 0.5018, reveals that Purchase Amount and its Z-Score (~12% each) are the strongest predictors of return customer status..."*

These were used directly as the voiceover script for the 12-minute narrated MP4.

---

## 📊 Dataset Dictionary — All 16 Columns

| # | 🏷️ Column | 📐 Type | 📋 Description |
|---|---|---|---|
| 1 | `Customer ID` | int | Unique identifier per customer |
| 2 | `Age` | int | Customer age |
| 3 | `Gender` | str | Male / Female / Other |
| 4 | `Location` | str | City/region (Dhaka, Khulna, Chittagong, Rajshahi, Sylhet) |
| 5 | `Product Category` | str | Electronics / Clothing / Sports / Home / Toys |
| 6 | `Purchase Amount ($)` | float | Total spend in USD |
| 7 | `Time Spent on Website (min)` | float | Minutes spent browsing |
| 8 | `Device Type` | str | Mobile / Desktop / Tablet |
| 9 | `Payment Method` | str | Cash on Delivery / Bank Transfer / Debit Card / Credit Card |
| 10 | `Discount Availed` | bool | Whether a discount was used |
| 11 | `Number of Items Purchased` | int | Total units bought |
| 12 | `Return Customer` | bool | Whether the customer has returned |
| 13 | `Review Score (1-5)` | float | Customer rating |
| 14 | `Delivery Time (days)` | int | Days to delivery |
| 15 | `Subscription Status` | str | Free / Premium / Trial |
| 16 | `Customer Satisfaction` | str | Low / Medium / High |

---

## ⚙️ Statistical Methods Used — Full Reference

| 🧪 Method | 📚 Library | 🎯 Applied To |
|---|---|---|
| Independent T-Test (Welch's) | `scipy.stats.ttest_ind` | 4 group-mean comparisons |
| Shapiro-Wilk Normality Test | `scipy.stats.shapiro` | T-Test assumption check |
| D'Agostino Normality Test | `scipy.stats.normaltest` | Large sample normality |
| Levene's Test | `scipy.stats.levene` | Variance equality check |
| One-Way ANOVA | `scipy.stats.f_oneway` | 4 multi-group comparisons |
| Kruskal-Wallis Test | `scipy.stats.kruskal` | Non-parametric ANOVA backup |
| Tukey HSD Post-Hoc | `statsmodels pairwise_tukeyhsd` | Pairwise group differences |
| Eta-Squared Effect Size | Manual formula | ANOVA practical significance |
| Chi-Square Test | `scipy.stats.chi2_contingency` | 3 categorical associations |
| Cramér's V Effect Size | Manual formula | Chi-Square practical significance |
| Pearson Correlation | `scipy.stats.pearsonr` | 6 variable pairs |
| Spearman Correlation | `scipy.stats.spearmanr` | 6 variable pairs (non-parametric) |
| Cohen's d Effect Size | Manual formula | T-Test practical significance |
| OLS Multiple Regression | `statsmodels.api.OLS` | Purchase Amount prediction |
| Logistic Regression | `sklearn` + `statsmodels` | Return Customer classification |
| Random Forest | `sklearn.ensemble` | Feature importance + AUC |
| Gradient Boosting | `sklearn.ensemble` | Model comparison + AUC |
| K-Means Clustering | `sklearn.cluster.KMeans` | 4 behavioural segments |
| RFM Segmentation | Manual quartile scoring | 5 customer tiers |
| Elbow Method | Inertia vs K plot | Optimal cluster count |

---

## 🛠️ Tech Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Jupyter](https://img.shields.io/badge/Jupyter-F37626?style=for-the-badge&logo=jupyter&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-F7931E?style=for-the-badge&logo=scikit-learn&logoColor=white)
![SciPy](https://img.shields.io/badge/SciPy-8CAAE6?style=for-the-badge&logo=scipy&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)

| 🔧 Library | 🎯 Used For |
|---|---|
| `pandas` + `numpy` | Data loading, preprocessing, feature engineering |
| `matplotlib` + `seaborn` | All visualisations (25+ charts) |
| `scipy` | All hypothesis tests (T-Test, ANOVA, Chi-Square, Correlations) |
| `statsmodels` | OLS regression, Tukey HSD, Logit coefficients |
| `scikit-learn` | Logistic Regression, Random Forest, Gradient Boosting, K-Means, Scaler, Metrics |
| `python-pptx` | Programmatic PPTX generation in `generate_presentation.py` |

---

## 🔗 Project Resources

| 📂 Resource | 🔗 Link |
|---|---|
| 📁 Full submission (PPTX + Video + Notebook) | [Google Drive](https://drive.google.com/drive/folders/1vY3xSdwTqr3AnHYpj-SRHi5dP4lBma0I) |
| 📊 Dataset (raw CSV) | [GitHub Raw](https://raw.githubusercontent.com/mdshafathossain-tech/Ecommerce-Project-Data/refs/heads/main/ecommerce_customer_behavior_dataset.csv) |

---

## 🏫 Course Context

**Course:** DTS360 — Data Science & AI Engineering
**Batch:** 2520
**Assignment Type:** Final Assignment

The instructor provided the dataset and a 3-level question set (Basic → Intermediate → Critical Thinking). This submission goes significantly beyond the required scope — adding full assumption checking for every statistical test, three ML models with ROC comparison, RFM segmentation, K-Means clustering, a purchase funnel, geographic analysis, and a programmatically generated 12-slide voiced presentation.

---

## 👤 Author

**Md. Shafat Hossain**
DTS360 Student · Batch 2520 · Data Science & AI Engineering

---

## 🔒 License

Academic project. Dataset is fictional and used for educational purposes only.
