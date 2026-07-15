# Redis PS Dashboard — Business Logic & Calculation Reference

This document captures the authoritative business logic used in the PS Dashboard (index.html + Code.gs).
Use this as the source of truth when building new skills, automations, or derivative dashboards.

---

## 1. Fiscal Year & Quarter

Redis uses a **Feb–Jan fiscal year**. The fiscal year label is `calendar year + 1`.

| FQ | Calendar Months | Example |
|----|----------------|---------|
| Q1 | February – April | Q1-2026 = Feb–Apr 2025 |
| Q2 | May – July | Q2-2026 = May–Jul 2025 |
| Q3 | August – October | Q3-2026 = Aug–Oct 2025 |
| Q4 | November – January | Q4-2026 = Nov 2025–Jan 2026 |

**Label rule:**
- January → stays in the same calendar year's Q4 (e.g. Jan 2026 = Q4-2026)
- February–December → year + 1

**Quarter label format:** `Q{N}-{FY}` — e.g. `Q2-2027`

**Quarter ranges (start/end dates):**
- Q1: Feb 1 – Apr 30
- Q2: May 1 – Jul 31
- Q3: Aug 1 – Oct 31
- Q4: Nov 1 – Jan 31 (crosses calendar year)

---

## 2. Reporting Period Boundaries

**Current quarter:** Period end is capped at the **last completed Saturday** (end of the last full Mon–Sat week before today).

```
priorWeekSaturday(today) = this week's Saturday − 7 days
```

**Past quarters:** Period end is the full quarter end date.

**Month filter override:** When a specific month is selected, it overrides the quarter period:
- Start = 1st of the month
- End = last calendar day of the month, OR `priorWeekSaturday(today)` if the month is still in progress
- If capped end < start, end = start

**Working hours:** `countWeekdays(periodStart, periodEnd) × 8`
- Weekdays = Mon–Fri only; Saturday and Sunday excluded

---

## 3. Ramp Logic

A resource is **Ramping** if: `today < startDate + 90 days`

A resource is **Ramped** if: `today ≥ startDate + 90 days`

**Ramp status is evaluated at `periodEnd`**, not today, so historical views reflect the person's status during that period.

**Effective start date for utilization:** `max(periodStart, startDate + 90 days)`
- A person's working hours only count from when they finished ramping.
- Available working hours = `countWeekdays(effectiveStart, periodEnd) × 8`

**Ramp exclusion in email automation:** Resources with `today < startDate + 90 days` are skipped entirely — no missing timecard email is sent.

---

## 4. Utilization Calculations (Resource-Level)

**Billable hours classification:**
A time entry is billable if ANY of the following:
- Has a Project (non-blank) AND the Project Type ≠ "internal"
- OR the Administration Type is one of:
  - `Program Management`
  - `General Product Issues`
  - `Customer Travel`

**PTO hours:** Administration Type = `PTO / Holiday / Sick`

**Three utilization metrics per resource:**

| Metric | Formula |
|--------|---------|
| **Billable Utilization** | `billableHrs / effectiveWorkHrs × 100` |
| **PTO-Adjusted Utilization** | `billableHrs / max(effectiveWorkHrs − ptoHrs, 0) × 100` |
| **Total Utilization** | `totalHrs / effectiveWorkHrs × 100` |

Where `effectiveWorkHrs = countWeekdays(max(periodStart, startDate+90), periodEnd) × 8`

---

## 5. Utilization KPI Averages (Team-Level)

KPI averages are **weighted** (not a simple mean of percentages):

```
avgBillUtil  = sum(billableHrs across all resources) / sum(effectiveWorkHrs across all resources) × 100
avgPTOUtil   = sum(billableHrs) / sum(max(effectiveWorkHrs − ptoHrs, 0)) × 100
```

Only resources with `effectiveWorkHrs > 0` are included in the totals.

---

## 6. Utilization Color Thresholds (Resource-Level)

Applied to all three utilization % columns in the table:

| Range | Color |
|-------|-------|
| ≥ 70% | Green |
| 55–69% | Yellow |
| < 55% | Red |

KPI header tiles (team average) use different thresholds:

| Range | Color |
|-------|-------|
| ≥ 80% | Green |
| 60–79% | Yellow |
| < 60% | Red |

---

## 7. Role Classification

Derived from the `Title` column in PS User Records:

- **Consultant** — title contains any of: `consulting`, `training`, `delivery`
- **Engagement Manager** — everything else

---

## 8. Annual Plan / Resident Engineer — Utilization Ratio

The Annual Plan tab shows open Annual Plan and Resident Engineer projects. These are identified from Active Projects by `PS SKUs` containing any of: `annual`, `resident`, `bundle`, `dedicated`.

Three metrics displayed per project:

| Column | Meaning |
|--------|---------|
| **% Through Contract** | How far through the contract period by date |
| **% Hours Consumed** | `Hrs Consumed / Hrs Entitlement × 100` |
| **Utilization Ratio** | `% Hours Consumed / % Through Contract` |

Utilization Ratio > 1.0 means hours are being consumed faster than the contract timeline pace. These are sourced directly from the sheet (not computed in the dashboard).

**Slide filter thresholds:**
- Slide 6a: Utilization Ratio < 50% (sorted lowest → highest)
- Slide 6b: Utilization Ratio 50–75% (sorted lowest → highest)
- Both slides exclude projects at or above 75% utilization

---

## 9. Deal Attach Rate

**Definition:** % of new customer accounts that also purchased Professional Services in the same quarter.

**Scope:** New Customer deal type only (strips Salesforce picklist prefix like "1. New Customer").

**Formula per geo:**
```
Deal Attach Rate = (# unique accounts with PS purchase) / (# total unique accounts in geo) × 100
```

- De-duplicated by `Opportunity ID` (or `Account + FiscalPeriod` as fallback)
- ARR is counted once per opportunity to avoid double-counting line items

**Geo mapping:**
- **AMER:** "north america", "latam", "latin america", "canada", "america"
- **EMEA:** "emea", "europe", "middle east", "africa"
- **APAC:** "apac", "apj", "asia", "pacific", "australia", "india", "japan"

**PS Product detection:** `Product Family` = "services - professional services", contains "professional services", or = "ps"

**PS Bookings total:** Sum of `Total Price (converted)` across all PS line items in the current quarter (all deal types, not just new customer).

---

## 10. Fiscal Period Key Format

Attach rate data uses fiscal period strings from Salesforce. These are normalized to a `{FY}_Q{N}` key:

Recognized formats (case-insensitive):
- `Q2-FY27`, `Q2-FY2027`
- `FY27-Q2`, `FY2027-Q2`
- `Q2-2027`, `Q2-27`

Two-digit years are interpreted as 20xx. Output key: `2027_Q2`

---

## 11. Missing Timecard Email Automation

**Trigger:** Runs every Monday at 9:00 AM (time-based GAS trigger).

**Week definition:** Previous Monday–Sunday
- `prevMon = today − 7 days`
- `prevSun = today − 1 day`

**Threshold:** A resource is emailed if logged hours for the week < 40.

**Skipped (no email sent):**
- No email address on record
- Resource is still within their 90-day ramp window (`today < startDate + 90 days`)
- Logged hours ≥ 40

**Email behavior:**
- Sent to the resource's email
- CC'd to their manager's email
- Subject: `[Action Required] Missing Timecard – Week of {date}`
- Content: shows hours logged, hours missing, manager name
- Uses Gmail dark-background HTML table layout

**Source sheets:**
- `Hours Data` — time entries (columns: `Work: Owner Name`, `Date`, `Hours (Number)`)
- `PS User records` — roster (columns: `Full Name`, `Email`, `Manager Email`, `Manager Name`, `User Start Day`)

---

## 12. KPI Push to External Sheet

Dashboard pushes live KPIs to spreadsheet `1HNr9QtJ_De_ZLQcHx6-rDBgHze3P7InQfFVQGax8Vjg`, sheet `PS`.

| Cell | Value | When |
|------|-------|------|
| A4 | Avg Billable Utilization % (as decimal, e.g. 0.72) | On Utilization tab visit, current quarter only |
| C4 | Avg PTO-Adjusted Utilization % (as decimal) | On Utilization tab visit, current quarter only |
| G3 | Current quarter Bookings target (total AMER+EMEA+APAC) | On every dashboard open |
| G4 | Current quarter PS Bookings total ($) | On every dashboard open |
| J3 | AMER Deal Attach Rate (as decimal) | On every dashboard open |
| J4 | EMEA Deal Attach Rate (as decimal) | On every dashboard open |
| J5 | APAC Deal Attach Rate (as decimal) | On every dashboard open |
| J6 | Grand (global) Deal Attach Rate (as decimal) | On every dashboard open |

Utilization KPIs (A4, C4) only fire when the Utilization tab is open **and** the selected quarter matches the current quarter.

---

## 13. Data Source Sheet Identification

Sheets are identified by column signature (not sheet name), except where name-based matching takes priority:

| Sheet | Identification Rule |
|-------|---------------------|
| `bookings` | Sheet name = "Bookings", "PS Bookings", or "PS Deals Bookings" |
| `attachRates` | Sheet name = "Attach Rates", "Attach Rate", "PS Attach Rates" (name wins unconditionally) |
| `goLive` | Has `Actual Go Live Date` column, ≤ 20 columns |
| `annualPlan` | Has `Utilization Ratio` but NOT `Health`, ≤ 35 columns |
| `cloudUsage` | Has a column containing `Cloud Usage Rate` |
| `activeProjects` | Has `Health` column AND > 20 columns |
| `hoursData` | Has `Work: Owner Name` or `Hours (Number)` column |
| `psUserRecords` | Has `Title` AND `Full Name` columns |

Annual Plan can also be derived virtually from Active Projects by filtering `PS SKUs` for keywords: `annual`, `resident`, `bundle`, `dedicated`.
