# Invoice Chaser — SaaS Financial Model

A 36-month, cohort-driven financial model for an early-stage SaaS product,
built programmatically in Python and delivered as an interactive Excel workbook.
Scenario-switchable, with proper deferred-revenue mechanics and a self-checking
integrity row.

> ⚠️ **The numbers are illustrative, not observed.** Every operating assumption
> in this model (churn, cost-per-click, conversion, etc.) is a placeholder drawn
> from typical low-ticket B2B SaaS ranges. **None of it is measured data for a
> real product.** This is a decision-framework and modelling-craft exercise, not
> a forecast. Read the outputs as "here is what these assumptions imply and what
> I'd need to measure to trust them," never as evidence.

---

## What this is

The product being modelled is a hypothetical automated invoice-chasing tool for
French *auto-entrepreneurs* (freelancers). Before committing time to building the
product, the question was: **do the unit economics even work, and if not, what
would I need to find out to know?**

The model is designed to answer that honestly — including admitting when it
*can't* answer yet.

## What's interesting about it (the transferable bits)

These hold regardless of the placeholder inputs:

- **Billings vs. recognised revenue are separated.** Annual customers pay ~10×
  the monthly price upfront but revenue is recognised straight-line over 12
  months. The gap sits in a **deferred-revenue rollforward**, so the P&L and the
  cash flow never get conflated.
- **Annual billing is modelled as financing, not pricing.** On illustrative
  assumptions, shifting customers onto annual plans changed the *funding
  requirement* far more than it changed the *P&L* — a lever that's invisible if
  you only build an income statement.
- **Two separate cohort waterfalls.** Monthly plans churn every month; annual
  plans are locked for 12 and face a single renewal decision at each anniversary.
  Blending them into one churn rate would hide the thing that matters most.
- **The model checks itself.** An integrity row forces
  `net cash − EBITDA − Δdeferred = 0` in every one of the 36 months. If the logic
  ever breaks, the sheet reports it instead of quietly returning a clean-looking
  wrong answer.
- **Sensitivity over scenarios.** The headline finding is that the verdict swings
  from insolvent to excellent on **three unmeasured drivers** — churn,
  cost-per-click, and trial-to-paid conversion — so the real output is a
  *validation to-do list*, not a number.

## Workbook structure

| Sheet | What it holds |
|---|---|
| **Read Me** | How to use it, colour legend, and an explicit health-warning on the assumptions |
| **Assumptions** | Every driver, three scenarios side by side, plus a derived block. The only place raw inputs live |
| **Cohorts** | Acquisition funnel + two 36×36 cohort waterfalls (monthly and annual) |
| **P&L** | Billings vs. recognised revenue, cost of revenue, gross margin, opex, EBITDA |
| **Cash Flow** | Cash walk, deferred-revenue rollforward, and the monthly integrity check |
| **Unit Economics** | LTV, CAC (paid and blended), payback, LTV:CAC, and two sensitivity grids |
| **Dashboard** | Headline outputs, breakeven months, funding requirement, NPV, and charts |

## Headline outputs (Base case, illustrative assumptions)

| Metric | Value |
|---|---|
| LTV : CAC (paid) | 2.13× |
| Range across scenarios | 0.41× → 9.72× |
| Additional funding required | ~€700 |
| First cash-flow-positive month | 7 |
| First EBITDA-positive month | 10 |
| Formulas / errors | 3,523 / 0 |

The point is not the 2.13×. It's that the number is *this sensitive* to three
things nobody has measured yet.

## How to use it

**Just exploring the numbers:** open the workbook and change **`Assumptions!B5`**
to `1` (Conservative), `2` (Base) or `3` (Aggressive). Everything recalculates.
To change a driver, edit the scenario columns on the Assumptions sheet — nothing
else is a manual input.

**Rebuilding from source:**

```bash
pip install -r requirements.txt
python build_model.py
```

This regenerates `Invoice_Chaser_SaaS_Model.xlsx` from scratch. The entire model
— all seven sheets and 3,500+ formulas — is defined in `build_model.py` using
[openpyxl](https://openpyxl.readthedocs.io/); no formula is hand-entered.

## What it deliberately does *not* model

VAT/TVA, corporate tax, founder opportunity cost beyond the salary line,
competitive response, terminal value, and TAM. A 36-month NPV also structurally
understates a business still compounding at month 36. Add these before using any
of it in a real decision.

## Files

```
├── README.md                       # you are here
├── build_model.py                  # generates the whole workbook (openpyxl)
├── Invoice_Chaser_SaaS_Model.xlsx  # the model itself
├── requirements.txt
└── LICENSE
```

## Tech

Python · openpyxl · Excel. No external data sources or dependencies beyond
openpyxl.

---

*Built as a modelling-craft and finance portfolio piece. The engineering is real;
the input numbers are placeholders pending validation.*
