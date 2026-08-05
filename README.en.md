# Cash Flow Statement Auto-Generator (Claude Code / WorkBuddy Skill)

Automatically compiles a **Cash Flow Statement** from the **Balance Sheet + Income
Statement (Profit & Loss) + a few off-balance-sheet inputs**. The formulas strictly
follow the Chinese template `编制现金流量表2024.xls` and its compilation guide:
**direct method for the main statement, indirect method for the supplementary
schedule**, classifying all cash flows into Operating / Investing / Financing activities.

**Author**: Sichuan Jinmu Zhide Technology Co., Ltd.（四川金沐志德科技有限公司）

Works out of the box with statements of companies reporting under the **new CAS
(2017+)** as well as legacy formats — accounts are auto-mapped (`合同负债→预收款项`,
`应收款项融资→应收票据`, `开发成本→存货`, `拆入资金→短期借款`, `消耗性生物资产→存货`,
`矿区权益→无形资产`, etc.), so **no manual account conversion is needed** for most
industries (real estate, construction, banking, agriculture, energy, manufacturing...).

## Quick Start

### Option A: Single workbook (easiest)

If the Excel workbook contains the three sheets "资产负债表 / 利润及利润分配表 /
表外数据录入", just pass one file:

```bash
pip install openpyxl xlrd          # only needed for Excel parsing
python scripts/cash_flow_generator.py --bs statements.xlsx --out output --fmt all
```

### Option B: Three separate inputs

```bash
python scripts/cash_flow_generator.py \
  --bs balance_sheet.xlsx --pl income_statement.xlsx --extra off_sheet.xlsx \
  --out output --fmt all
```

Supported input formats: `.xlsx` / `.xls` / `.csv` (UTF-8 or GBK) / pasted
Markdown table text. Any of the three arguments can be the table text itself.

### Off-Balance-Sheet Collection Form (recommended)

`examples/表外数据收集表.xlsx` (plus .csv) is the key to accuracy. Fill in
**"支付给职工的工资" (wages, required)** plus whatever true values you have
(taxes, depreciation, amortization, disposal gains/losses...) and pass it as
`--extra`. Empty cells are auto-estimated by the engine.

> Measured accuracy: with complete off-balance-sheet data, inflow items such as
> cash received from sales stay within ±7%. The operating net cash flow is
> balanced against the change in monetary funds; firms with large time deposits /
> wealth-management products (monetary funds ≠ cash equivalents) will show a
> systematic deviation — always cross-check with the official cash flow footnote.

### High-Precision Mode (errors → ~0)

In the second section of the collection form you can enter **true values straight
from your official cash flow statement** (sales received, purchases paid, wages,
taxes, other operating inflows, etc.). Key switch: once you fill in
"收到的其他与经营活动有关的现金" (other operating cash received), the balancing
trick is disabled and the operating net = sum of the true items. Verified on
SANY Heavy Industry 2025: **+0.00%** error when all true values were provided.

```
Accuracy tiers: default ±7%~20% → with off-sheet data ±5% → high-precision ±0%
```

### Prior-Year Amounts

```bash
python scripts/cash_flow_generator.py --bs bs.xlsx --pl pl.xlsx --extra extra.xlsx \
  --bs-prior last_balance_sheet.xlsx --pl-prior last_income_statement.xlsx
```

## Outputs

| File | Contents |
| --- | --- |
| `现金流量表.md` | Main statement (line no. / current / prior year) + supplementary schedule + cash reconciliation + negative-adjustment notes + validations + compilation notes |
| `现金流量表.xlsx` | Multi-sheet workbook: main statement, supplementary schedule, notes & validations |
| `现金流量表.json` | Structured result for programmatic use |

## Core Formulas (summary — full mapping in `references/公式口径与编制说明.md`)

- **Cash received from sales** = Revenue×(1+output VAT) + (Notes receivable, begin−end) + (AR, begin−end) + (Advances from customers, end−begin) − discount interest
- **Cash paid for purchases** = (COGS + Inventory, end−begin)×(1+input VAT) + (Notes payable, begin−end) + (AP, begin−end) + (Prepayments, end−begin)
- **Taxes paid** = VAT payable + other taxes + income tax + taxes in admin expenses + taxes in other business costs − (Taxes payable & other payables, end−begin)
- **Other operating cash received** = balancing item (auto-computed to balance the statement)
- **Supplementary (indirect method)**: Net profit → +impairment +depreciation +amortization ±prepaid/accrued ±disposal losses ±finance costs ±investment losses ±deferred tax ±inventory / operating receivables / operating payables changes → operating net (with "other" as the balancing item)
- **Negative adjustment** (guide rule #4): the main statement must not show negative inflows; a negative inflow is moved to the corresponding outflow item.

## Overriding Parameters (set VAT rates / ratios to your company's actuals)

```bash
python scripts/cash_flow_generator.py --bs statements.xlsx \
  --json-params '{"sale_vat_rate":0.09, "extra_sale_tax_rate":0.06, "four_gold_rate":0.2}'
```

| Param | Default | Meaning |
| --- | --- | --- |
| `sale_vat_rate` | 0.13 | Output VAT for sales conversion |
| `purchase_vat_rate` | 0.17 | Input VAT for purchases conversion |
| `extra_sale_tax_rate` | 0.06 | Output VAT for off-sheet estimates |
| `extra_pur_tax_rate` | 0.06 | Input VAT for off-sheet estimates |
| `four_gold_rate` | 0.266 | Social insurance & housing fund ratio |
| `welfare_rate` | 0.0106 | Other welfare expense ratio |
| `bad_debt_rate` | 0.005 | Bad-debt provision ratio (reference only) |

## Repository Layout

```
cashflow-statement-skill/
├── SKILL.md                          # Skill manifest (triggers & usage)
├── README.md                         # Chinese readme
├── README.en.md                      # This file
├── LICENSE
├── .gitignore
├── scripts/
│   ├── cash_flow_generator.py        # Core engine (standalone CLI)
│   ├── gen_extra_form.py             # Regenerates the collection form
│   └── requirements.txt
├── references/
│   ├── 公式口径与编制说明.md          # Line-by-line formula reference
│   ├── 验证报告-三一重工2025.md       # Validation report vs. public data
│   └── 现金流量表编制说明.txt          # Original compilation guide
└── examples/
    ├── 示例_资产负债表.csv            # Sample data extracted from the template
    ├── 示例_利润表.csv
    ├── 示例_表外数据.csv
    ├── 表外数据收集表.xlsx / .csv     # Off-balance-sheet collection form
    ├── 对比报告-默认vs高精度.md       # Default vs. high-precision comparison
    └── output/                       # Sample outputs (md / xlsx / json)
```

## Important Notes

- The statement reflects the **approximate** cash position for external reporting;
  for corporate decisions compile from detailed ledgers (per the original guide).
- VAT rates and provision ratios **must be adjusted** to the company's actuals.
- When off-balance-sheet data is missing, the engine estimates using the template
  ratios (wages → 26.6% four-funds, 1.06% welfare; VAT from revenue/cost) and
  marks these in the notes.

## Validation

Verified against the template's cached values (zero error on 13 key items) and
against SANY Heavy Industry's 2025 public statements (sales received +6.8% with
balance-sheet method only; 0.00% in high-precision mode). See
`references/验证报告-三一重工2025.md` for the full breakdown.
