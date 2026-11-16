# Supply Planning Model (MRP in Excel, generated with Python)

A forecast-driven planning model for a fictional snack manufacturer, **Maple Ridge Foods**. One Python script writes an Excel workbook built entirely from **live formulas**. Change the demand forecast or the scenario lever, and production, component requirements, purchase orders, shortage alerts and supplier cash needs all recalculate in Excel. There are no macros and nothing is hard-coded.

**Download the workbook:** [`PlanningDemo-MapleRidgeFoods.xlsx`](PlanningDemo-MapleRidgeFoods.xlsx). Open it in Excel, go to `Inputs`, set `C4` to `+25%` and watch the Dashboard.

## Why

Many small manufacturers plan purchasing in disconnected spreadsheets: sales sends a forecast, someone guesses what to order, and shortages show up when a line stops. This model shows the standard way to connect it all (**material requirements planning, MRP**) in a tool every operations team already has.

## How the model flows

```
Inputs (forecast by product x month, scenario %)
   │
   ▼
Production = forecast x (1 + scenario), whole units
   │
   ▼  explode through the bill of materials (SUMPRODUCT)
Requirements = component quantity needed per month
   │
   ▼  projected stock vs. safety stock
Purchasing (MRP)
   • planned receipt when stock would fall below safety stock
   • rounded UP to the supplier's minimum order quantity
   • "place PO" month = receipt month minus supplier lead time
   • flags receipts needed sooner than the lead time allows
   │
   ▼
Cash = supplier payments by month, using each component's payment terms
   │
   ▼
Dashboard: one page for the monthly planning meeting (charts + alerts)
```

## Sheets

| Sheet | What it holds |
|---|---|
| Start here | How to use the model |
| Dashboard | Production, purchasing and cash summary with charts and shortage alerts |
| Inputs | Monthly forecast for 3 products plus one demand-scenario lever (yellow cells) |
| BOM | Quantity of each of 7 components per unit of product |
| Components | Unit cost, lead time, on hand, safety stock, minimum order quantity, payment terms |
| Production | Planned production |
| Requirements | Component needs per month |
| Purchasing | Projected stock, planned receipts, PO placement dates, lead-time alerts |
| Cash | Supplier payments by month |

## Rebuild it

```sh
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python3 build_demo.py                 # writes PlanningDemo-MapleRidgeFoods.xlsx
```

Change `PRODUCTS`, `COMPONENTS` or `BOM` at the top of `build_demo.py` to model a different business.

## Notes

- All data is invented. No client information is included.
- Designed by Sergey Gruzd and built with Claude Code (AI pair programming).
