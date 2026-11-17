# supply planning model

an mrp (material requirements planning) model in excel for a made up snack company, maple ridge foods. a python script writes the workbook and everything in it is live formulas, no macros, nothing hardcoded.

change the forecast or the scenario % and production, component needs, purchase orders, shortage alerts and supplier payments all update.

[download the workbook](PlanningDemo-MapleRidgeFoods.xlsx), go to `Inputs`, set `C4` to +25% and look at the dashboard.

## how it flows

```
forecast x scenario      -> production
production x bill of materials  -> component requirements
requirements vs stock + safety stock -> purchase orders
  (rounded up to min order qty, placed lead-time months early,
   flagged if there isn't enough lead time)
purchase orders x payment terms -> cash out by month
```

sheets: start here, dashboard, inputs, bom, components, production, requirements, purchasing, cash.

## rebuild

```sh
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python3 build_demo.py
```

edit `PRODUCTS`, `COMPONENTS` or `BOM` at the top of `build_demo.py` to model something else. all data is invented. built with claude code.
