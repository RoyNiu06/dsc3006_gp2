# DSC3006 Final Project — Group 2

**Used-Car Price Prediction: Temporal Validation and Error Analysis**

| | |
|---|---|
| Course | DSC3006 Statistical Machine Learning (2026/27 Semester A) |
| Group | Group 2 |
| Members | GENG Xi (59282183), LUO Yue (59282392), NIU Ziru (59273230) |
| Date | 29 November 2026 |

---

## 1. What we study

We predict the auction `sellingprice` of used cars from **auction-time attributes only** (model year, odometer, make, model, trim, body, transmission, state, colour, interior, condition). The dataset is the public Kaggle **Vehicle Sales Data**: 558,837 raw records, 558,592 after cleaning (245 rows removed for invalid price, date, state, year or odometer).

The goal is deliberately **not** to chase the highest R². The project asks three questions:

1. **How much of the price can auction-time attributes explain?**
2. **Does a more flexible model beat a linear baseline?**
3. **Where should we not trust the prediction?**

MMR (an existing price estimate) and VIN are excluded as predictors; seller and sale date are excluded as identity/split variables.

## 2. Validation design

The deployment question is about a **future** period, so we use a **temporal holdout** rather than random K-fold:

| Split | Period | Rows |
|---|---|---|
| Train | Dec 2014 – Feb 2015 | 153,684 |
| Validation | Mar – Apr 2015 | 354,936 |
| Test | May – Jul 2015 | 47,526 |

Random K-fold is rejected because it would put future prices — and the same vehicle — into training. VIN overlap between splits is **zero** (2,245 train rows and 201 validation rows were purged to keep vehicles disjoint). The test set is evaluated once.

## 3. Models

A transparent model ladder, each step added only if it earns its place on validation log RMSE:

- **Mean baseline** — reference floor
- **Ridge** — regularised linear model on the same predictors
- **Spline Ridge** — adds non-linear terms in the numeric predictors
- **HGB31** — histogram gradient boosting (selected on validation log RMSE, then refit on all eligible pre-May records)

## 4. Repository layout

```
report/   DSC3006_Final_Project_Group2.pdf   # 10-page final report
slides/   DSC3006_VehicleSales_Slides.pdf    # presentation deck (PDF)
          DSC3006_VehicleSales_Slides.pptx   # presentation deck (editable)
scripts/  build_vehicle_slides_v3.py         # regenerates the deck from source tables
```

## 5. Reproducing the deck

```bash
pip install python-pptx pillow
python scripts/build_vehicle_slides_v3.py
```

The script writes `DSC3006_VehicleSales_Slides.pptx` next to itself.

## 6. Data

The Kaggle Vehicle Sales Data dataset is **public** and is **not redistributed** here. Download it from Kaggle (`vehicle-sales-data`) and place the CSV where your notebook expects it. All figures quoted in the report are computed from the public release, so the pipeline is reproducible from the original download.

## 7. Scope and limitations (short version)

- Prices are **not** claimed to be causal effects of the attributes.
- The model is **not** a valuation product and is not claimed to transfer to the 2026 market.
- A UK replication was considered but not included in the main analysis, because its price definition and timestamps differ from the US auction data.
