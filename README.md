# MLB Pitcher Strikeout Projections

A Python pipeline that projects how many strikeouts each probable MLB starter will record, grades itself against what actually happened, and benchmarks the results against sportsbook strikeout lines. I built it as a personal project to practice forecasting, data pipelines, and honest model evaluation.

## What it does

Every day the script:

1. Pulls the day's probable pitchers and each team's strikeout rate from the free MLB Stats API.
2. Projects strikeouts for every starter (model below).
3. Builds a multi-tab Excel workbook: Today's Slate, History Log, and Accuracy (plus a personal log tab that stays private).
4. Emails the workbook to me, and the next day pulls the actual results to grade yesterday's projections.

## The model

A deliberately simple, explainable baseline:

1. **Recency-weighted K/9** from the pitcher's last five starts (weights 40/30/20/7/3%).
2. Multiply by an expected-innings figure (5.5) to get a raw strikeout projection.
3. **Opponent adjustment:** scale by the opposing team's K% vs. a 22% league average, capped at +/-5%.
4. **Anti-inflation cap:** a projection can't exceed 2x the pitcher's recent average.

## Results (honest version)

Evaluated on 2,383 starts from March 31 to July 12, 2026. Run `python evaluate.py` to reproduce.

| Metric | Value |
|---|---|
| Mean absolute error | 1.99 Ks |
| RMSE | 2.51 Ks |
| Bias | +0.19 Ks (slightly over-projects) |
| Within 1 K / within 2 Ks | 33% / 59% |
| Baseline: always predict the league average (4.97 Ks) | 1.90 MAE |

On the 248 starts where I logged a sportsbook line:

| | MAE |
|---|---|
| Model | 1.99 |
| Sportsbook line | 1.80 |

When the model disagreed with the line by more than 0.5 Ks (191 starts, pushes excluded), it landed on the right side 113 times (59.2%).

**Takeaways**
- The model does **not** beat a "predict the average" baseline on raw error, and the sportsbook line is more accurate than the model. A single weighted K/9 plus a small opponent tweak leaves most of the variance unexplained.
- The directional signal is more interesting than the raw error: the model's disagreements with the line were right more often than not. That is a small sample, so I treat it as a lead, not a conclusion.
- Next steps I'd test: pitch-level features (whiff rate, pitch mix), umpire and park effects, projecting innings pitched instead of fixing it at 5.5, and calibrating the opponent adjustment instead of capping it by hand.

## Data notes

- Starts from late April through mid-July were reconstructed with `backfill_history.py`, which rebuilds each projection using only game logs from **before** that start (no lookahead). Sportsbook lines can't be recovered after the fact, so those rows have no line.
- Sportsbook lines were entered by hand, so only a subset of starts has one.
- `data/history.csv` is the full log of projections and results. Personal bet records are not included.

## Run it yourself

```bash
pip install -r requirements.txt
cp .env.example .env      # then add your Gmail app password (only needed for the email step)
python run_today.py       # builds today's workbook in Output/
python evaluate.py        # prints the accuracy numbers above
```

## Files

| File | Purpose |
|---|---|
| `run_today.py` | Daily runner: pulls data, projects, builds the Excel report, sends the email |
| `backfill_history.py` | One-time catch-up for missed days (no-lookahead reconstruction) |
| `evaluate.py` | Computes the accuracy metrics from `data/history.csv` |
| `data/history.csv` | Projections vs. actual strikeouts, 2026 season |

**Stack:** Python, MLB Stats API, openpyxl, SMTP.
