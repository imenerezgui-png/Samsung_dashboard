# Samsung Campaigns Dashboard

KPI dashboard for the Samsung paid-social campaigns (Aug 1 – Nov 5, 2026), built by **Ykone**.

## Features

- **KPI cards** — budget, delivered spend, reach, impressions, engagements, ThruPlays, weighted CPM/CPC, ad recall lift.
- **Budget speedometer** — gauge of delivered spend in $ and % of the total budget, with a Samsung S26 Fe as the needle.
- **Highlights & take-aways** — auto-generated insight cards (top engagement, cheapest CPM/CPC, market split, video champion, pacing alerts).
- **Smart campaign-name parsing** — country (`TN`/`tn`/`Tn` → Tunisia, `LYB`/`Lyb`/`lyb` → Libya), budgets (`2000$`, `500 $`, `333 USD`, `1 066$`, `50$ … +50$`), platform (FB/IG) and objective extracted directly from the `Campaign name` column.
- **Filters** — country, platform, objective.
- **Charts** — budget by market, budget vs delivered spend by objective, impressions & engagement rate per campaign, CPM/CPC efficiency scatter, engagement mix.
- **Samsung-blue theme** featuring the Samsung and Ykone logos.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The Meta export `Samsung-Campaigns-Aug-1-2026-Nov-5-2026.xlsx` must sit next to `app.py`.
