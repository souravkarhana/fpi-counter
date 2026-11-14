# FPI Holders Tracker

Nifty 500 ke har stock mein kitne FPIs (foreign funds) hold kar rahe hain, Excel mein.
Data: NSE Shareholding Pattern filings (quarterly, free).

Output: `fpi_holders.xlsx`

| Stock Name | Symbol | No. of FPI Holders | Quarter |
|---|---|---|---|
| HDFC Bank Ltd. | HDFCBANK | ... | ... |

## GitHub pe chalana
1. Repo ke **Actions** tab mein jao, "Update FPI holders Excel" chuno, **Run workflow** dabao.
2. 20-30 min baad `fpi_holders.xlsx` repo mein commit ho jayegi (aur run ke Artifacts mein bhi milegi).
3. Har mahine ki 1 tarikh ko ye apne aap chalti hai.

## Apne computer pe
```
pip install -r requirements.txt
python fpi_holders.py
```

## Dhyan do
- NSE kabhi GitHub ke servers ko block kar deta hai. Workflow fail ho to script apne computer pe chalao.
- Data quarterly update hota hai, isliye mahine mein ek baar chalana kaafi hai.
