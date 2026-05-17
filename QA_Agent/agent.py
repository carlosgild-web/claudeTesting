import os
import anthropic
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(BASE_DIR, 'HPIAdmins_v1.xlsx')

_client = None

_df = None
_system_prompt = None


def _load_df():
    global _df
    if _df is None:
        df = pd.read_excel(DATA_PATH)
        df['Date'] = pd.to_datetime(df['Date'])
        _df = df[df['DataGapFlag'] == False].copy()
    return _df


def _build_system_prompt(df: pd.DataFrame) -> str:
    regions = sorted(df['RegionName'].unique().tolist())
    date_min = df['Date'].min()
    date_max = df['Date'].max()

    first_avg = df[df['Date'] == date_min]['AveragePrice'].mean()
    last_avg  = df[df['Date'] == date_max]['AveragePrice'].mean()
    overall_growth = (last_avg / first_avg - 1) * 100

    first_snap  = df[df['Date'] == date_min]
    latest_snap = df[df['Date'] == date_max]

    region_lines = []
    for r in sorted(regions, key=lambda r: -(
        latest_snap[latest_snap['RegionName'] == r]['AveragePrice'].values[0] /
        first_snap[first_snap['RegionName'] == r]['AveragePrice'].values[0]
    )):
        p0  = float(first_snap[first_snap['RegionName'] == r]['AveragePrice'].values[0])
        p1  = float(latest_snap[latest_snap['RegionName'] == r]['AveragePrice'].values[0])
        vol = int(latest_snap[latest_snap['RegionName'] == r]['SalesVolume'].values[0])
        det = float(latest_snap[latest_snap['RegionName'] == r]['DetachedPrice'].values[0])
        sem = float(latest_snap[latest_snap['RegionName'] == r]['SemiDetachedPrice'].values[0])
        ter = float(latest_snap[latest_snap['RegionName'] == r]['TerracedPrice'].values[0])
        flt = float(latest_snap[latest_snap['RegionName'] == r]['FlatPrice'].values[0])
        g   = (p1 / p0 - 1) * 100
        region_lines.append(
            f"  {r}:\n"
            f"    Jan 1995: GBP {round(p0):,}  ->  Oct 2016: GBP {round(p1):,}  (+{round(g,1)}%)\n"
            f"    Oct 2016 volume: {vol:,} transactions\n"
            f"    Oct 2016 property prices: Detached GBP {round(det):,} | Semi GBP {round(sem):,} | "
            f"Terraced GBP {round(ter):,} | Flat GBP {round(flt):,}"
        )

    yearly = df.groupby(df['Date'].dt.year)['AveragePrice'].mean().round(0)
    yearly_lines = "  " + "\n  ".join(
        f"{yr}: GBP {int(price):,}" for yr, price in yearly.items()
    )

    region_yearly_lines = []
    for r in regions:
        rdf = df[df['RegionName'] == r]
        yr_avg = rdf.groupby(rdf['Date'].dt.year)['AveragePrice'].mean().round(0)
        row = "  " + r + ": " + " | ".join(
            f"{yr} GBP {int(p):,}" for yr, p in yr_avg.items()
        )
        region_yearly_lines.append(row)

    yoy_peaks = []
    for r in regions:
        rdf = df[df['RegionName'] == r].dropna(subset=['12mPctChange'])
        if len(rdf) == 0:
            continue
        idx_max = rdf['12mPctChange'].idxmax()
        idx_min = rdf['12mPctChange'].idxmin()
        yoy_peaks.append(
            f"  {r}: peak +{round(rdf.loc[idx_max,'12mPctChange'],1)}% "
            f"({rdf.loc[idx_max,'Date'].strftime('%b %Y')})  /  "
            f"trough {round(rdf.loc[idx_min,'12mPctChange'],1)}% "
            f"({rdf.loc[idx_min,'Date'].strftime('%b %Y')})"
        )

    return (
        "You are a data analyst assistant embedded in the UK House Price Index (HPI) dashboard.\n"
        "Your role is to answer questions ONLY about the data that powers this dashboard.\n"
        "If asked about unrelated topics, politely redirect to HPI data questions.\n\n"

        "=== DATASET OVERVIEW ===\n"
        f"Coverage : 6 English metropolitan regions, {date_min.strftime('%B %Y')} to {date_max.strftime('%B %Y')}\n"
        f"Regions  : {', '.join(regions)}\n"
        f"Records  : {len(df):,} (6 regions x monthly observations)\n"
        "Source   : UK Land Registry / HPI Official Statistics\n\n"

        "=== DATA COLUMNS ===\n"
        "Date                : Monthly observation (first of month)\n"
        "RegionName          : One of the 6 metropolitan regions\n"
        "Area                : Local authority area code/name\n"
        "AveragePrice        : Mean residential property price (GBP)\n"
        "SalesVolume         : Number of completed transactions\n"
        "1mPctChange         : Month-over-month % price change (null for first month)\n"
        "12mPctChange        : Year-over-year % price change (null for first 12 months)\n"
        "DetachedPrice       : Mean detached house price (GBP)\n"
        "SemiDetachedPrice   : Mean semi-detached house price (GBP)\n"
        "TerracedPrice       : Mean terraced house price (GBP)\n"
        "FlatPrice           : Mean flat/apartment price (GBP)\n\n"

        "=== MARKET OVERVIEW ===\n"
        f"Market average price Jan 1995 : GBP {round(first_avg):,}\n"
        f"Market average price Oct 2016 : GBP {round(last_avg):,}\n"
        f"Overall market growth (21 yrs): +{round(overall_growth,1)}%\n\n"

        "=== REGIONAL SUMMARY (Jan 1995 to Oct 2016, ranked by total growth) ===\n"
        + "\n".join(region_lines) + "\n\n"

        "=== MARKET AVERAGE PRICE BY YEAR (all regions combined) ===\n"
        + yearly_lines + "\n\n"

        "=== AVERAGE PRICE BY YEAR - PER REGION ===\n"
        + "\n".join(region_yearly_lines) + "\n\n"

        "=== YoY GROWTH PEAKS AND TROUGHS BY REGION ===\n"
        + "\n".join(yoy_peaks) + "\n\n"

        "=== DATA QUALITY NOTES ===\n"
        "- Nov-Dec 2016 rows excluded (SalesVolume = 0, DataGapFlag = True)\n"
        "- 1mPctChange null for Jan 1995 (first month - structural, not an error)\n"
        "- 12mPctChange null for Jan 1995 to Dec 1995 (first 12 months - structural)\n"
        "- No other data quality issues detected\n\n"

        "=== RESPONSE GUIDELINES ===\n"
        "- Format prices with pound sign and commas (e.g. £145,200).\n"
        "- Format percentages to 1 decimal place (e.g. +12.4%).\n"
        "- Be concise and precise; use bullet points for multi-part answers.\n"
        "- If you cite a figure, state the month/year it applies to.\n"
    )


def get_system_prompt() -> str:
    global _system_prompt
    if _system_prompt is None:
        _system_prompt = _build_system_prompt(_load_df())
    return _system_prompt


def _get_client():
    global _client
    if _client is None:
        _client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    return _client


def ask(question: str, history: list | None = None) -> str:
    """
    Ask the QA agent a question.
    history: list of {"role": "user"|"assistant", "content": str}
    Returns the assistant reply as a string.
    """
    hist = list(history or [])[-20:]
    messages = [
        {"role": m["role"], "content": m["content"]}
        for m in hist
    ]
    messages.append({"role": "user", "content": question})

    response = _get_client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        system=get_system_prompt(),
        messages=messages,
    )
    return response.content[0].text


if __name__ == "__main__":
    print("Loading data and building context...")
    print(get_system_prompt()[:600])
    print("\n[context ready]\n")
    print("Test: Which region had the highest average price in October 2016?")
    print(ask("Which region had the highest average price in October 2016?"))
