# TicketVault Report Processor

A Streamlit app that combines, deduplicates, and splits monthly TicketVault sales exports into two formatted Excel files — one per seller account (`ystickets` and `yitzknopf`).

## Features

- Upload one or more TicketVault export files (handles the 100k row export limit by accepting multiple files)
- Upload a company → account mapping file that can be updated over time
- Deduplicates rows across files automatically
- Excludes $0 sales
- Assigns any unlisted customer to "Offsite"
- Outputs two `.xlsx` files sorted by invoice date (newest first), fully formatted with centered cells, table borders, and a navy header row
- Warns if any company names are missing from the mapping file

## Project structure

```
ticketvault-processor/
├── app.py            # Streamlit UI
├── processor.py      # Core logic (loading, processing, Excel writing)
├── requirements.txt
├── .streamlit/
│   └── config.toml   # Theme
└── .gitignore
```

## Getting started

### 1. Clone the repo

```bash
git clone https://github.com/your-org/ticketvault-processor.git
cd ticketvault-processor
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run locally

```bash
streamlit run app.py
```

## Deploying to Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set **Main file path** to `app.py`
4. Click **Deploy**

## Mapping file format

The mapping file must be an `.xlsx` with exactly these two columns:

| Company Name     | Account    |
|------------------|------------|
| YS Tickets Spec  | ystickets  |
| Jacks YS         | ystickets  |
| Indiana Promotions | yitzknopf |
| The Ticket Guy   | yitzknopf  |

The app supports prefix matching — e.g. `Bearhawk - Dylan` will automatically match `Bearhawk Group`.

## Named customers

The following customers are mapped by name. Any other client value is labeled **Offsite**:

`TicketsNow` · `Gametime` · `Vivid Seats` · `StubHub` · `SeatGeek` · `GoTickets` · `TicketNetwork` · `TickPick` · `Ticket Evolution` · `Mercury`
