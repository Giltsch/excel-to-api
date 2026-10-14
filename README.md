# Excel → API

**Upload a spreadsheet, get a complete FastAPI project back** — data model, validation,
CRUD endpoints, seed data and Dockerfile, generated from the structure of your Excel file.

![tests](https://github.com/Giltsch/excel-to-api/actions/workflows/test.yml/badge.svg)

```
northwind.xlsx (13 sheets)  →  analyze  →  13 SQLAlchemy models, Pydantic schemas,
                                            CRUD routes, 7 foreign keys, seeded SQLite DB
```

## How it works

1. **Rule-based analysis** (`app/analyzer.py`) — no LLM required
   - Column types from header names *and* the data (int, float, bool, date, email, URL, enum …)
   - Nullability, value ranges and string lengths from the actual values
   - Primary keys by naming convention (`id`, `category_id`, `CategoryID`, `produkt_id`),
     only accepted if the values are actually unique; otherwise a surrogate key is generated
   - Foreign keys and relationships between sheets, including English and German
     singular/plural (`categories ↔ category`, `Produkte ↔ Produkt`)
   - Excel error values (`#N/A`, `#DIV/0!` …) are treated as empty cells
2. **Optional LLM review** (`app/llm_analyzer.py`) — sends a compact schema summary
   (column names, detected types, keys and up to 3 sample values per column, never the
   full sheet) to a model via OpenRouter to double-check types and relations.
3. **Code generation** (`app/generator.py` + Jinja2 `templates/`) — a standalone FastAPI
   project with SQLAlchemy models, Pydantic schemas, paginated CRUD endpoints, a seed
   script with your original rows, `requirements.txt`, `Dockerfile` and README.
4. **Live preview** — start the generated API inside the app and try it in Swagger UI,
   or download it as a ZIP.

## Quickstart

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app/main.py          # http://localhost:8501
```

Try it with `northwind.xlsx` (classic Northwind sample DB) or `examples/produkte.xlsx`
(small German demo, regenerate with `python create_demo_excel.py`).

The generated project runs on its own:

```bash
unzip generated_api.zip && cd generated_api
pip install -r requirements.txt
python seed_data.py                # loads your Excel rows into SQLite
uvicorn main:app --reload          # http://localhost:8000/docs
```

### Docker

```bash
cp .env.example .env
docker compose up -d --build       # http://localhost:8501
```

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `OPENROUTER_API_KEY` | – | Enables the optional LLM review |
| `OPENROUTER_MODEL` | `anthropic/claude-sonnet-4` | Model used for the review |
| `PUBLIC_MODE` | `false` | Hosted mode: users bring their own key, live preview disabled |
| `MAX_FILE_SIZE_MB`, `MAX_SHEETS`, `MAX_ROWS_PER_SHEET`, `RATE_LIMIT_UPLOADS` | see `.env.example` | Upload limits |

Uploads are validated before parsing (file signature, size, sheet/row/column limits) and
rate-limited per session. The container runs as a non-root user on a read-only filesystem.
See [DEPLOYMENT.md](DEPLOYMENT.md) for hosting behind Caddy with TLS.

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

## Roadmap

Next steps are messy real-world spreadsheets (merged cells, headers not in row 1, several
tables per sheet), CSV/JSON/Google Sheets input and composite keys — see
[roadmap.md](roadmap.md) (German).

## Stack

Python · Streamlit · FastAPI · SQLAlchemy · Pydantic · openpyxl · Jinja2 · Docker

## License

[MIT](LICENSE)
