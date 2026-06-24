# AGENTS.md — proyecto_tosi

## Project Overview
Python desktop app (Flask + pywebview) for Stellantis autoparts inventory management. Single-file backend (`main.py`) + single-file frontend (`frontend/index.html`). SQLite database (`datos_inventario.db`) created at runtime.

## Run Commands
```bash
python main.py                    # Launch desktop app (auto-installs deps via pip)
python -m pytest                  # No test suite exists
```

## Architecture
- **Backend**: Flask on port 7777 (threaded), serves API + static frontend
- **Frontend**: Vanilla HTML/CSS/JS (no build step), loaded via pywebview
- **DB**: SQLite file `datos_inventario.db` in CWD, two tables:
  - `inventario` — container stock (Fuente A: plant export)
  - `piezas` — supplier parts with daily flow (Fuente B: .xlsm)
- **Entry point**: `main.py` → `verificar_e_instalar_dependencias()` → Flask + pywebview

## Key API Endpoints
| Method | Path | Purpose |
|--------|------|---------|
| POST | `/procesar` | Import Fuente A (Excel → `inventario` table) |
| POST | `/procesar_piezas` | Import Fuente B (.xlsm → `piezas` table) |
| GET | `/contenedores` | List unique containers |
| GET | `/contenedores/<id>/piezas` | Parts in a container |
| GET | `/piezas` | All parts from Fuente B |
| GET | `/plan_apertura` | Compute opening plan (max 10 containers/day) |

## Import Logic (Critical)
- **Fuente A** (`/procesar`): Filters rows where `DISPONIBLE=SI`, `EN PLANTA=SI`, `PAGADO=SI`. Saves `NUM_CONTENEDOR`, `PTO_DESCARGA`, `PLANO`, `DESC_PLANO`, `CANTIDAD`.
- **Fuente B** (`/procesar_piezas`): Auto-detects sheet/header row; maps columns `Ref`, `Designação`, `Flujo Espe`, `Darsena`, `MAG + BDL`, plus today/tomorrow flow columns by date matching.
- Both use `df.to_sql(..., if_exists='replace')` — **full replace** on each import.

## Plan de Apertura Algorithm (Optimized)
1. Join `piezas` (stock + flow today/tomorrow) with `inventario` (containers per `PLANO`=Ref)
2. Compute deficit = (flow_today + flow_tomorrow) - stock per part
3. Calculate urgency = deficit / total_need, coverage_hours = stock / (total_need/16h)
4. **Global greedy selection (max coverage)**: Build pool of all available containers across all parts. Iteratively pick the container with highest **marginal deficit coverage weighted by urgency** (gain = min(container_qty, remaining_deficit) * urgency). Repeat up to 10 containers.
5. Returns prioritized list (selection order = priority) with `estado` (red/amber/green) based on coverage hours

**Key improvement**: Instead of per-part greedy then global limit, the algorithm now globally maximizes total deficit units covered within the 10-container constraint. This avoids wasting containers on parts where a single container barely covers deficit, favoring containers that fully/nearly cover their part's deficit.

## Dependencies (auto-installed)
- `pywebview`, `Flask`, `pandas`, `openpyxl`, `xlrd`

## Frontend Conventions
- Single-page app with screen switching (`#login-screen`, `#app-screen`)
- CSS custom properties for theming (light/dark via `html.dark`)
- No framework — vanilla ES6 modules not used; all JS in `<script>` block
- Fetch calls to relative paths (e.g., `/plan_apertura`)

## Gotchas
- **Port 7777** hardcoded in both Flask (`app.run(port=7777)`) and pywebview URL
- **No migrations** — schema changes require deleting `datos_inventario.db`
- **Column names case-sensitive** in Excel imports; backend normalizes to uppercase
- **Date parsing** in Fuente B handles multiple formats (`%d/%m`, `%d-%m`, `%Y-%m-%d`, etc.)
- **No authentication** beyond simple login screen (stores initials in localStorage)
- **No tests** — verify manually via UI

## File Structure
```
proyecto_tosi/
├── main.py                 # Backend + launcher (706 lines)
├── frontend/
│   └── index.html          # Full frontend (1816 lines, inline CSS/JS)
├── datos_inventario.db     # SQLite DB (created at runtime)
├── excel/                  # Sample Excel files (reference only)
└── *.py                    # Inspection scripts (dump_sheets.py, etc.)
```

## Common Tasks
- **Add API endpoint**: Edit `main.py` Flask routes, update frontend `fetch()` calls
- **Change DB schema**: Delete `datos_inventario.db`, update `to_sql` columns
- **Modify import logic**: Edit `/procesar` or `/procesar_piezas` in `main.py`
- **Adjust plan algorithm**: Edit `/plan_apertura` route (lines 488-725)
- **Update UI**: Edit `frontend/index.html` (CSS in `<style>`, JS in `<script>`)