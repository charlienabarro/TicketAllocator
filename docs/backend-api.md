# Ticket Bundle Splitter API

## Run

```bash
pip install -r requirements.txt
uvicorn backend.app:app --reload --reload-dir .
```

Open [http://127.0.0.1:8000](http://127.0.0.1:8000).

## Core workflow endpoints

- `POST /ticket-bundles/preview`
- `POST /ticket-bundles/generate`

Both endpoints require multipart form-data fields:
- `allocation_csv` (file)
- `tickets_pdf` (file)

`allocation_csv` can be either:
- CSV file (`.csv`)
- Apple Numbers file (`.numbers`) converted server-side via Numbers on macOS

For `.numbers` uploads:
- Apple Numbers must be installed on the same Mac running the API.
- The first run may prompt for Automation permissions for `osascript` controlling Numbers.

## Allocation CSV expectations

Required logical columns (aliases supported):
- Booking reference: `Booking Ref` / `booking_reference`
- Email: `Email` / `customer email`
- Seats: `Assigned Seats` / `seats`

Optional:
- `Customer Name`

Seat list format can include:
- single seats: `C1 C2`
- comma/semicolon lists: `C1, C2; C3`
- ranges: `C1-C5`
- `to` ranges: `C1 to C5` or `C1 to 5`

## Output

`POST /ticket-bundles/generate` returns `ticket_bundle.zip` containing:
- one PDF per booking
- one unsigned `.pkpass` per matched ticket page under `wallet/`
- `manifest.csv` with booking ref, customer, email, seats, page numbers, missing seats

`POST /ticket-bundles/preview` returns one row per booking with:
- grouped PDF metadata and inline PDF data
- `wallet_passes`: one entry per parsed ticket page, each including:
  - `seat_label`
  - `pass_file`
  - `pass_url`
  - `pass_download_url`
  - `pass_data_url`
