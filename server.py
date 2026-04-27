from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import google.auth.transport.requests
import urllib.request
import urllib.parse
import json
import os

load_dotenv()
app = FastAPI()

SHEET_ID   = "1GiVtLEjCH1WLQa4B79DHUyIQN7TnY2PcufJEhsJEJ20"
SHEET_NAME = "Globelink Invoice Data"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets"]

_sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
if not _sa_json:
    raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set")
SERVICE_ACCOUNT = json.loads(_sa_json)

HEADERS = [
    "File",
    "Invoice Number",
    "Date",
    "Arrival Date",
    "Container Number",
    "Country",
    "Entry Number",
    "BL Number",
    "Vessel",
    "Entry Date",
    "Origin",
    "Destination",
    "Pieces",
    "Weight",
    "Invoice Value",
    "Additional Duty %",
    "Duty Per Item %",
    "MPF & HMF %",
    "Freight",
    "Globelink Bill",
    "PO Number",
    "Extracted At",
]


def col_letter(n):
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def get_token():
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT, scopes=SCOPES)
    creds.refresh(google.auth.transport.requests.Request())
    return creds.token


def sheets_request(method, path, body=None):
    url  = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}{urllib.parse.quote(path, safe='/?=&:!')}"
    data = json.dumps(body).encode() if body else None
    req  = urllib.request.Request(
        url, data=data, method=method,
        headers={"Authorization": f"Bearer {get_token()}", "Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def ensure_headers():
    last_col = col_letter(len(HEADERS))
    result   = sheets_request("GET", f"/values/{SHEET_NAME}!A1:{last_col}1")
    existing = (result.get("values") or [[]])[0]
    if existing != HEADERS:
        sheets_request("PUT", f"/values/{SHEET_NAME}!A1:{last_col}1?valueInputOption=RAW",
                       {"values": [HEADERS]})


def append_rows(rows):
    ensure_headers()
    last_col = col_letter(len(HEADERS))
    sheets_request(
        "POST",
        f"/values/{SHEET_NAME}!A:{last_col}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS",
        {"values": rows}
    )


@app.get("/")
def health():
    return {"status": "ok", "service": "globelink-mcp"}


@app.post("/mcp")
async def mcp(request: Request):
    try:
        body = await request.json()
    except Exception:
        return JSONResponse(
            {"jsonrpc": "2.0", "id": None,
             "error": {"code": -32700, "message": "Parse error: invalid JSON"}},
            status_code=400)

    method = body.get("method", "")
    req_id = body.get("id", 1)

    if method == "tools/list":
        return JSONResponse({
            "jsonrpc": "2.0", "id": req_id,
            "result": {
                "tools": [{
                    "name": "append_rows_to_globelink_sheet",
                    "description": (
                        "Appends one row per PO number to the shared Google Sheet. "
                        "All financial and metadata fields are duplicated across each PO row. "
                        "Pass a list of row objects — one per PO."
                    ),
                    "inputSchema": {
                        "type": "object",
                        "required": ["rows"],
                        "properties": {
                            "rows": {
                                "type": "array",
                                "description": "List of row objects — one per PO number",
                                "items": {
                                    "type": "object",
                                    "required": ["file_name", "po_number"],
                                    "properties": {
                                        # Core identifiers
                                        "file_name":         {"type": "string", "description": "PDF filename"},
                                        "invoice_number":    {"type": "string", "description": "Last number in PDF filename e.g. 44081"},
                                        "date":              {"type": "string", "description": "Invoice Date from invoice page e.g. 3/16/2026"},
                                        "arrival_date":      {"type": "string", "description": "Arrival Date from invoice page e.g. 3/13/2026"},
                                        "container_number":  {"type": "string", "description": "Container # from BOL page e.g. KMTU9361131"},
                                        "country":           {"type": "string", "description": "Derived country e.g. USA"},
                                        # Shipment metadata
                                        "entry_number":      {"type": "string", "description": "Customs entry number e.g. L93-00042996"},
                                        "bl_number":         {"type": "string", "description": "Bill of Lading number"},
                                        "vessel":            {"type": "string", "description": "Ship name e.g. KMTC SEOUL"},
                                        "entry_date":        {"type": "string", "description": "Date of customs entry"},
                                        "origin":            {"type": "string", "description": "Country of origin e.g. CHINA"},
                                        "destination":       {"type": "string", "description": "US port/city e.g. LOS ANGELES, CA"},
                                        "pieces":            {"type": "string", "description": "Number of cartons"},
                                        "weight":            {"type": "string", "description": "Gross weight in KG"},
                                        # Financial
                                        "invoice_value":     {"type": "number", "description": "Entered Value from Entry Summary"},
                                        "additional_duty":   {"type": "number", "description": "Sum of all 9903.xx.xx duty %"},
                                        "duty_per_item":     {"type": "number", "description": "Last non-9903 HTS duty % (null if FREE)"},
                                        "mpf_hmf":           {"type": "number", "description": "MPF % + HMF % combined"},
                                        "freight":           {"type": "number", "description": "All charges except Customs Duty"},
                                        "globelink_bill":    {"type": "number", "description": "Balance Due amount"},
                                        # PO + timestamp
                                        "po_number":         {"type": "string", "description": "PO number for this row"},
                                        "extracted_at":      {"type": "string", "description": "Extraction timestamp"},
                                    }
                                }
                            }
                        }
                    }
                }]
            }
        })

    if method == "tools/call":
        tool_name = body.get("params", {}).get("name", "")
        args      = body.get("params", {}).get("arguments", {})

        if tool_name != "append_rows_to_globelink_sheet":
            return JSONResponse({
                "jsonrpc": "2.0", "id": req_id,
                "error": {"code": -32601, "message": f"Unknown tool: {tool_name}"}
            })

        try:
            from datetime import datetime
            default_ts = datetime.now().strftime("%Y-%m-%d %H:%M")
            rows_input = args.get("rows", [])

            sheet_rows = []
            for row in rows_input:
                sheet_rows.append([
                    row.get("file_name",        ""),
                    row.get("invoice_number",   ""),
                    row.get("date",             ""),
                    row.get("arrival_date",     ""),
                    row.get("container_number", ""),
                    row.get("country",          ""),
                    row.get("entry_number",     ""),
                    row.get("bl_number",        ""),
                    row.get("vessel",           ""),
                    row.get("entry_date",       ""),
                    row.get("origin",           ""),
                    row.get("destination",      ""),
                    row.get("pieces",           ""),
                    row.get("weight",           ""),
                    row.get("invoice_value",    ""),
                    row.get("additional_duty",  ""),
                    row.get("duty_per_item",    ""),
                    row.get("mpf_hmf",          ""),
                    row.get("freight",          ""),
                    row.get("globelink_bill",   ""),
                    row.get("po_number",        ""),
                    row.get("extracted_at",     default_ts),
                ])

            append_rows(sheet_rows)
            first_file = rows_input[0].get("file_name", "") if rows_input else ""
            return JSONResponse({
                "jsonrpc": "2.0", "id": req_id,
                "result": {
                    "content": [{
                        "type": "text",
                        "text": f"{len(sheet_rows)} row(s) appended for file: {first_file}"
                    }]
                }
            })

        except Exception as e:
            return JSONResponse({
                "jsonrpc": "2.0", "id": req_id,
                "error": {"code": -32000, "message": str(e)}
            }, status_code=500)

    return JSONResponse({
        "jsonrpc": "2.0", "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"}
    }, status_code=404)


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=False)