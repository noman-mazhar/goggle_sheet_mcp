"""
Globelink Google Sheets MCP Server
-----------------------------------
Exposes one MCP tool: append_to_globelink_sheet
Claude calls this after extracting data from a PDF.
Writes one row to the configured Google Sheet.

Deploy on Render (free tier) as a web service.
"""

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import google.auth.transport.requests
import urllib.request
import urllib.error
import urllib.parse
import json
import os

load_dotenv()

app = FastAPI()

# ── CONFIG ────────────────────────────────────────────────────────────────────
SHEET_ID   = "1GiVtLEjCH1WLQa4B79DHUyIQN7TnY2PcufJEhsJEJ20"
SHEET_NAME = "Globelink Invoice Data"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets"]

# Load service account credentials from environment variable.
# Set GOOGLE_SERVICE_ACCOUNT_JSON to the full JSON content of the service account key file.
_sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
if not _sa_json:
    raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set")
SERVICE_ACCOUNT = json.loads(_sa_json)

# Fixed columns before PO numbers. PO#1, PO#2, ... are appended dynamically,
# followed by "Extracted At" as the last column.
BASE_HEADERS = [
    "File", "Entry Number", "BL Number", "Vessel", "Entry Date",
    "Origin", "Destination", "Pieces", "Weight",
    "Invoice Value", "Additional Duty %", "Duty Per Item %",
    "MPF & HMF %", "Freight", "Globelink Bill"
]
# ─────────────────────────────────────────────────────────────────────────────


def col_letter(n: int) -> str:
    """Convert 1-based column index to spreadsheet letter(s): 1→A, 26→Z, 27→AA."""
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
    url   = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}{urllib.parse.quote(path, safe='/?=&:!')}"
    data  = json.dumps(body).encode() if body else None
    req   = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {get_token()}",
            "Content-Type":  "application/json"
        }
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def ensure_headers(po_count: int):
    """
    Write or extend the header row to accommodate po_count PO# columns.
    Headers: BASE_HEADERS + [PO#1, PO#2, ...] + [Extracted At]
    If the current header row is shorter than needed, it is rewritten.
    """
    headers  = BASE_HEADERS + [f"PO#{i + 1}" for i in range(po_count)] + ["Extracted At"]
    last_col = col_letter(len(headers))
    result   = sheets_request("GET", f"/values/{SHEET_NAME}!A1:{last_col}1")
    existing = (result.get("values") or [[]])[0]
    if len(existing) < len(headers):
        sheets_request("PUT", f"/values/{SHEET_NAME}!A1:{last_col}1?valueInputOption=RAW", {
            "values": [headers]
        })


def append_row(base_values: list, po_numbers: list, extracted_at: str):
    ensure_headers(len(po_numbers))
    row      = base_values + po_numbers + [extracted_at]
    last_col = col_letter(len(row))
    sheets_request("POST", f"/values/{SHEET_NAME}!A:{last_col}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS", {
        "values": [row]
    })


# ── MCP PROTOCOL ─────────────────────────────────────────────────────────────

@app.get("/")
def health():
    return {"status": "ok", "service": "globelink-mcp"}


@app.post("/mcp")
async def mcp(request: Request):
    """
    Minimal MCP-compatible SSE/JSON endpoint.
    Handles:  tools/list  and  tools/call
    """
    try:
        body = await request.json()
    except Exception:
        return JSONResponse(
            {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error: invalid JSON"}},
            status_code=400
        )
    method = body.get("method", "")
    req_id = body.get("id", 1)

    # ── tools/list ────────────────────────────────────────────────────────────
    if method == "tools/list":
        return JSONResponse({
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "tools": [{
                    "name": "append_to_globelink_sheet",
                    "description": (
                        "Appends one row of extracted Globelink customs data "
                        "to the shared Google Sheet. Call this after extracting "
                        "all values from a Globelink PDF. Multiple PO numbers are "
                        "each written to their own cell (PO#1, PO#2, ...) in the same row."
                    ),
                    "inputSchema": {
                        "type": "object",
                        "required": ["file_name", "invoice_value", "additional_duty",
                                     "duty_per_item", "mpf_hmf", "freight",
                                     "globelink_bill", "po_number"],
                        "properties": {
                            "file_name":        {"type": "string",  "description": "PDF filename"},
                            "entry_number":     {"type": "string",  "description": "Entry # e.g. L93-00042335"},
                            "bl_number":        {"type": "string",  "description": "Master/House B/L number"},
                            "vessel":           {"type": "string",  "description": "Vessel / steamship name"},
                            "entry_date":       {"type": "string",  "description": "Date of entry MM/DD/YYYY"},
                            "origin":           {"type": "string",  "description": "Country of origin"},
                            "destination":      {"type": "string",  "description": "Port of destination"},
                            "pieces":           {"type": "string",  "description": "Number of pieces/cartons"},
                            "weight":           {"type": "string",  "description": "Gross weight in KG"},
                            "invoice_value":    {"type": "number",  "description": "Invoice value in USD"},
                            "additional_duty":  {"type": "number",  "description": "Additional duty % (e.g. 10.0)"},
                            "duty_per_item":    {"type": "number",  "description": "Duty per item % (e.g. 2.7)"},
                            "mpf_hmf":          {"type": "number",  "description": "MPF + HMF % combined"},
                            "freight":          {"type": "number",  "description": "Total freight excl. customs duty"},
                            "globelink_bill":   {"type": "number",  "description": "Balance Due / Globelink Bill"},
                            "po_number":        {"type": "array", "items": {"type": "string"},
                                                 "description": "One or more PO numbers — each gets its own cell (PO#1, PO#2, ...)"},
                            "extracted_at":     {"type": "string",  "description": "Timestamp of extraction"}
                        }
                    }
                }]
            }
        })

    # ── tools/call ────────────────────────────────────────────────────────────
    if method == "tools/call":
        tool_name = body.get("params", {}).get("name", "")
        args      = body.get("params", {}).get("arguments", {})

        if tool_name != "append_to_globelink_sheet":
            return JSONResponse({
                "jsonrpc": "2.0", "id": req_id,
                "error": {"code": -32601, "message": f"Unknown tool: {tool_name}"}
            })

        try:
            from datetime import datetime
            base_values = [
                args.get("file_name",      ""),
                args.get("entry_number",   ""),
                args.get("bl_number",      ""),
                args.get("vessel",         ""),
                args.get("entry_date",     ""),
                args.get("origin",         ""),
                args.get("destination",    ""),
                args.get("pieces",         ""),
                args.get("weight",         ""),
                args.get("invoice_value",  ""),
                args.get("additional_duty",""),
                args.get("duty_per_item",  ""),
                args.get("mpf_hmf",        ""),
                args.get("freight",        ""),
                args.get("globelink_bill", ""),
            ]
            po_numbers   = args.get("po_number", [])
            extracted_at = args.get("extracted_at", datetime.now().strftime("%Y-%m-%d %H:%M"))
            append_row(base_values, po_numbers, extracted_at)
            return JSONResponse({
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{
                        "type": "text",
                        "text": f"Row appended to Google Sheet for file: {args.get('file_name','')}"
                    }]
                }
            })
        except Exception as e:
            return JSONResponse({
                "jsonrpc": "2.0", "id": req_id,
                "error": {"code": -32000, "message": str(e)}
            }, status_code=500)

    # ── unknown method ────────────────────────────────────────────────────────
    return JSONResponse({
        "jsonrpc": "2.0", "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"}
    }, status_code=404)


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=False)
