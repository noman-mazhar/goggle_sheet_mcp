"""
Globelink Google Sheets MCP Server
-----------------------------------
Exposes one MCP tool: append_to_globelink_sheet
Claude calls this after extracting data from a PDF.
Writes one row to the configured Google Sheet.

Deploy on Render (free tier) as a web service.
"""

from flask import Flask, request, jsonify
from google.oauth2.service_account import Credentials
import google.auth.transport.requests
import urllib.request
import urllib.error
import json
import os

app = Flask(__name__)

# ── CONFIG ────────────────────────────────────────────────────────────────────
SHEET_ID   = "1GiVtLEjCH1WLQa4B79DHUyIQN7TnY2PcufJEhsJEJ20"
SHEET_NAME = "Globelink Data"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets"]

# Load service account credentials from environment variable.
# Set GOOGLE_SERVICE_ACCOUNT_JSON to the full JSON content of the service account key file.
_sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
if not _sa_json:
    raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set")
SERVICE_ACCOUNT = json.loads(_sa_json)

HEADERS = [
    "File", "Entry Number", "BL Number", "Vessel", "Entry Date",
    "Origin", "Destination", "Pieces", "Weight",
    "Invoice Value", "Additional Duty %", "Duty Per Item %",
    "MPF & HMF %", "Freight", "Globelink Bill", "Extracted At"
]
# ─────────────────────────────────────────────────────────────────────────────


def get_token():
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT, scopes=SCOPES)
    creds.refresh(google.auth.transport.requests.Request())
    return creds.token


def sheets_request(method, path, body=None):
    url   = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}{path}"
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


def ensure_headers():
    """Write header row if sheet is empty."""
    result = sheets_request("GET", f"/values/{SHEET_NAME}!A1:P1")
    if not result.get("values"):
        sheets_request("PUT", f"/values/{SHEET_NAME}!A1:P1?valueInputOption=RAW", {
            "values": [HEADERS]
        })


def append_row(row: list):
    ensure_headers()
    sheets_request("POST", f"/values/{SHEET_NAME}!A:P:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS", {
        "values": [row]
    })


# ── MCP PROTOCOL ─────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "globelink-mcp"})


@app.route("/mcp", methods=["POST"])
def mcp():
    """
    Minimal MCP-compatible SSE/JSON endpoint.
    Handles:  tools/list  and  tools/call
    """
    body   = request.get_json(force=True, silent=True) or {}
    method = body.get("method", "")
    req_id = body.get("id", 1)

    # ── tools/list ────────────────────────────────────────────────────────────
    if method == "tools/list":
        return jsonify({
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "tools": [{
                    "name": "append_to_globelink_sheet",
                    "description": (
                        "Appends one row of extracted Globelink customs data "
                        "to the shared Google Sheet. Call this after extracting "
                        "all 6 values from a Globelink PDF."
                    ),
                    "inputSchema": {
                        "type": "object",
                        "required": ["file_name", "invoice_value", "additional_duty",
                                     "duty_per_item", "mpf_hmf", "freight",
                                     "globelink_bill"],
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
            return jsonify({
                "jsonrpc": "2.0", "id": req_id,
                "error": {"code": -32601, "message": f"Unknown tool: {tool_name}"}
            })

        try:
            from datetime import datetime
            row = [
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
                args.get("extracted_at",   datetime.now().strftime("%Y-%m-%d %H:%M")),
            ]
            append_row(row)
            return jsonify({
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{
                        "type": "text",
                        "text": f"✓ Row appended to Google Sheet for file: {args.get('file_name','')}"
                    }]
                }
            })
        except Exception as e:
            return jsonify({
                "jsonrpc": "2.0", "id": req_id,
                "error": {"code": -32000, "message": str(e)}
            }), 500

    # ── unknown method ────────────────────────────────────────────────────────
    return jsonify({
        "jsonrpc": "2.0", "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"}
    }), 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)