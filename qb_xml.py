"""
qb_xml.py — QBXML request builders and response parsers.

Each function returns a complete QBXML string ready to hand to QBWC,
or parses the XML response QB sends back.

QB XML reference:
  https://developer.intuit.com/app/developer/qbdesktop/docs/api-reference
"""

import re
import xml.etree.ElementTree as ET
from datetime import date


# ── QBXML envelope ─────────────────────────────────────────────────────────

def _wrap(request_body: str) -> str:
    """Wrap a request element in the required QBXML envelope."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<?qbxml version="16.0"?>'
        "<QBXML>"
        "<QBXMLMsgsRq onError=\"stopOnError\">"
        + request_body
        + "</QBXMLMsgsRq>"
        "</QBXML>"
    )


# ── Sales Order (SalesOrderAdd) ─────────────────────────────────────────────

def build_sales_order_xml(job_id: str, payload: dict) -> str:
    """
    Build a SalesOrderAddRq QBXML string.

    Expected payload keys:
      customer_name  : str   — must match QB customer name exactly
      ref_number     : str   — your PO / order reference
      line_items     : list  — [{item_code, description, qty, unit_price}]
      memo           : str   — optional note on the order
      txn_date       : str   — 'YYYY-MM-DD', defaults to today
    """
    txn_date = payload.get("txn_date") or date.today().isoformat()
    memo = payload.get("memo", "")
    ref = payload.get("ref_number", "")
    customer = payload["customer_name"]

    lines = ""
    for item in payload.get("line_items", []):
        individual_sku = item.get("individual_sku")
        if individual_sku:
            # QB Group item — must use SalesOrderLineGroupAdd (not SalesOrderLineAdd).
            # Qty = number of cases; QB expands to individual units automatically.
            lines += (
                "<SalesOrderLineGroupAdd>"
                f"<ItemGroupRef><FullName>{_esc(item['item_code'])}</FullName></ItemGroupRef>"
                f"<Quantity>{item['qty']}</Quantity>"
                "</SalesOrderLineGroupAdd>"
            )
        else:
            lines += (
                "<SalesOrderLineAdd>"
                f"<ItemRef><FullName>{_esc(item['item_code'])}</FullName></ItemRef>"
                f"<Desc>{_esc(item.get('description', ''))}</Desc>"
                f"<Quantity>{item['qty']}</Quantity>"
                f"<Rate>{item['unit_price']}</Rate>"
                "</SalesOrderLineAdd>"
            )

    body = (
        f'<SalesOrderAddRq requestID="{_esc(job_id)}">'
        "<SalesOrderAdd>"
        f"<CustomerRef><FullName>{_esc(customer)}</FullName></CustomerRef>"
        f"<TxnDate>{txn_date}</TxnDate>"
        f"<RefNumber>{_esc(ref)}</RefNumber>"
        f"<Memo>{_esc(memo)}</Memo>"
        + lines
        + "</SalesOrderAdd>"
        "</SalesOrderAddRq>"
    )
    return _wrap(body)


# ── Sales Order price patch (SalesOrderMod) ────────────────────────────────

def build_sales_order_mod_xml(job_id: str, payload: dict) -> str:
    """
    Build a SalesOrderModRq that patches component Rate fields on group lines.

    Expected payload keys:
      txn_id        : str   — TxnID returned by the SalesOrderAdd
      edit_sequence : str   — EditSequence returned by the SalesOrderAdd
      group_mods    : list  — [{group_txn_line_id, component_txn_line_id, rate}]
    """
    txn_id   = payload["txn_id"]
    edit_seq = payload["edit_sequence"]

    # Replay all lines in the exact original order (QB error 3290 if out of order).
    # Individual lines get a bare TxnLineID — QB keeps them unchanged.
    # Group lines get the component Rate patched to the correct dispatch price.
    mods = ""
    for ml in payload.get("mod_lines", []):
        if ml["type"] == "individual":
            mods += (
                "<SalesOrderLineMod>"
                f"<TxnLineID>{_esc(ml['txn_line_id'])}</TxnLineID>"
                "</SalesOrderLineMod>"
            )
        elif ml["type"] == "group":
            mods += (
                "<SalesOrderLineGroupMod>"
                f"<TxnLineID>{_esc(ml['txn_line_id'])}</TxnLineID>"
                "<SalesOrderLineMod>"
                f"<TxnLineID>{_esc(ml['component_txn_line_id'])}</TxnLineID>"
                f"<Rate>{ml['rate']}</Rate>"
                "</SalesOrderLineMod>"
                "</SalesOrderLineGroupMod>"
            )

    body = (
        f'<SalesOrderModRq requestID="{_esc(job_id)}">'
        "<SalesOrderMod>"
        f"<TxnID>{_esc(txn_id)}</TxnID>"
        f"<EditSequence>{_esc(edit_seq)}</EditSequence>"
        + mods +
        "</SalesOrderMod>"
        "</SalesOrderModRq>"
    )
    return _wrap(body)


# ── Estimate / Quote (EstimateAdd) ──────────────────────────────────────────

def build_estimate_xml(job_id: str, payload: dict) -> str:
    """
    Build an EstimateAddRq QBXML string.
    Payload keys are identical to build_sales_order_xml.
    """
    txn_date = payload.get("txn_date") or date.today().isoformat()
    memo = payload.get("memo", "")
    ref = payload.get("ref_number", "")
    customer = payload["customer_name"]

    lines = ""
    for item in payload.get("line_items", []):
        lines += (
            "<EstimateLineAdd>"
            f"<ItemRef><FullName>{_esc(item['item_code'])}</FullName></ItemRef>"
            f"<Desc>{_esc(item.get('description', ''))}</Desc>"
            f"<Quantity>{item['qty']}</Quantity>"
            f"<Rate>{item['unit_price']}</Rate>"
            "</EstimateLineAdd>"
        )

    body = (
        f'<EstimateAddRq requestID="{_esc(job_id)}">'
        "<EstimateAdd>"
        f"<CustomerRef><FullName>{_esc(customer)}</FullName></CustomerRef>"
        f"<TxnDate>{txn_date}</TxnDate>"
        f"<RefNumber>{_esc(ref)}</RefNumber>"
        f"<Memo>{_esc(memo)}</Memo>"
        + lines
        + "</EstimateAdd>"
        "</EstimateAddRq>"
    )
    return _wrap(body)


# ── Inventory Query (ItemInventoryQuery) ────────────────────────────────────

def build_inventory_query_xml(job_id: str, payload: dict) -> str:
    """
    Build an ItemInventoryQueryRq QBXML string.

    Expected payload keys:
      item_codes : list[str] — list of QB item names / codes to query
    """
    filters = ""
    for code in payload.get("item_codes", []):
        filters += f"<FullName>{_esc(code)}</FullName>"

    body = (
        f'<ItemInventoryQueryRq requestID="{_esc(job_id)}">'
        + filters
        + "<IncludeRetElement>Name</IncludeRetElement>"
        "<IncludeRetElement>QuantityOnHand</IncludeRetElement>"
        "<IncludeRetElement>QuantityOnOrder</IncludeRetElement>"
        "<IncludeRetElement>QuantityOnSalesOrder</IncludeRetElement>"
        "</ItemInventoryQueryRq>"
    )
    return _wrap(body)


# ── Response parsers ────────────────────────────────────────────────────────

def parse_response(xml_str: str) -> dict:
    """
    Parse any QBXML response string from QB.
    Returns a dict with:
      ok        : bool
      status    : 'success' | 'error' | 'warning'
      message   : human-readable summary
      data      : parsed payload (txn_id for orders, items list for inventory)
    """
    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError as e:
        # Regex fallback: ET can't parse QB's response (bad char in error message).
        # Extract statusCode + statusMessage without a full XML parser so we at
        # least surface the real QB error to the user.
        sc = re.search(r'statusCode=["\'](\d+)["\']', xml_str)
        sm = re.search(r'statusMessage=["\']([^"\']{0,400})["\']', xml_str)
        code = sc.group(1) if sc else "1"
        msg  = sm.group(1) if sm else f"XML parse error: {e}"
        return {
            "ok": code == "0",
            "status": "success" if code == "0" else "error",
            "message": msg,
            "data": {},
        }

    ns = ""
    status_code = "0"
    status_msg = ""
    data = {}

    # Find the first *Ret or *Rs element
    for elem in root.iter():
        tag = elem.tag

        # Grab status from any *Rs element
        if tag.endswith("Rs"):
            status_code = elem.get("statusCode", "0")
            status_msg = elem.get("statusMessage", "")

        # Sales order created
        if tag == "SalesOrderRet":
            data["txn_id"]        = _find_text(elem, "TxnID")
            data["edit_sequence"] = _find_text(elem, "EditSequence")
            data["ref_number"]    = _find_text(elem, "RefNumber")
            data["type"]          = "sales_order"
            # Walk SalesOrderRet children IN ORDER — QB error 3290 fires if the
            # mod request sends lines out of sequence.  We build one flat ordered
            # list so the mod can replay them in exactly the right order.
            has_groups = False
            all_lines  = []
            for child in elem:
                if child.tag == "SalesOrderLineRet":
                    lid = _find_text(child, "TxnLineID")
                    if lid:
                        all_lines.append({"type": "individual", "txn_line_id": lid})
                elif child.tag == "SalesOrderLineGroupRet":
                    has_groups = True
                    grp_lid  = _find_text(child, "TxnLineID")
                    comp_lid = ""
                    for comp in child.findall("SalesOrderLineRet"):
                        comp_lid = _find_text(comp, "TxnLineID")
                        break   # always 1 component
                    all_lines.append({
                        "type":                  "group",
                        "txn_line_id":           grp_lid,
                        "component_txn_line_id": comp_lid,
                    })
            if has_groups:
                data["all_lines"] = all_lines

        # Sales order modified (price patch job)
        if tag == "SalesOrderModRet":
            data["txn_id"] = _find_text(elem, "TxnID")
            data["type"]   = "sales_order_mod"

        # Estimate created
        if tag == "EstimateRet":
            data["txn_id"] = _find_text(elem, "TxnID")
            data["ref_number"] = _find_text(elem, "RefNumber")
            data["type"] = "estimate"

        # Inventory items
        if tag == "ItemInventoryRet":
            items = data.setdefault("items", [])
            items.append({
                "name": _find_text(elem, "Name"),
                "qty_on_hand": _find_text(elem, "QuantityOnHand"),
                "qty_on_order": _find_text(elem, "QuantityOnOrder"),
                "qty_on_so": _find_text(elem, "QuantityOnSalesOrder"),
            })

    ok = status_code == "0"
    return {
        "ok": ok,
        "status": "success" if ok else "error",
        "message": status_msg,
        "data": data,
    }


# ── Helpers ─────────────────────────────────────────────────────────────────

def _esc(val) -> str:
    """Escape a value for safe inclusion in QBXML element content or attributes.
    - Escapes the five XML special characters
    - Strips control characters that are illegal in XML 1.0
    - Encodes non-ASCII characters as XML numeric character references so
      QB's parser never sees raw Unicode (avoids 'invalid token' errors)
    """
    s = str(val)
    # Standard XML escaping
    s = (s.replace("&", "&amp;")
          .replace("<", "&lt;")
          .replace(">", "&gt;")
          .replace('"', "&quot;")
          .replace("'", "&apos;"))
    # Strip XML 1.0 illegal control characters
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', s)
    # Encode non-ASCII as numeric character references (e.g. ® → &#174;)
    s = ''.join(c if ord(c) < 128 else f'&#{ord(c)};' for c in s)
    return s


def _find_text(elem, tag: str) -> str:
    child = elem.find(tag)
    return child.text or "" if child is not None else ""
