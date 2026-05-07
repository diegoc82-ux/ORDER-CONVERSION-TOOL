"""
qb_server.py — QBWC SOAP listener for QuickBooks Desktop integration.
Fixed to handle QBWC 34.x namespace format and serverVersion/clientVersion calls.
"""
 
import os
import json
import logging
from datetime import datetime
from flask import Flask, request, Response
from qb_queue import init_db, get_next_pending, complete_job, fail_job, get_job, enqueue
from qb_xml import (
    build_sales_order_xml,
    build_sales_order_mod_xml,
    build_estimate_xml,
    build_inventory_query_xml,
    parse_response,
)
 
QB_SERVER_PORT = int(os.environ.get("QB_SERVER_PORT", 8443))
QBWC_USERNAME  = os.environ.get("QBWC_USERNAME", "u1p_qb_connector")
QBWC_PASSWORD  = os.environ.get("QBWC_PASSWORD", "ChangeMe123!")
SOAP_NS        = "http://developer.intuit.com/qbwc/rdp/QBWebConnectorSvc"
INTUIT_NS      = "http://developer.intuit.com/"
 
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [QB] %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("qb_server")
 
_session: dict = {"token": None, "job_id": None, "last_error": ""}
 
soap_app = Flask("qb_soap")
 
 
@soap_app.route("/qbwc", methods=["GET", "POST"])
def qbwc_endpoint():
    if request.method == "GET":
        return Response(_wsdl(), content_type="text/xml; charset=utf-8")
 
    body = request.data.decode("utf-8", errors="replace")
    log.info("QBWC → %d bytes", len(body))
 
    method = _detect_method(body)
    log.info("Method: %s", method)
 
    if method == "serverVersion":
        # QBWC asks what version of the web service we support
        resp = _soap_response("serverVersion",
            "<serverVersionResult><strVersion>2.0</strVersion></serverVersionResult>")
 
    elif method == "clientVersion":
        # QBWC tells us its version — we just say OK (empty string = no warning)
        resp = _soap_response("clientVersion",
            "<clientVersionResult><strVersion></strVersion></clientVersionResult>")
 
    elif method == "authenticate":
        resp = _handle_authenticate(body)
 
    elif method == "sendRequestXML":
        resp = _handle_send_request(body)
 
    elif method == "receiveResponseXML":
        resp = _handle_receive_response(body)
 
    elif method == "getLastError":
        resp = _handle_get_last_error(body)
 
    elif method == "closeConnection":
        resp = _handle_close_connection(body)
 
    else:
        log.warning("Unknown method — body snippet: %s", body[:200])
        resp = _soap_fault("Unknown method")
 
    return Response(resp, content_type="text/xml; charset=utf-8")
 
 
# ── Handlers ──────────────────────────────────────────────────────────────────
 
def _handle_authenticate(body: str) -> str:
    username = _extract(body, "strUserName")
    password = _extract(body, "strPassword")
    log.info("authenticate: user=%s", username)
 
    if username != QBWC_USERNAME or password != QBWC_PASSWORD:
        log.warning("authenticate: INVALID credentials")
        return _soap_response("authenticate",
            "<authenticateResult>"
            "<string>BAD_LOGIN</string>"
            "<string>nvu</string>"
            "</authenticateResult>")
 
    token = f"u1p_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}"
    _session["token"] = token
    _session["last_error"] = ""
    log.info("authenticate: OK token=%s", token)
 
    return _soap_response("authenticate",
        "<authenticateResult>"
        f"<string>{token}</string>"
        "<string></string>"
        "</authenticateResult>")
 
 
def _handle_send_request(body: str) -> str:
    log.info("sendRequestXML called")
    job = get_next_pending()
 
    if job is None:
        log.info("No pending jobs")
        return _soap_response("sendRequestXML",
            "<sendRequestXMLResult><string></string></sendRequestXMLResult>")
 
    _session["job_id"] = job["id"]
    payload  = json.loads(job["payload"])
    job_type = job["type"]
    log.info("Processing job %s type=%s", job["id"], job_type)
 
    try:
        if job_type == "sales_order":
            xml = build_sales_order_xml(job["id"], payload)
        elif job_type == "estimate":
            xml = build_estimate_xml(job["id"], payload)
        elif job_type == "inventory_query":
            xml = build_inventory_query_xml(job["id"], payload)
        elif job_type == "sales_order_mod":
            xml = build_sales_order_mod_xml(job["id"], payload)
        else:
            raise ValueError(f"Unknown job type: {job_type}")
    except Exception as e:
        log.error("Build failed: %s", e)
        fail_job(job["id"], str(e))
        _session["job_id"] = None
        return _soap_response("sendRequestXML",
            "<sendRequestXMLResult><string></string></sendRequestXMLResult>")
 
    return _soap_response("sendRequestXML",
        f"<sendRequestXMLResult><string>{_xml_escape(xml)}</string></sendRequestXMLResult>")
 
 
def _handle_receive_response(body: str) -> str:
    log.info("receiveResponseXML called")
    raw = _extract(body, "response")
    xml = _xml_unescape(raw or "")
    log.info("QB response XML (%d chars): %r", len(xml), xml[:500])
    job_id = _session.get("job_id")
 
    if job_id:
        try:
            result = parse_response(xml)
            if result["ok"]:
                # If the SO has group lines, enqueue a price-patch mod job
                # before marking this job complete.
                all_lines = result["data"].get("all_lines")
                if all_lines:
                    orig_job     = get_job(job_id)
                    orig_payload = json.loads(orig_job["payload"])
                    # Prices for group items in payload order
                    group_prices = [it["unit_price"]
                                    for it in orig_payload.get("line_items", [])
                                    if it.get("individual_sku")]
                    group_idx = 0
                    mod_lines = []
                    for al in all_lines:
                        if al["type"] == "individual":
                            mod_lines.append({
                                "type":        "individual",
                                "txn_line_id": al["txn_line_id"],
                            })
                        elif al["type"] == "group":
                            rate = group_prices[group_idx] if group_idx < len(group_prices) else 0
                            mod_lines.append({
                                "type":                  "group",
                                "txn_line_id":           al["txn_line_id"],
                                "component_txn_line_id": al["component_txn_line_id"],
                                "rate":                  rate,
                            })
                            group_idx += 1
                    if any(ml["type"] == "group" for ml in mod_lines):
                        mod_payload = {
                            "txn_id":        result["data"]["txn_id"],
                            "edit_sequence": result["data"]["edit_sequence"],
                            "mod_lines":     mod_lines,
                        }
                        enqueue("sales_order_mod", mod_payload)
                        log.info("Enqueued price-patch mod for SO %s (%d lines, %d groups)",
                                 result["data"]["txn_id"], len(mod_lines), group_idx)

                complete_job(job_id, result)
                log.info("Job %s DONE", job_id)
            else:
                fail_job(job_id, result.get("message", "QB error"))
                _session["last_error"] = result.get("message", "")
                log.error("Job %s FAILED: %s", job_id, result.get("message"))
        except Exception as e:
            fail_job(job_id, str(e))
            _session["last_error"] = str(e)
            log.error("receiveResponseXML exception: %s", e)
 
    _session["job_id"] = None
    return _soap_response("receiveResponseXML",
        "<receiveResponseXMLResult><int>100</int></receiveResponseXMLResult>")
 
 
def _handle_get_last_error(body: str) -> str:
    error = _session.get("last_error", "")
    log.info("getLastError: %s", error or "(none)")
    return _soap_response("getLastError",
        f"<getLastErrorResult><string>{error}</string></getLastErrorResult>")
 
 
def _handle_close_connection(body: str) -> str:
    log.info("closeConnection")
    _session["token"] = None
    _session["job_id"] = None
    _session["last_error"] = ""
    return _soap_response("closeConnection",
        "<closeConnectionResult><string>OK</string></closeConnectionResult>")
 
 
# ── SOAP helpers ──────────────────────────────────────────────────────────────
 
def _soap_response(method: str, inner: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope '
        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        "<soap:Body>"
        f'<{method}Response xmlns="{INTUIT_NS}">'
        + inner +
        f"</{method}Response>"
        "</soap:Body>"
        "</soap:Envelope>"
    )
 
 
def _soap_fault(message: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body><soap:Fault>"
        f"<faultstring>{message}</faultstring>"
        "</soap:Fault></soap:Body></soap:Envelope>"
    )
 
 
def _wsdl() -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        f'<definitions xmlns="http://schemas.xmlsoap.org/wsdl/" '
        f'targetNamespace="{INTUIT_NS}">'
        '<service name="QBWebConnectorSvc">'
        '<port name="QBWebConnectorSvcSoap" binding="QBWebConnectorSvcSoap">'
        f'<soap:address xmlns:soap="http://schemas.xmlsoap.org/wsdl/soap/" '
        f'location="http://localhost:{QB_SERVER_PORT}/qbwc"/>'
        "</port></service></definitions>"
    )
 
 
def _detect_method(body: str) -> str:
    """
    Detect SOAP method name from body.
    QBWC 34 sends: <methodName xmlns="http://developer.intuit.com/">
    So we look for the tag name ignoring namespace attributes.
    """
    methods = [
        "serverVersion", "clientVersion", "authenticate",
        "sendRequestXML", "receiveResponseXML",
        "getLastError", "closeConnection"
    ]
    for method in methods:
        # Match <method> or <method xmlns=...> or <ns:method>
        if (f"<{method}>" in body or
            f"<{method} " in body or
            f":{method}>" in body or
            f":{method} " in body):
            return method
    return "unknown"
 
 
def _extract(body: str, tag: str) -> str:
    """Extract text content of a tag, handling namespace prefixes."""
    # Try plain tag first
    for open_t in [f"<{tag}>", f"<{tag} "]:
        start = body.find(open_t)
        if start != -1:
            # Find the actual > to skip attributes
            gt = body.find(">", start)
            close_t = f"</{tag}>"
            end = body.find(close_t, gt)
            if end != -1:
                return body[gt + 1:end].strip()
 
    # Try namespaced tag e.g. <q1:tag>
    for line in body.splitlines():
        if f":{tag}>" in line or f":{tag} " in line:
            s = line.find(">") + 1
            e = line.rfind("<")
            if s < e:
                return line[s:e].strip()
    return ""
 
 
def _xml_escape(val: str) -> str:
    return (val.replace("&", "&amp;").replace("<", "&lt;")
               .replace(">", "&gt;").replace('"', "&quot;"))
 
 
def _xml_unescape(val: str) -> str:
    return (val.replace("&lt;", "<").replace("&gt;", ">")
               .replace("&amp;", "&").replace("&quot;", '"'))
 
 
# ── Entry point ───────────────────────────────────────────────────────────────
 
if __name__ == "__main__":
    init_db()
    log.info("QB SOAP server starting on port %d", QB_SERVER_PORT)
    log.info("QBWC endpoint → http://localhost:%d/qbwc", QB_SERVER_PORT)
    soap_app.run(host="0.0.0.0", port=QB_SERVER_PORT, debug=False)