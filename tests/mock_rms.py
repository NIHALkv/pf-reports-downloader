"""Mock of the IntelliPrism (RMS) reports pages, for testing rms_automation without the real site.

Mirrors the real structure: mat-card.report-card cards, a right-side mat-drawer filter
panel whose first date field (div.date-filter > input + div.calendar-btn) opens an
ngb-datepicker (month/year selects, day cells with aria-labels, SET DATES), a
"Report is generating" dialog, a Download Center table (mat-column-* cells) whose
row turns Success after READY_S seconds, and DOWNLOAD opening a new tab that
delivers the file. Only one report can generate at a time (buttons disabled).

    python tests/mock_rms.py [port] [ready_seconds]
"""
import io
import sys
import time
from datetime import datetime

from flask import Flask, jsonify, request, send_file

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 5066
READY_S = float(sys.argv[2]) if len(sys.argv) > 2 else 5
REPORTS = {
    "Request Profile Report": "Profile created date range",
    "Positions Report": "Profile created date range",
    "Candidate Evaluation Report": "Candidate shared date range",
    "IJP - Demand View Report": "Posted to IJP on",
    "IJP - Applicant View Report": "Posted to IJP on",
}
app = Flask(__name__)
jobs = []  # {"id", "name", "requested", "ready_at", "dates"}


def _busy():
    return any(time.time() < j["ready_at"] for j in jobs)


@app.post("/api/generate")
def generate():
    data = request.get_json()
    if _busy():
        return jsonify(error="busy"), 409
    jobs.insert(0, {"id": str(len(jobs) + 1), "name": data["name"], "dates": data["dates"],
                    "requested": datetime.now(), "ready_at": time.time() + READY_S})
    return jsonify(ok=True)


@app.get("/api/state")
def state():
    return jsonify(busy=_busy(), jobs=[{"name": j["name"], "dates": j["dates"]} for j in jobs])


@app.get("/rms/file-download")
def file_download():
    j = next(j for j in jobs if j["id"] == request.args["fileId"])
    data = io.BytesIO(f"{j['name']} | {j['dates']}".encode())
    return send_file(data, as_attachment=True, download_name=f"{j['name']}.xlsx")


@app.get("/rms/download-center")
def downloads():
    rows = []
    for j in jobs:
        status = "Success" if time.time() >= j["ready_at"] else "In Progress"
        req = j["requested"].strftime("%d %b %Y, %I:%M %p").lstrip("0") + " GMT+5:30"
        btn = (f'<button class="primary-btn" onclick="window.open(\'/rms/file-download?fileId={j["id"]}&fileType=REPORT\')">'
               'DOWNLOAD</button>' if status == "Success" else "")
        rows.append(f'<tr><td class="mat-column-name">{j["name"]}</td><td class="mat-column-source">Intelliprism Reports</td>'
                    f'<td class="mat-column-requeston">{req}</td><td class="mat-column-status">{status}</td>'
                    f'<td class="mat-column-filter">VIEW FILTERS</td><td class="mat-column-action">{btn}</td></tr>')
    # an old row for the same report, which must be ignored
    rows.append('<tr><td class="mat-column-name">Positions Report</td><td class="mat-column-source">x</td>'
                '<td class="mat-column-requeston">18 Nov 2026, 2:51 PM GMT+5:30</td><td class="mat-column-status">Success</td>'
                '<td class="mat-column-filter"></td><td class="mat-column-action"><button class="primary-btn" '
                'onclick="alert(1)">DOWNLOAD</button></td></tr>')
    return "<h2>Downloads</h2><table><tbody>" + "".join(rows) + "</tbody></table>"


PAGE = r"""<html><body>
<h2>Intelliprism Reports</h2>
<div id="cards"></div>
<mat-drawer class="right-side-nav" style="display:none">
  <div>Select filters to download report <mat-icon onclick="closeDrawer()">close</mat-icon></div>
  <div class="filter-sidebar-container"><div class="filter-content">
    <label id="lab"></label>
    <div class="date-filter main-filter-date dropdown">
      <div class="input-group"><input class="form-control" placeholder="dd-mm-yyyy" readonly id="f1">
        <div class="dropdown-toggle calendar-btn pointer" onclick="openCal()"><svg width="10" height="10"></svg></div></div>
      <div id="cal" style="display:none">
        <ngb-datepicker class="custom-datepicker">
          <select class="form-select" aria-label="Select month" id="m"></select>
          <select class="form-select" aria-label="Select year" id="y"></select>
          <div id="days"></div>
        </ngb-datepicker>
        <button onclick="clearDates()">CLEAR</button><button class="primary-btn w-100" onclick="setDates()">SET DATES</button>
      </div>
    </div>
    <label>Optional date</label>
    <div class="date-filter"><div class="input-group"><input class="form-control" placeholder="dd-mm-yyyy" readonly value="Start date to End date"></div></div>
  </div></div>
  <button class="primary-btn w-100" onclick="gen()">GENERATE REPORT</button>
</mat-drawer>
<mat-dialog-container id="dlg" style="display:none"><h3>Report is generating</h3>
  <button onclick="location.href='/rms/download-center'">GO TO DOWNLOADS</button></mat-dialog-container>
<script>
const REPORTS = __REPORTS__;
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const DAYL = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
let current = null, from = null, to = null;
const fmt = d => String(d.getDate()).padStart(2,"0") + "-" + MON[d.getMonth()] + "-" + d.getFullYear();
async function render() {
  const st = await (await fetch("/api/state")).json();
  document.getElementById("cards").innerHTML = Object.keys(REPORTS).map(n =>
    `<mat-card class="report-card"><div><h3> ${n} </h3></div><button class="secondary-btn" ${st.busy ? "disabled" : ""}
      onclick="openDrawer('${n}')"> GENERATE REPORT </button></mat-card>`).join("");
}
function openDrawer(n) {
  current = n; document.getElementById("lab").textContent = REPORTS[n] + "*";
  const t = new Date(), s = new Date(t); s.setDate(t.getDate() - 30);
  document.getElementById("f1").value = fmt(s) + " to " + fmt(t); from = to = null;
  setTimeout(() => document.querySelector("mat-drawer").style.display = "block", 400);
}
function closeDrawer() { document.querySelector("mat-drawer").style.display = "none"; }
function openCal() {
  const m = document.getElementById("m"), y = document.getElementById("y");
  m.innerHTML = MON.map((x,i) => `<option value="${i+1}">${x}</option>`).join("");
  y.innerHTML = Array.from({length: 11}, (_, i) => `<option value="${2016+i}">${2016+i}</option>`).join("");
  const t = new Date(); m.value = String(t.getMonth()+1); y.value = String(t.getFullYear());
  m.onchange = y.onchange = drawDays; drawDays();
  document.getElementById("cal").style.display = "block";
}
function drawDays() {
  const m = +document.getElementById("m").value - 1, y = +document.getElementById("y").value;
  const n = new Date(y, m+1, 0).getDate(); let h = "";
  for (let d = 1; d <= n; d++) { const dt = new Date(y, m, d);
    h += `<div role="gridcell" class="ngb-dp-day" aria-label="${DAYL[dt.getDay()]}, ${MONL[m]} ${d}, ${y}"><span class="custom-day" onclick="pick(${y},${m},${d})"> ${d} </span></div>`; }
  document.getElementById("days").innerHTML = h;
}
function pick(y, m, d) { const dt = new Date(y, m, d); if (!from || to) { from = dt; to = null; } else if (dt >= from) { to = dt; } else { from = dt; } }
function clearDates() { from = to = null; }
function setDates() { if (from && to) document.getElementById("f1").value = fmt(from) + " to " + fmt(to); document.getElementById("cal").style.display = "none"; }
async function gen() {
  const r = await fetch("/api/generate", {method: "POST", headers: {"Content-Type": "application/json"},
    body: JSON.stringify({name: current, dates: document.getElementById("f1").value})});
  closeDrawer();
  if (r.ok) document.getElementById("dlg").style.display = "block";
  render();
}
render();
</script></body></html>"""


@app.get("/rms/rms-reports")
def reports():
    import json
    return PAGE.replace("__REPORTS__", json.dumps(REPORTS))


if __name__ == "__main__":
    app.run(port=PORT)
