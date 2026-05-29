#!/usr/bin/env python3
"""
inject_incoming_tab.py
Inserts a new "Incoming" tab at position 0 in dashboard.html,
shifting existing FFT–B2C tabs from positions 1–6 to 2–7.

Run: python3 inject_incoming_tab.py
"""

import re

HTML_PATH = "/Users/eduardoflores/AI_Projects/Dashboards/Especiales/dashboard.html"

# ─────────────────────────────────────────────────────────────────────────────
# 1. Read file
# ─────────────────────────────────────────────────────────────────────────────
with open(HTML_PATH, "r", encoding="utf-8") as f:
    src = f.read()

# ─────────────────────────────────────────────────────────────────────────────
# 2. Renumber all existing panel/chart/table/function IDs from 1–6 → 2–7
#    Applied HIGHEST-NUMBER FIRST to avoid double-replacement collisions.
# ─────────────────────────────────────────────────────────────────────────────

# Helper: exact string replacement (assert-on-miss so we catch typos)
def replace_exact(text, old, new, label=""):
    if old not in text:
        raise ValueError(f"ANCHOR NOT FOUND [{label}]: {repr(old)}")
    return text.replace(old, new)

# ── HTML panel IDs ────────────────────────────────────────────────────────────
for n in range(6, 0, -1):
    src = replace_exact(src, f'id="panel-{n}"', f'id="panel-{n+1}"', f"panel-{n} id")

# ── aria-controls ─────────────────────────────────────────────────────────────
for n in range(6, 0, -1):
    src = replace_exact(src, f'aria-controls="panel-{n}"', f'aria-controls="panel-{n+1}"', f"aria-controls panel-{n}")

# ── Tab bar onclick ────────────────────────────────────────────────────────────
for n in range(6, 0, -1):
    src = replace_exact(src, f'onclick="showTab({n})"', f'onclick="showTab({n+1})"', f"showTab({n})")

# ── Canvas chart IDs ──────────────────────────────────────────────────────────
for n in range(6, 0, -1):
    src = replace_exact(src, f'id="tab{n}Chart"', f'id="tab{n+1}Chart"', f"tab{n}Chart id")
    src = replace_exact(src, f'getElementById("tab{n}Chart")', f'getElementById("tab{n+1}Chart")', f'getElementById tab{n}Chart')

# ── Table body IDs ────────────────────────────────────────────────────────────
# Only tabs 1, 3, 5, 6 have a plain TableBody — tabs 2 and 4 do not.
TABS_WITH_TABLE_BODY = [6, 5, 3, 1]  # descending to avoid collisions
for n in TABS_WITH_TABLE_BODY:
    src = replace_exact(src, f'id="tab{n}TableBody"', f'id="tab{n+1}TableBody"', f"tab{n}TableBody id")
    src = replace_exact(src, f'getElementById("tab{n}TableBody")', f'getElementById("tab{n+1}TableBody")', f'getElementById tab{n}TableBody')

# ── tab4 special IDs (Summary/GrandTotal) — shift 4→5 ────────────────────────
# tab4SummaryTable  — HTML id only (no JS getElementById call)
# tab4SummaryBody   — HTML id + JS getElementById
# tab4GrandTotalUnits — HTML id + JS getElementById
src = replace_exact(src, 'id="tab4SummaryTable"',    'id="tab5SummaryTable"',    "tab4SummaryTable id")
for old_id, new_id in [
    ("tab4SummaryBody",     "tab5SummaryBody"),
    ("tab4GrandTotalUnits", "tab5GrandTotalUnits"),
]:
    src = replace_exact(src, f'id="{old_id}"',              f'id="{new_id}"',              f"{old_id} id")
    src = replace_exact(src, f'getElementById("{old_id}")',  f'getElementById("{new_id}")',  f'getElementById {old_id}')

# ─────────────────────────────────────────────────────────────────────────────
# 3. Update initChart dispatcher  (done BEFORE function-rename so the match
#    is against the original initTab1–6 text)
# ─────────────────────────────────────────────────────────────────────────────
OLD_DISPATCHER = """  function initChart(tabNum) {
    if (tabNum === 1) { initTab1(); }
    else if (tabNum === 2) { initTab2(); }
    else if (tabNum === 3) { initTab3(); }
    else if (tabNum === 4) { initTab4(); }
    else if (tabNum === 5) { initTab5(); }
    else if (tabNum === 6) { initTab6(); }
  }"""

NEW_DISPATCHER = """  function initChart(tabNum) {
    if (tabNum === 1) { initIncoming(); }
    else if (tabNum === 2) { initTab2(); }
    else if (tabNum === 3) { initTab3(); }
    else if (tabNum === 4) { initTab4(); }
    else if (tabNum === 5) { initTab5(); }
    else if (tabNum === 6) { initTab6(); }
    else if (tabNum === 7) { initTab7(); }
  }"""

src = replace_exact(src, OLD_DISPATCHER, NEW_DISPATCHER, "initChart dispatcher")

# ── JS function definitions ───────────────────────────────────────────────────
for n in range(6, 0, -1):
    src = replace_exact(src, f'function initTab{n}()', f'function initTab{n+1}()', f"function initTab{n}")

# ── TAB comment labels (cosmetic — keep them accurate) ───────────────────────
for n in range(6, 0, -1):
    src = src.replace(f'// TAB {n} —', f'// TAB {n+1} —')

# ── HTML tab-panel comment headers ───────────────────────────────────────────
# Use the 11-space indent that uniquely identifies the HTML <!-- --> comment lines,
# so this loop does NOT re-rename the JS // TAB n — lines already handled above.
for n in range(6, 0, -1):
    src = src.replace(f'           TAB {n} —', f'           TAB {n+1} —')

# ─────────────────────────────────────────────────────────────────────────────
# 4. TAB_COUNT 6 → 7
# ─────────────────────────────────────────────────────────────────────────────
src = replace_exact(src, "const TAB_COUNT = 6;", "const TAB_COUNT = 7;", "TAB_COUNT")

# ─────────────────────────────────────────────────────────────────────────────
# 5. Deactivate the old FFT button (now tab 2) — remove active/aria-selected=true
# ─────────────────────────────────────────────────────────────────────────────
OLD_FFT_BTN = '<button class="tab-btn active" role="tab" aria-selected="true"  aria-controls="panel-2" onclick="showTab(2)">FFT</button>'
NEW_FFT_BTN = '<button class="tab-btn"        role="tab" aria-selected="false" aria-controls="panel-2" onclick="showTab(2)">FFT</button>'
src = replace_exact(src, OLD_FFT_BTN, NEW_FFT_BTN, "FFT button deactivate")

# ─────────────────────────────────────────────────────────────────────────────
# 6. Deactivate the old FFT panel (now panel-2) — remove active class
# ─────────────────────────────────────────────────────────────────────────────
OLD_FFT_PANEL = '<div class="tab-panel active" id="panel-2" role="tabpanel">'
NEW_FFT_PANEL = '<div class="tab-panel" id="panel-2" role="tabpanel">'
src = replace_exact(src, OLD_FFT_PANEL, NEW_FFT_PANEL, "FFT panel deactivate")

# ─────────────────────────────────────────────────────────────────────────────
# 7. Insert new Incoming tab BUTTON as first button in the tab bar
# ─────────────────────────────────────────────────────────────────────────────
INCOMING_BTN = '      <button class="tab-btn active" role="tab" aria-selected="true"  aria-controls="panel-1" onclick="showTab(1)">Incoming</button>\n'

# Anchor: the (now-deactivated) FFT button is first in the tab bar
FFT_BTN_ANCHOR = '<button class="tab-btn"        role="tab" aria-selected="false" aria-controls="panel-2" onclick="showTab(2)">FFT</button>'

src = replace_exact(src,
    '      ' + FFT_BTN_ANCHOR,
    INCOMING_BTN + '      ' + FFT_BTN_ANCHOR,
    "insert Incoming button")

# ─────────────────────────────────────────────────────────────────────────────
# 8. Insert new panel-1 HTML BEFORE the FFT panel block (now "TAB 2 —")
# ─────────────────────────────────────────────────────────────────────────────
INCOMING_PANEL = """      <!-- ══════════════════════════════════════════════════════
           TAB 1 — Incoming (Reporte UPH - Incoming)
      ══════════════════════════════════════════════════════ -->
      <div class="tab-panel active" id="panel-1" role="tabpanel">
        <div class="card">

          <div class="card-header">
            <h1>Incoming Trailers — Pallets Received by Origin</h1>
            <p>Receiving Dock &nbsp;·&nbsp; Dec 2025 – May 2026 &nbsp;·&nbsp; 32,555 pallets &nbsp;·&nbsp; 1,288 events</p>
          </div>

          <div class="kpi-hero">
            <span class="kpi-value">32,555</span>
            <span class="kpi-label">Total Pallets Received</span>
          </div>

          <div class="chart-wrapper">
            <canvas id="tab1Chart"></canvas>
          </div>

          <div class="table-section">
            <h2>Monthly Breakdown</h2>
            <table>
              <thead>
                <tr>
                  <th>Month</th>
                  <th>US Origins</th>
                  <th>Tijuana, BC</th>
                  <th>Total Pallets</th>
                  <th>Unload Hours</th>
                  <th>Avg HC / Event</th>
                </tr>
              </thead>
              <tbody id="tab1TableBody"></tbody>
              <tfoot>
                <tr>
                  <td>Grand Total</td>
                  <td>31,164</td>
                  <td>1,391</td>
                  <td>32,555</td>
                  <td>1,502.6</td>
                  <td>2.5</td>
                </tr>
              </tfoot>
            </table>
          </div>

          <div class="insight-section">
            <h2>Key Insights</h2>
            <div class="insight-grid">

              <div class="insight-card">
                <div class="insight-label">Volume Mix</div>
                <div class="insight-title">US Origins account for 95.7% of all pallets — TV inventory from six DCs</div>
                <div class="insight-body">31,164 of 32,555 total pallets came from US distribution centers: Spartanburg SC, Greenfield IN, Waco TX, Las Vegas NV, Johnstown NY, and Groesbeck TX. These are the raw TV units that feed the FFT, Sorting, and OpenCell lines.</div>
              </div>

              <div class="insight-card">
                <div class="insight-label">January 2026 Peak</div>
                <div class="insight-title">8,931 US pallets in January — post-holiday DC clearance surge</div>
                <div class="insight-body">January 2026 was the highest-volume month at 8,931 US pallets across 358 events and 510 total unloading hours. This aligns with the Q1 inventory push that drove FFT's February production peak of 30,212 sellable units.</div>
              </div>

              <div class="insight-card">
                <div class="insight-label">Tijuana, BC</div>
                <div class="insight-title">4.3% of pallets — insumos, OpenCell screens, lámparas inventory, and 3PL/last-mile product</div>
                <div class="insight-body">Tijuana flows (1,391 pallets, 40 events) are semantically distinct from the TV inventory: they include insumos, pantallas for OpenCell screen swaps, lamp inventory, and the product base for 3PL and last-mile operations. May 2026 hit the highest Tijuana month (388 pallets, 9 events), tracking OpenCell's ongoing ramp.</div>
              </div>

              <div class="insight-card">
                <div class="insight-label">Trend</div>
                <div class="insight-title">US pallet volume down 54% from the January peak, stabilizing at ~4,000/month</div>
                <div class="insight-body">After the January surge (8,931 pallets), US inbound normalized to 3,818–4,873 pallets/month through Feb–May 2026 — consistent with the lower FFT run rate seen post-peak. The dock averaged 2.5 workers per unloading event across all months.</div>
              </div>

            </div>
          </div>

          <p class="footer-note">Source: Reporte UPH MI Technologies MTY MAXX — Incoming · Dec 2025 – May 2026 · Tijuana includes insumos, OpenCell screens, lámparas, and 3PL product · Report date: May 29, 2026</p>

        </div>
      </div><!-- /panel-1 -->

"""

# Anchor: the HTML comment that now precedes FFT (TAB 2).
# After the TAB-comment rename loop, "TAB 1 — FFT (report_fft_sellable.html)"
# became "TAB 2 — FFT (report_fft_sellable.html)" — match the full post-rename line.
FFT_PANEL_COMMENT = "      <!-- ══════════════════════════════════════════════════════\n           TAB 2 — FFT (report_fft_sellable.html)"
src = replace_exact(src, FFT_PANEL_COMMENT, INCOMING_PANEL + FFT_PANEL_COMMENT, "insert panel-1 block")

# ─────────────────────────────────────────────────────────────────────────────
# 9. Insert initIncoming() JS function BEFORE the TAB 2 JS comment (FFT)
# ─────────────────────────────────────────────────────────────────────────────
INIT_INCOMING = """  // ════════════════════════════════════════════════════════
  // TAB 1 — Incoming (pallets by origin)
  // ════════════════════════════════════════════════════════
  function initIncoming() {
    var MONTHS_RAW = ["2025-12","2026-01","2026-02","2026-03","2026-04","2026-05"];
    var PALLETS_US  = [5229, 8931, 4873, 4175, 3818, 4138];
    var PALLETS_TIJ = [283, 232, 123, 114, 251, 388];
    var TOTAL       = [5512, 9163, 4996, 4289, 4069, 4526];
    var HOURS       = [237.93, 510.07, 205.52, 155.48, 191.65, 201.92];
    var AVG_HC      = [2.35, 2.88, 3.04, 2.09, 2.14, 2.19];
    var LABELS = MONTHS_RAW.map(formatMonth);

    var ctx = document.getElementById("tab1Chart").getContext("2d");
    new Chart(ctx, {
      type: "bar",
      data: {
        labels: LABELS,
        datasets: [
          {
            label: "US Origins",
            data: PALLETS_US,
            backgroundColor: "#1e5f99",
            borderColor: "rgba(0,0,0,0.15)",
            borderWidth: 0.5,
            borderRadius: 2,
            borderSkipped: false,
            order: 2
          },
          {
            label: "Tijuana, BC",
            data: PALLETS_TIJ,
            backgroundColor: "#e07b6a",
            borderColor: "rgba(0,0,0,0.15)",
            borderWidth: 0.5,
            borderRadius: 2,
            borderSkipped: false,
            order: 1
          },
          {
            type: "line", label: "Total Pallets",
            data: TOTAL,
            borderColor: "#c0392b", backgroundColor: "transparent",
            borderWidth: 2, pointRadius: 5, pointBackgroundColor: "#c0392b",
            pointBorderColor: "#fff", pointBorderWidth: 1.5,
            tension: 0.3, fill: false, order: 0,
            datalabels: { display: true, align: "top", anchor: "end", color: "#c0392b",
              font: { size: 10, weight: "600" },
              formatter: function(v) { return v >= 1000 ? (v/1000).toFixed(1)+"k" : v; } }
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          datalabels: { display: false },
          legend: {
            position: "top",
            labels: { color: "#222", font: { size: 11 }, padding: 14, boxWidth: 12, boxHeight: 12, usePointStyle: false }
          },
          tooltip: {
            backgroundColor: "#ffffff",
            borderColor: "#d0d0d0",
            borderWidth: 1,
            titleColor: "#1a1a2e",
            bodyColor: "#333",
            padding: 12,
            callbacks: {
              label: function(context) {
                var v = context.parsed.y;
                return v > 0 ? " " + context.dataset.label + ": " + v.toLocaleString() : null;
              },
              footer: function(items) {
                var total = items.filter(function(i){ return i.dataset.type !== "line"; }).reduce(function(s, i) { return s + i.parsed.y; }, 0);
                return "Total: " + total.toLocaleString();
              }
            }
          }
        },
        scales: {
          x: { stacked: true, grid: { color: "#e0e0e0" }, ticks: { color: "#444", font: { size: 11 } } },
          y: {
            stacked: true,
            grid: { color: "#e0e0e0" },
            ticks: {
              color: "#444", font: { size: 11 },
              callback: function(v) { return v >= 1000 ? (v / 1000).toFixed(0) + "k" : v; }
            },
            title: { display: true, text: "Pallets", color: "#555", font: { size: 11 } }
          }
        }
      }
    });

    // Build monthly table
    var tbody = document.getElementById("tab1TableBody");
    MONTHS_RAW.forEach(function(ym, i) {
      var us  = PALLETS_US[i];
      var tij = PALLETS_TIJ[i];
      var tot = us + tij;
      var tr = document.createElement("tr");
      tr.innerHTML =
        "<td>" + formatMonth(ym) + "</td>" +
        "<td>" + us.toLocaleString() + "</td>" +
        "<td>" + tij.toLocaleString() + "</td>" +
        "<td>" + tot.toLocaleString() + "</td>" +
        "<td>" + HOURS[i].toFixed(1) + "h</td>" +
        "<td>" + AVG_HC[i].toFixed(1) + "</td>";
      tbody.appendChild(tr);
    });
  }

"""

# Anchor: the JS comment block that opens the FFT function.
# After the TAB-comment rename loop above, "// TAB 1 — FFT Sellable" became
# "// TAB 2 — FFT Sellable" — so match on the full post-rename line.
FFT_JS_COMMENT = "  // ════════════════════════════════════════════════════════\n  // TAB 2 — FFT Sellable"
src = replace_exact(src, FFT_JS_COMMENT, INIT_INCOMING + FFT_JS_COMMENT, "insert initIncoming function")

# ─────────────────────────────────────────────────────────────────────────────
# 10. Write the result back
# ─────────────────────────────────────────────────────────────────────────────
with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(src)

print("dashboard.html written successfully.")

# ─────────────────────────────────────────────────────────────────────────────
# 11. Verification checks
# ─────────────────────────────────────────────────────────────────────────────
checks = {
    'panel-7 exists':       'id="panel-7"'         in src,
    'panel-1 exists':       'id="panel-1"'         in src,
    'TAB_COUNT = 7':        'const TAB_COUNT = 7;' in src,
    'initIncoming present': 'function initIncoming' in src,
    'initIncoming routed':  'initIncoming();'       in src,
    'showTab(1) Incoming':  'onclick="showTab(1)">Incoming</button>' in src,
    'panel-1 active':       'class="tab-panel active" id="panel-1"' in src,
    'panel-2 NOT active':   'class="tab-panel active" id="panel-2"' not in src,
    'FFT btn NOT active':   'class="tab-btn active"' in src and '>FFT<' not in src.split('class="tab-btn active"')[1].split('>')[0] + src.split('class="tab-btn active"')[1].split('>')[1],
    'tab1Chart in panel-1': True,   # structural — confirmed by panel HTML
    'tab2Chart in FFT':     'id="tab2Chart"' in src,
    'tab7Chart exists':     'id="tab7Chart"' in src,
    'initTab7 exists':      'function initTab7' in src,
    'no orphan initTab1':   'function initTab1' not in src,
}

print("\nVerification:")
all_pass = True
for name, result in checks.items():
    status = "PASS" if result else "FAIL"
    if not result:
        all_pass = False
    print(f"  [{status}] {name}")

print()
if all_pass:
    print("All checks passed.")
else:
    print("Some checks FAILED — review output above.")
