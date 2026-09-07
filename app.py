"""
SmartWork AI - Customer Acquisition & Outreach Hub Server
=========================================================
Two-Panel High-Precision Operator Console on http://localhost:8000:
- Left: Territory Explorer (All 333 Districts with live count & auto-harvester)
- Right: Direct Customer Outreach CRM (WhatsApp demo pitch, 1-click copy)
"""

import os
import sys
import math
import json
import re
import urllib.parse
from collections import OrderedDict
from typing import Optional

import pandas as pd
from fastapi import FastAPI, BackgroundTasks, Query
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

app = FastAPI(title="SmartWork AI - Outreach Hub")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRAPERS_DIR = os.path.join(BASE_DIR, "scrapers")
if SCRAPERS_DIR not in sys.path:
    sys.path.append(SCRAPERS_DIR)

try:
    from unified_harvester import harvest_district_all_sources
    from directory_scrapers import auto_harvest_pending_districts
except Exception:
    try:
        from directory_scrapers import (
            harvest_district_all_sources,
            auto_harvest_pending_districts
        )
    except Exception:
        harvest_district_all_sources = None
        auto_harvest_pending_districts = None

EXCEL_MASTER = os.path.join(BASE_DIR, "FINAL_COMMON_TYPISTS.xlsx")
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "dashboard.html")
DISTRICTS_PATH = os.path.join(BASE_DIR, "hindi_belt_districts.json")

scraper_status = {
    "running": False,
    "last_run": None,
    "last_message": "Engine Ready"
}

def get_all_districts():
    if os.path.exists(DISTRICTS_PATH):
        try:
            with open(DISTRICTS_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get("districts", [])
        except Exception as e:
            print(f"Error loading districts: {e}")
    return []

ALL_DISTRICTS = get_all_districts()

def run_auto_harvest_batch_task(batch_size: int = 3):
    global scraper_status
    scraper_status["running"] = True
    scraper_status["last_message"] = f"Auto-harvesting next {batch_size} pending districts (Web + CSC + Govt)..."
    try:
        results = auto_harvest_pending_districts(batch_size=batch_size, districts_path=DISTRICTS_PATH, master_file=EXCEL_MASTER)
        dist_names = ", ".join(r["district"] for r in results) if results else "None (all districts have leads!)"
        scraper_status["last_message"] = f"Auto-harvest finished for: {dist_names}!"
    except Exception as e:
        scraper_status["last_message"] = f"Auto-harvest error: {str(e)}"
    finally:
        scraper_status["running"] = False
        scraper_status["last_run"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

def run_district_multi_source_task(district: str, state: str):
    global scraper_status
    scraper_status["running"] = True
    scraper_status["last_message"] = f"Harvesting {district} ({state}) using Web + CSC + Govt directories..."
    try:
        summary = harvest_district_all_sources(district=district, state=state)
        scraper_status["last_message"] = f"Harvested {district}: {summary['new_unique_saved']} fresh unique leads added!"
    except Exception as e:
        scraper_status["last_message"] = f"District harvest error: {str(e)}"
    finally:
        scraper_status["running"] = False
_CACHE = {
    "mtime": 0,
    "df": pd.DataFrame(),
    "counts": {},
    "total": 0,
    "total_phones": 0,
    "sources": []
}

def get_master_data():
    global _CACHE
    if not os.path.exists(EXCEL_MASTER):
        return pd.DataFrame(), {}, 0, 0, []
    mtime = os.path.getmtime(EXCEL_MASTER)
    if _CACHE["df"].empty or _CACHE["mtime"] != mtime:
        try:
            df = pd.read_excel(EXCEL_MASTER, dtype=str).fillna("")
            total = len(df)
            phone_s = df["Phone"].str.strip()
            total_phones = int((phone_s.str.len() >= 10).sum())
            if "Source" not in df.columns:
                df["Source"] = "Google Maps"
            sources = sorted(list(set(s for s in df["Source"].unique() if s.strip())))
            dist_counts = {}
            if not df.empty:
                loc_series = df["Location"].str.lower()
                dist_series = df["District"].str.lower()
                vc_dist = dist_series.value_counts().to_dict()
                vc_loc = loc_series.value_counts().to_dict()
                all_d = get_all_districts()
                for d in all_d:
                    d_name = d["district"].strip()
                    d_lower = d_name.lower()
                    base_name = d_name.split(" (")[0].strip().lower()
                    cnt = vc_dist.get(d_lower, 0)
                    if cnt == 0 and base_name != d_lower:
                        cnt = vc_dist.get(base_name, 0)
                    if cnt == 0:
                        cnt = vc_loc.get(d_lower, 0)
                    if cnt == 0 and base_name != d_lower:
                        cnt = vc_loc.get(base_name, 0)
                    dist_counts[d_name] = cnt
            _CACHE = {
                "mtime": mtime,
                "df": df,
                "counts": dist_counts,
                "total": total,
                "total_phones": total_phones,
                "sources": sources
            }
        except Exception as e:
            print(f"Error loading master excel cache: {e}")
    return _CACHE["df"], _CACHE["counts"], _CACHE["total"], _CACHE.get("total_phones", 0), _CACHE.get("sources", [])

@app.get("/", response_class=HTMLResponse)
def dashboard(
    district: Optional[str] = None,
    state: Optional[str] = None,
    phones_only: Optional[int] = 0,
    search: Optional[str] = None,
    category: Optional[str] = None,
    source: Optional[str] = None,
    page: int = 1,
    page_size: int = 25
):
    df_master, district_counts, total_master_leads, total_mobile_leads, available_sources = get_master_data()
    all_districts_list = get_all_districts()

    # Determine Active State & District
    is_mobile_view = bool(phones_only)
    is_state_view = False
    if is_mobile_view:
        active_state = "All States"
        active_district = "Verified Mobiles"
    elif state and not district:
        active_state = state
        active_district = state
        is_state_view = True
    elif district:
        active_district = district
        active_state = state or "Uttar Pradesh"
        for d in all_districts_list:
            if d["district"].lower() == active_district.lower():
                active_state = d.get("state", active_state)
                break
    else:
        active_state = "Uttar Pradesh"
        active_district = "Uttar Pradesh"
        is_state_view = True

    # Build Left Panel: State-to-District Accordions
    states_dict = OrderedDict()
    for d in all_districts_list:
        st = d.get("state", "Other").strip()
        if st not in states_dict:
            states_dict[st] = []
        states_dict[st].append(d)

    state_stats = []
    for st_name, d_list in states_dict.items():
        st_leads = sum(district_counts.get(d["district"].strip(), 0) for d in d_list)
        state_stats.append((st_name, d_list, st_leads))

    # Sort states: states with leads first, then alphabetically
    state_stats.sort(key=lambda x: (x[2], x[0]), reverse=True)

    states_list_html = ""
    for st_idx, (st_name, st_districts, st_leads) in enumerate(state_stats):
        slug = re.sub(r'[^a-zA-Z0-9]', '', st_name).lower()
        is_active_state = (st_name.lower() == active_state.lower())

        expanded_cls = "expanded" if is_active_state else ""
        open_cls = "open" if is_active_state else ""
        active_state_cls = "active-state" if (is_active_state and is_state_view) else ""
        chevron_char = "▲" if is_active_state else "▼"

        state_badge = (
            f'<span class="sc-badge sc-badge-active">{st_leads:,} Leads</span>'
            if st_leads > 0 else
            f'<span class="sc-badge sc-badge-zero">0 Leads</span>'
        )

        sorted_state_districts = sorted(
            st_districts,
            key=lambda d: (district_counts.get(d["district"].strip(), 0), d["district"].strip()),
            reverse=True
        )

        districts_in_state_html = ""
        for d in sorted_state_districts:
            dname = d["district"].strip()
            cnt = district_counts.get(dname, 0)
            is_active = "active" if (not is_state_view and dname.lower() == active_district.lower()) else ""

            badge_html = ""
            if cnt > 0:
                badge_html = f'<span class="d-badge-count">{cnt:,} Leads</span>'
            else:
                badge_html = f'<button onclick="event.preventDefault(); event.stopPropagation(); harvestDistrict(\'{dname}\', \'{st_name}\');" class="btn-harvest-mini">+ Get Leads</button>'

            districts_in_state_html += f"""
            <a href="/?district={urllib.parse.quote(dname)}&state={urllib.parse.quote(st_name)}" 
               class="d-item {is_active}" 
               data-name="{dname}" 
               data-state="{st_name}"
               onclick="selectDistrict(event, '{dname}', '{st_name}', this)">
                <div class="d-info">
                    <span class="d-info-name">{dname}</span>
                </div>
                <div class="d-action">
                    {badge_html}
                </div>
            </a>
            """

        states_list_html += f"""
        <div class="state-group" data-state="{st_name}">
            <div class="state-card {expanded_cls} {active_state_cls}" 
                 onclick="selectState(event, '{st_name}', '{slug}', this)" 
                 ondblclick="toggleStateAccordion(event, '{slug}', this)"
                 title="Click to load {st_name} leads | Double-click to minimize districts">
                <div class="sc-left">
                    <span class="sc-name">{st_name}</span>
                    <span class="sc-count">{len(st_districts)} Districts</span>
                </div>
                <div class="sc-right">
                    {state_badge}
                    <span class="sc-chevron">{chevron_char}</span>
                </div>
            </div>
            <div class="state-districts-list {open_cls}" id="state-districts-{slug}">
                {districts_in_state_html}
            </div>
        </div>
        """

    # Filter Leads for Active District or State
    df_district = pd.DataFrame()
    if not df_master.empty:
        if is_mobile_view:
            phone_s = df_master["Phone"].str.strip()
            df_district = df_master[phone_s.str.len() >= 10].copy()
        elif is_state_view:
            df_district = df_master[df_master["State"].str.contains(active_state, case=False, na=False)]
        else:
            active_base = active_district.split(" (")[0].strip().lower()
            df_district = df_master[
                (df_master["Location"].str.lower() == active_district.lower()) |
                (df_master["District"].str.lower() == active_district.lower()) |
                (df_master["District"].str.lower() == active_base) |
                (df_master["Location"].str.lower() == active_base)
            ]
            if df_district.empty and active_district == "All":
                df_district = df_master.copy()

    # Category dropdown options (No emojis)
    cat_defs = [
        ("stamp", "Stamp Vendors (e-Stamp ACC)"),
        ("csc", "CSC / Jan Seva Kendras"),
        ("typist", "Legal Typists & Katibs"),
        ("advocate", "Advocates")
    ]
    category_dropdown_options = ""
    for c_val, c_lbl in cat_defs:
        c_sel = "selected" if category and category.lower() == c_val else ""
        category_dropdown_options += f'<option value="{c_val}" {c_sel}>{c_lbl}</option>'

    # Source dropdown options (Multi-Engine & Tools)
    source_defs = [
        ("Google Maps", "Google Maps"),
        ("Justdial", "Justdial"),
        ("Trade Directory", "Trade Directories (Sulekha / IndiaMART)"),
        ("Stamp & Registration Dept", "Stamp & Registration Dept"),
        ("Bar Council / Court", "Bar Council & Courts"),
        ("CSC Cloud Locator", "CSC Cloud Locator")
    ]
    existing_src_names = {s[0].lower() for s in source_defs}
    for s_item in available_sources:
        if s_item and s_item.lower() not in existing_src_names:
            source_defs.append((s_item, s_item))

    source_dropdown_options = ""
    for s_val, s_lbl in source_defs:
        s_sel = "selected" if source and source.lower() == s_val.lower() else ""
        source_dropdown_options += f'<option value="{s_val}" {s_sel}>{s_lbl}</option>'

    # Apply search, category, and source filters
    df_filtered = df_district.copy()
    if not df_filtered.empty:
        if category:
            df_filtered = df_filtered[df_filtered["Category"].str.contains(category, case=False, na=False)]
        if source:
            df_filtered = df_filtered[df_filtered["Source"].str.lower() == source.lower()]
        if search:
            s = search.lower()
            df_filtered = df_filtered[
                df_filtered["Name"].str.lower().str.contains(s) |
                df_filtered["Phone"].str.contains(s) |
                df_filtered["Address"].str.lower().str.contains(s) |
                df_filtered["Location"].str.lower().str.contains(s) |
                df_filtered["Source"].str.lower().str.contains(s)
            ]

    filtered_count = len(df_filtered)
    total_pages = max(1, math.ceil(filtered_count / page_size))
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    df_page = df_filtered.iloc[start_idx:end_idx] if not df_filtered.empty else pd.DataFrame()
    leads = df_page.to_dict(orient="records")

    # Render Table Rows
    pitch_text = f"Namaste! We noticed your document typing and registry work in {active_district}. We built SmartWork AI (https://thesmartwork.onrender.com) which converts handwritten notes & voice memos directly into editable MS Word (.docx) files in 5 seconds. Would you like a free demo?"
    encoded_pitch = urllib.parse.quote(pitch_text)

    rows_html = ""
    for idx, lead in enumerate(leads, (page - 1) * page_size + 1):
        name = lead.get("Name", "N/A")
        phone = lead.get("Phone", "")
        cat = lead.get("Category", "Legal Typist")
        addr = lead.get("Address", "")
        lead_source = lead.get("Source", "Google Maps")

        wa_btn = ""
        copy_btn = ""
        if phone:
            wa_url = f"https://wa.me/91{phone}?text={encoded_pitch}"
            wa_btn = f'<a href="{wa_url}" target="_blank" class="btn-wa-hacker" title="Transmit AI pitch to {name}">[>_ PING_WA]</a>'
            copy_btn = f'<button onclick="navigator.clipboard.writeText(\'{phone}\'); alert(\'[COMM_KEY COPIED] {phone}\');" class="btn-cp" title="Copy Number">[CP]</button>'
        else:
            wa_btn = '<span style="color:#14532d; font-size:10.5px; font-weight:700;">[NO_COMM]</span>'

        cat_badge = '<span class="cat-badge cat-typist">[TYPIST]</span>'
        if "stamp" in cat.lower():
            cat_badge = '<span class="cat-badge cat-stamp">[STAMP_ACC]</span>'
        elif "csc" in cat.lower():
            cat_badge = '<span class="cat-badge cat-csc">[CSC_GOV]</span>'
        elif "advocate" in cat.lower():
            cat_badge = '<span class="cat-badge cat-advocate">[LEGAL_BAR]</span>'

        src_badge = '<span class="cat-badge cat-gmaps">[G-MAPS]</span>'
        if "justdial" in lead_source.lower():
            src_badge = '<span class="cat-badge cat-jd">[JUSTDIAL]</span>'
        elif "trade" in lead_source.lower():
            src_badge = '<span class="cat-badge cat-trade">[TRADE_DIR]</span>'
        elif "stamp" in lead_source.lower():
            src_badge = '<span class="cat-badge cat-stamp">[STAMP_ACC]</span>'
        elif "bar" in lead_source.lower():
            src_badge = '<span class="cat-badge cat-advocate">[LEGAL_BAR]</span>'
        elif "csc" in lead_source.lower():
            src_badge = '<span class="cat-badge cat-csc">[CSC_GOV]</span>'

        rows_html += f"""
        <tr>
            <td style="color:#22c55e; font-weight:800; font-size:11px; opacity:0.65;">#{idx:02d}</td>
            <td>
                <div style="display:flex; flex-direction:column; gap:2px;">
                    {cat_badge}
                    {src_badge}
                </div>
            </td>
            <td>
                <div style="font-weight:700; color:#f0fdf4; font-size:12.5px;">{name}</div>
            </td>
            <td>
                <div style="display:flex; align-items:center;">
                    <span class="phone-display phone-text">{phone if phone else "—"}</span>
                    {copy_btn}
                </div>
            </td>
            <td style="color:#86efac; font-size:11px; opacity:0.85;" title="{addr}">
                &gt; {addr[:65] + '...' if len(addr) > 65 else addr}
            </td>
            <td style="text-align:center;">{wa_btn}</td>
        </tr>
        """

    if not rows_html:
        rows_html = f'''
        <tr>
            <td colspan="6" style="text-align:center; padding: 48px; color: #94a3b8; font-size: 14px;">
                No customer leads found for <strong>{active_district}</strong> yet.<br><br>
                <button onclick="harvestCurrentDistrict()" class="btn-tb btn-tb-harvest" style="margin: 0 auto; display:inline-flex;">
                    [+ HARVEST LEADS FOR {active_district.upper()} NOW]
                </button>
            </td>
        </tr>
        '''

    prev_disabled = "disabled" if page <= 1 else ""
    next_disabled = "disabled" if page >= total_pages else ""
    prev_page = max(1, page - 1)
    next_page = min(total_pages, page + 1)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        template = f.read()

    replacements = {
        "{{ total_master_leads }}": f"{total_master_leads:,}",
        "{{ total_mobile_leads }}": f"{total_mobile_leads:,}",
        "{{ active_district }}": active_district,
        "{{ active_state }}": active_state,
        "{{ active_district_count }}": f"{len(df_district):,}",
        "{{ states_list_html }}": states_list_html,
        "{{ category_dropdown_options }}": category_dropdown_options,
        "{{ source_dropdown_options }}": source_dropdown_options,
        "{{ search }}": search or "",
        "{{ category }}": category or "",
        "{{ source }}": source or "",
        "{{ phones_only }}": "1" if is_mobile_view else "0",
        "{{ rows_html }}": rows_html,
        "{{ filtered_count }}": f"{filtered_count:,}",
        "{{ page }}": str(page),
        "{{ total_pages }}": str(total_pages),
        "{{ prev_page }}": str(prev_page),
        "{{ next_page }}": str(next_page),
        "{{ prev_disabled }}": prev_disabled,
        "{{ next_disabled }}": next_disabled,
    }

    for k, v in replacements.items():
        template = template.replace(k, v)

    return HTMLResponse(content=template)

@app.get("/download")
def download_excel():
    if os.path.exists(EXCEL_MASTER):
        return FileResponse(
            EXCEL_MASTER,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="FINAL_COMMON_TYPISTS.xlsx"
        )
    return JSONResponse(status_code=404, content={"error": "File not found."})

@app.post("/api/auto-harvest-batch")
def trigger_auto_harvest_batch(background_tasks: BackgroundTasks, batch_size: int = 3):
    global scraper_status
    if scraper_status["running"]:
        return {"status": "busy", "message": f"Scraper is currently running ({scraper_status['last_message']}). Please wait."}
    background_tasks.add_task(run_auto_harvest_batch_task, batch_size)
    return {"status": "started", "message": f"[AUTO-HARVEST INITIATED] Next {batch_size} pending districts being harvested with Web + CSC + Govt sources in background."}

@app.post("/api/harvest-district")
def trigger_district_harvest(background_tasks: BackgroundTasks, district: str = Query(...), state: str = Query("Uttarakhand")):
    global scraper_status
    if scraper_status["running"]:
        return {"status": "busy", "message": f"Scraper is currently running ({scraper_status['last_message']}). Please wait."}
    background_tasks.add_task(run_district_multi_source_task, district, state)
    return {"status": "started", "message": f"[HARVEST INITIATED] Full multi-source harvest started for {district} ({state}) in background."}

@app.get("/api/status")
def get_scraper_status():
    return scraper_status

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
