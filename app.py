"""
SmartWork AI - Customer Acquisition & Outreach Hub Server
=========================================================
High-Precision CRM & Operator Console:
- Left: Territory Explorer (19 States & 456 Districts with live count)
- Right: Direct Customer Outreach CRM with Mobile-First Grid, Reply Tracking & Direct WA Chat
"""

import os
import sys
import math
import json
import re
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta
from collections import OrderedDict
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Query, Body, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

app = FastAPI(title="SmartWork AI - Outreach Hub")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_MASTER = os.path.join(BASE_DIR, "FINAL_COMMON_TYPISTS.xlsx")
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "dashboard.html")
DISTRICTS_PATH = os.path.join(BASE_DIR, "hindi_belt_districts.json")

def get_all_districts():
    if os.path.exists(DISTRICTS_PATH):
        try:
            with open(DISTRICTS_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get("districts", [])
        except Exception as e:
            print(f"Error loading districts: {e}")
    return []

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
            
            # Ensure outreach columns exist
            for col, def_val in [("OutreachStatus", "QUEUED"), ("SentAt", ""), ("ReplyText", ""), ("ReplyAt", "")]:
                if col not in df.columns:
                    df[col] = def_val
                    
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

def save_master_data(df: pd.DataFrame):
    """Saves updated dataframe back to Excel safely."""
    global _CACHE
    try:
        tmp_path = os.path.join(BASE_DIR, "FINAL_COMMON_TYPISTS_tmp.xlsx")
        df.to_excel(tmp_path, index=False)
        if os.path.exists(EXCEL_MASTER):
            os.remove(EXCEL_MASTER)
        os.rename(tmp_path, EXCEL_MASTER)
        _CACHE["mtime"] = os.path.getmtime(EXCEL_MASTER)
        _CACHE["df"] = df
        return True
    except Exception as e:
        print(f"Error saving master excel: {e}")
        return False

@app.get("/", response_class=HTMLResponse)
def dashboard(
    district: Optional[str] = None,
    state: Optional[str] = None,
    phones_only: Optional[int] = 0,
    search: Optional[str] = None,
    category: Optional[str] = None,
    source: Optional[str] = None,
    status: Optional[str] = None,
    today_only: Optional[int] = 0,
    page: int = 1,
    page_size: int = 25
):
    df_master, district_counts, total_master_leads, total_mobile_leads, available_sources = get_master_data()
    all_districts_list = get_all_districts()

    # Calculate Global Outreach Metrics
    total_queued = 0
    total_sent = 0
    total_replied = 0
    total_hot_leads = 0
    sent_today = 0
    today_str = datetime.now().strftime("%Y-%m-%d")

    if not df_master.empty and "OutreachStatus" in df_master.columns:
        status_counts = df_master["OutreachStatus"].str.upper().value_counts().to_dict()
        total_queued = status_counts.get("QUEUED", 0)
        total_sent = sum(status_counts.get(s, 0) for s in ["SENT", "REPLIED", "HOT_LEAD", "CONVERTED"])
        total_replied = sum(status_counts.get(s, 0) for s in ["REPLIED", "HOT_LEAD", "CONVERTED"])
        total_hot_leads = status_counts.get("HOT_LEAD", 0) + status_counts.get("CONVERTED", 0)
        if "SentAt" in df_master.columns:
            sent_today = int(df_master["SentAt"].str.startswith(today_str).sum())

    reply_rate = f"{(total_replied / total_sent * 100):.1f}%" if total_sent > 0 else "0.0%"

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

            badge_html = (
                f'<span class="d-badge-count">{cnt:,} Leads</span>'
                if cnt > 0 else
                f'<span class="d-badge-count sc-badge-zero" style="opacity:0.4;">0</span>'
            )

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
        if today_only:
            df_district = df_master[df_master["OutreachStatus"].str.upper() == "QUEUED"].head(120).copy()
        elif is_mobile_view:
            phone_s = df_master["Phone"].str.strip()
            df_district = df_master[phone_s.str.len() >= 10].copy()
        elif is_state_view:
            df_district = df_master[df_master["State"].str.contains(active_state, case=False, na=False)].copy()
        else:
            active_base = active_district.split(" (")[0].strip().lower()
            df_district = df_master[
                (df_master["Location"].str.lower() == active_district.lower()) |
                (df_master["District"].str.lower() == active_district.lower()) |
                (df_master["District"].str.lower() == active_base) |
                (df_master["Location"].str.lower() == active_base)
            ].copy()
            if df_district.empty and active_district == "All":
                df_district = df_master.copy()

    # Category dropdown options
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

    # Status dropdown options
    status_defs = [
        ("QUEUED", f"Queued / Not Contacted ({total_queued:,})"),
        ("SENT", f"Sent / Awaiting Reply ({total_sent:,})"),
        ("REPLIED", f"Replied / Conversing ({total_replied:,})"),
        ("HOT_LEAD", f"Hot Leads / Demo Ready ({total_hot_leads:,})")
    ]
    status_dropdown_options = ""
    for st_val, st_lbl in status_defs:
        st_sel = "selected" if status and status.upper() == st_val else ""
        status_dropdown_options += f'<option value="{st_val}" {st_sel}>{st_lbl}</option>'

    # Apply search, category, status, and text filters
    df_filtered = df_district.copy()
    if not df_filtered.empty:
        if category:
            df_filtered = df_filtered[df_filtered["Category"].str.contains(category, case=False, na=False)]
        if status:
            df_filtered = df_filtered[df_filtered["OutreachStatus"].str.upper() == status.upper()]
        if search:
            s = search.lower()
            df_filtered = df_filtered[
                df_filtered["Name"].str.lower().str.contains(s) |
                df_filtered["Phone"].str.contains(s) |
                df_filtered["Address"].str.lower().str.contains(s) |
                df_filtered["Location"].str.lower().str.contains(s) |
                df_filtered["District"].str.lower().str.contains(s) |
                df_filtered["ReplyText"].str.lower().str.contains(s)
            ]

        # Always pin Test Lead 9084714807 at the top of results on Page 1
    df_test = df_master[df_master["Phone"].str.endswith("9084714807")]
    if not df_test.empty and page == 1:
        df_filtered_clean = df_filtered[~df_filtered["Phone"].str.endswith("9084714807")]
        df_filtered = pd.concat([df_test, df_filtered_clean], ignore_index=True)

    filtered_count = len(df_filtered)
    total_pages = max(1, math.ceil(filtered_count / page_size))
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    df_page = df_filtered.iloc[start_idx:end_idx] if not df_filtered.empty else pd.DataFrame()
    leads = df_page.to_dict(orient="records")

    # Render Table Rows: Mobile Number 1st Column, Name & Cat 2nd, Reply Status & Text 3rd, Direct WA & Reply Buttons 4th
    rows_html = ""
    for idx, lead in enumerate(leads, (page - 1) * page_size + 1):
        name = lead.get("Name", "Customer Lead")
        clean_name = re.sub(r"\[.*?\]", "", name).strip() or name
        if "/" in clean_name:
            clean_name = clean_name.split("/")[0].strip()
        phone = lead.get("Phone", "")
        cat = lead.get("Category", "Legal Typist")
        addr = lead.get("Address", "")
        lead_dist = lead.get("District", active_district)
        lead_status = str(lead.get("OutreachStatus", "QUEUED")).upper().strip()
        reply_msg = str(lead.get("ReplyText", "")).strip()
        sent_time = str(lead.get("SentAt", "")).strip()
        reply_time = str(lead.get("ReplyAt", "")).strip()

        is_test_lead = (phone == "9084714807" or "TESTING LAB" in name.upper())
        row_style = 'style="background: rgba(245, 158, 11, 0.12); border-left: 3px solid #f59e0b;"' if is_test_lead else ""
        id_badge = '<span style="background:#f59e0b; color:#000; padding:2px 5px; border-radius:3px; font-weight:900; font-size:10px;">★ TEST</span>' if is_test_lead else f'#{idx:02d}'

        # Category Badge
        cat_badge = '<span class="cat-badge cat-typist">[TYPIST]</span>'
        if "stamp" in cat.lower():
            cat_badge = '<span class="cat-badge cat-stamp">[STAMP_VENDOR]</span>'
        elif "csc" in cat.lower() or "jan seva" in cat.lower():
            cat_badge = '<span class="cat-badge cat-csc">[CSC_KENDRA]</span>'
        elif "advocate" in cat.lower():
            cat_badge = '<span class="cat-badge cat-advocate">[ADVOCATE]</span>'

        # Status Badge
        if lead_status == "REPLIED":
            status_badge = '<span class="status-badge status-replied">[REPLIED]</span>'
        elif lead_status in ["HOT_LEAD", "CONVERTED"]:
            status_badge = '<span class="status-badge status-hot">[HOT_LEAD]</span>'
        elif lead_status == "SENT":
            status_badge = f'<span class="status-badge status-sent" title="Sent at: {sent_time}">[SENT]</span>'
        else:
            status_badge = '<span class="status-badge status-queued">[QUEUED]</span>'

        # Reply Content Column
        reply_content_html = ""
        if reply_msg:
            reply_content_html = f'''
            <div class="reply-bubble">
                <div class="reply-bubble-header">&#128172; REPLIED ({reply_time.split(" ")[0] if reply_time else "Recent"}):</div>
                <div class="reply-bubble-text">"{reply_msg}"</div>
            </div>
            '''
        elif lead_status == "SENT":
            reply_content_html = f'''
            <div class="reply-pending">
                <span class="clock-icon">&#9203;</span> Message Transmitted ({sent_time}) &bull; Awaiting Reply
            </div>
            '''
        else:
            reply_content_html = '<span class="reply-not-sent">&mdash; Not contacted yet</span>'

        # Official Matching & Pitching Script for all 4 Categories
        clean_greeting_name = clean_name if is_test_lead or len(name) < 25 else "bhai"
        pitch_text = f'''Hello {clean_greeting_name}, हमने तुम्हारे काम का एक AI tool बनाया है जो दो काम करता है:

1. Handwritten paper की photo लेकर उसे editable MS Word document में बदल देता है। (matlab koi customer hath se likha hua paper leke aaya, aap website se unka photo khinchna or document ms word me teyar ho jayega, kuch v likhne ki jarurat nhi)
2. Voice से application/document तैयार कर देता है। (matlab koi customer aata hai kuch letter likhwane ko to unko website ke mic me bolne ko kahe, website khud pura letter teyar kr degi, agr wo bolte time glti v krta hai fir v ai usko sahi kr dega)
यानी typing का काफी काम bahut aasan हो सकता है।
tool free hai jb chahe website se kaam krwa skte ho. Abhi try v kr skte ho link ye raha: https://thesmartwork.onrender.com
ya google me ye search kro: thesmartwork.onrender.com'''
        encoded_pitch = urllib.parse.quote(pitch_text)

        # Action Buttons: Direct Chat on WhatsApp + Quick Reply Logger
        wa_btn = ""
        copy_btn = ""
        reply_action_btn = ""

        if phone:
            wa_url = f"https://wa.me/91{phone}?text={encoded_pitch}"
            
            if is_test_lead:
                wa_btn = f'''
                <div style="display:flex; flex-direction:column; gap:3px; width:100%;">
                    <button onclick="sendCloudMessageDirect('{phone}', `{pitch_text}`)" class="btn-wa-direct" style="background:rgba(245,158,11,0.25); border-color:#f59e0b; color:#fbbf24; font-size:9.5px;" title="Send Direct Cloud Message via Render Evolution Gateway">
                        [ ⚡ SEND CLOUD MSG ]
                    </button>
                    <a href="{wa_url}" target="_blank" onclick="markLeadSent('{phone}')" class="btn-wa-direct" style="font-size:9px;" title="Open WhatsApp Web">
                        [ &gt;_ CHAT ON WA ]
                    </a>
                </div>
                '''
                reply_action_btn = f'''
                <div style="display:flex; gap:3px; margin-top:2px;">
                    <button onclick="simulateInboundReply('{phone}')" class="btn-log-reply" style="font-size:8.5px;" title="Simulate Inbound Customer Reply">
                        [ 💬 SIMULATE ]
                    </button>
                    <button onclick="resetTestLead('{phone}')" class="btn-log-reply" style="border-color:#ef4444; color:#ef4444; font-size:8.5px;" title="Reset Status">
                        [ ⟳ RESET ]
                    </button>
                </div>
                '''
            else:
                wa_btn = f'''
                <a href="{wa_url}" target="_blank" onclick="markLeadSent('{phone}')" class="btn-wa-direct" title="Open WhatsApp Chat & Send AI Pitch">
                    [ &gt;_ CHAT ON WA ]
                </a>
                '''
                reply_action_btn = f'''
                <button onclick="promptLogReply('{phone}', '{name}')" class="btn-log-reply" title="Log/Update customer response">
                    [ + LOG REPLY ]
                </button>
                '''
            copy_btn = f'<button onclick="navigator.clipboard.writeText(\'{phone}\'); alert(\'[COMM_KEY COPIED] {phone}\');" class="btn-cp" title="Copy Number">[CP]</button>'
        else:
            wa_btn = '<span style="color:#14532d; font-size:10.5px; font-weight:700;">[NO_PHONE]</span>'

        rows_html += f"""
        <tr {row_style}>
            <td style="color:#22c55e; font-weight:800; font-size:11px; text-align:center;">{id_badge}</td>
            <td>
                <div style="display:flex; align-items:center; gap:6px;">
                    <span class="phone-display phone-text" style="font-size:12.5px; font-weight:800; color:{'#fbbf24' if is_test_lead else '#39ff14'}; letter-spacing:0.5px;">{phone if phone else "—"}</span>
                    {copy_btn}
                </div>
                {'<div style="font-size:9px; color:#f59e0b; font-weight:800; margin-top:1px;">[🧪 YOUR TEST NUMBER]</div>' if is_test_lead else ''}
            </td>
            <td>
                <div style="font-weight:700; color:{'#fbbf24' if is_test_lead else '#f0fdf4'}; font-size:12.5px; margin-bottom:3px;">{name}</div>
                {cat_badge}
            </td>
            <td style="text-align:center;">
                {status_badge}
            </td>
            <td>
                {reply_content_html}
            </td>
            <td style="color:#86efac; font-size:11px; opacity:0.85;" title="{addr}">
                <div style="font-weight:700; color:#a7f3d0;">{lead_dist}</div>
                <div style="opacity:0.75; font-size:10.5px;">&gt; {addr[:45] + '...' if len(addr) > 45 else addr}</div>
            </td>
            <td style="text-align:center;">
                <div style="display:flex; flex-direction:column; gap:4px; align-items:center;">
                    {wa_btn}
                    {reply_action_btn}
                </div>
            </td>
        </tr>
        """

    if not rows_html:
        rows_html = f'''
        <tr>
            <td colspan="7" style="text-align:center; padding: 48px; color: #86efac; font-size: 13px; opacity: 0.8;">
                [RECON_NOTICE] No customer records found matching current filters for <strong>{active_district}</strong>.
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
        "{{ total_queued }}": f"{total_queued:,}",
        "{{ total_sent }}": f"{total_sent:,}",
        "{{ total_replied }}": f"{total_replied:,}",
        "{{ total_hot_leads }}": f"{total_hot_leads:,}",
        "{{ sent_today }}": f"{sent_today}",
        "{{ reply_rate }}": reply_rate,
        "{{ active_district }}": active_district,
        "{{ active_state }}": active_state,
        "{{ active_district_count }}": f"{len(df_district):,}",
        "{{ states_list_html }}": states_list_html,
        "{{ category_dropdown_options }}": category_dropdown_options,
        "{{ status_dropdown_options }}": status_dropdown_options,
        "{{ search }}": search or "",
        "{{ category }}": category or "",
        "{{ status }}": status or "",
        "{{ phones_only }}": "1" if is_mobile_view else "0",
        "{{ today_only }}": "1" if today_only else "0",
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

# ============================================================================
# OUTREACH APIS & DATA UPDATE ENDPOINTS
# ============================================================================

@app.post("/api/outreach/mark-status")
def mark_outreach_status(
    phone: str = Query(...),
    status: str = Query("SENT"),
    reply_text: Optional[str] = Query(None)
):
    """Updates lead outreach status (SENT, REPLIED, HOT_LEAD) with timestamps."""
    df_master, _, _, _, _ = get_master_data()
    if df_master.empty:
        return JSONResponse(status_code=404, content={"error": "Database empty"})

    clean_ph = re.sub(r"\D", "", str(phone))
    if len(clean_ph) >= 10:
        clean_ph = clean_ph[-10:]

    match_idx = df_master[df_master["Phone"].str.endswith(clean_ph)].index
    if len(match_idx) == 0:
        return JSONResponse(status_code=404, content={"error": f"Lead {phone} not found"})

    idx = match_idx[0]
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df_master.at[idx, "OutreachStatus"] = status.upper()
    if status.upper() == "SENT" and not df_master.at[idx, "SentAt"]:
        df_master.at[idx, "SentAt"] = now_iso
    elif status.upper() in ["REPLIED", "HOT_LEAD", "CONVERTED"]:
        df_master.at[idx, "ReplyAt"] = now_iso
        if reply_text:
            df_master.at[idx, "ReplyText"] = reply_text

    save_master_data(df_master)
    return {
        "status": "success",
        "phone": clean_ph,
        "new_status": status.upper(),
        "timestamp": now_iso
    }

@app.post("/api/outreach/add-lead")
def add_new_lead(
    name: str = Query(...),
    phone: str = Query(...),
    district: str = Query(...),
    state: str = Query(...),
    category: str = Query("Legal Typist"),
    address: Optional[str] = Query("")
):
    """Adds a new verified lead directly to the database."""
    clean_ph = re.sub(r"\D", "", str(phone))
    if len(clean_ph) == 12 and clean_ph.startswith("91"):
        clean_ph = clean_ph[2:]
    if len(clean_ph) != 10 or clean_ph[0] not in "6789":
        return JSONResponse(status_code=400, content={"error": "Invalid 10-digit mobile number."})

    df_master, _, _, _, _ = get_master_data()
    if not df_master.empty and df_master["Phone"].str.endswith(clean_ph).any():
        return JSONResponse(status_code=409, content={"error": "Lead with this mobile number already exists."})

    new_row = {
        "Name": name.strip(),
        "Phone": clean_ph,
        "Location": district.strip(),
        "District": district.strip(),
        "State": state.strip(),
        "Address": address or f"{district}, {state}",
        "Category": category.strip(),
        "Source": "Manual Lead",
        "Rating": "5.0",
        "ReviewsCount": 0,
        "OutreachStatus": "QUEUED",
        "SentAt": "",
        "ReplyText": "",
        "ReplyAt": "",
        "ScrapedAt": datetime.now(timezone.utc).isoformat()
    }

    df_new = pd.concat([pd.DataFrame([new_row]), df_master], ignore_index=True)
    save_master_data(df_new)
    return {"status": "success", "message": f"Lead {name} ({clean_ph}) added successfully!"}


# ============================================================================
# GEMINI 3.6 FLASH AI SALES BRAIN & AUTO-CHATTING ENGINE
# ============================================================================

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", os.environ.get("VITE_GEMINI_API_KEY", ""))
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.6-flash")

CATEGORY_KNOWLEDGE = {
    "typist": {
        "title": "Legal / Court Typist (कोर्ट टाइपिस्ट)",
        "pain_point": "वकीलों की खराब हैंडराइटिंग पढ़ना और 20-30 पेज कृतिदेव/मंगल फॉन्ट में टाइप करने में घंटों लगना।",
        "pitch": "SmartWork AI (https://thesmartwork.onrender.com) हाथ से लिखे किसी भी केस/ड्राफ्ट की फोटो या ऑडियो से 5 सेकंड में 100% सही KrutiDev/Mangal MS Word फाइल तैयार कर देता है।",
        "benefits": "दिन भर में 3 गुना ज्यादा टाइपिंग काम करें और 3-4 घंटे बचाएं।"
    },
    "stamp": {
        "title": "Stamp Vendor / Deed Writer (स्टाम्प वेंडर / वसीका नवीस)",
        "pain_point": "किरायानामा (Rent Agreement), शपथ पत्र (Affidavit), बैनामा (Sale Deed) और मुख्तारनामा का फॉर्मेट बार-बार तैयार करना।",
        "pitch": "सभी रेडीमेड लीगल फॉर्मेट्स उपलब्ध हैं। बस पार्टी की डिटेल बोलिए या रफ नोट अपलोड कीजिए, तुरंत स्टाम्प पेपर के हिसाब से रेडी ड्राफ्ट मिल जाता है।",
        "benefits": "कस्टमर को बिना इंतज़ार कराए 2 मिनट में प्रिंटेड एग्रीमेंट दें।"
    },
    "csc": {
        "title": "CSC Kendra / Jan Seva Kendra / Cyber Cafe (सीएससी केंद्र संचालक)",
        "pain_point": "कस्टमर के आवेदन पत्र, आय/जाति/निवास के प्रार्थना पत्र, बायोडाटा/रिज्यूमे बनाना।",
        "pitch": "कस्टमर के व्हाट्सएप फोटो या हाथ से लिखे कागज को तुरंत साफ़-सुथरे सरकारी प्रार्थना पत्र या रिज्यूमे में बदलें।",
        "benefits": "प्रति कस्टमर 15-20 मिनट की बचत और दुकान पर भीड़ का तुरंत समाधान।"
    },
    "advocate": {
        "title": "Advocate / Legal Practitioner (अधिवक्ता / वकील साहब)",
        "pain_point": "केस डायरी, जमानत याचिका (Bail App), लीगल नोटिस और रिट पिटीशन की फास्ट ड्राफ्टिंग।",
        "pitch": "केस के मुख्य तथ्य बस बोलकर या डायरी की फोटो डालकर उचित सेक्शन्स (IPC/BNS, CrPC/BNSS) के साथ स्टैंडर्ड लीगल ड्राफ्ट बनाएं।",
        "benefits": "जूनियर या टाइपिस्ट पर निर्भरता खत्म, कोर्ट के लिए तुरंत तैयार ड्राफ्ट।"
    }
}

def get_category_key(cat_str: str) -> str:
    cat_lower = str(cat_str).lower()
    if "stamp" in cat_lower or "vendor" in cat_lower or "deed" in cat_lower:
        return "stamp"
    elif "csc" in cat_lower or "jan seva" in cat_lower or "cyber" in cat_lower or "kendra" in cat_lower:
        return "csc"
    elif "advocate" in cat_lower or "lawyer" in cat_lower or "vakeel" in cat_lower:
        return "advocate"
    return "typist"

def generate_gemini_sales_reply(lead: dict, incoming_msg: str) -> tuple[str, bool]:
    """
    Generates tailored, conversational Hindi/Hinglish reply using Gemini 3.6 Flash
    with deep category-specific knowledge and intelligent objection handling.
    """
    name = lead.get("Name", "सर")
    cat = lead.get("Category", "Legal Typist")
    dist = lead.get("District", "")
    state = lead.get("State", "")
    
    cat_key = get_category_key(cat)
    cat_info = CATEGORY_KNOWLEDGE[cat_key]
    
    # Clean name if it has brackets or simulator text
    clean_name = re.sub(r"\[.*?\]", "", name).strip() or "सर"
    if "/" in clean_name:
        clean_name = clean_name.split("/")[0].strip()

    lower_msg = incoming_msg.lower().strip()

    # Detect if customer is a HOT LEAD
    hot_keywords = ["demo", "link", "trial", "price", "kharidna", "call", "bhejo", "de do", "batao", "free", "try", "use", "karein", "kaise", "haan", "ha", "yes", "ok", "intrested", "interested", "accha", "sahi", "krutidev", "mangal", "format"]
    is_hot = any(k in lower_msg for k in hot_keywords)

    system_prompt = f"""
आप SmartWork AI (https://thesmartwork.onrender.com) के आधिकारिक AI सेल्स असिस्टेंट हैं।
कस्टमर से व्हाट्सएप पर बहुत ही स्वाभाविक, दोस्ताना और मददगार हिंदी/हिंग्लिश में बात करें (2-3 छोटी पंक्तियाँ)।

[हमारा मुख्य समाधान और 2 फीचर्स]
1. Handwritten Paper to Word: हाथ से लिखे किसी भी कागज की फोटो खींचकर 5 सेकंड में MS Word (.docx) में बदल देता है (बिना टाइप किए)।
2. Voice to Letter / Document: माइक में बोलने पर पूरा लेटर/आवेदन तैयार हो जाता है, और बोलने में कोई गलती हो तो AI खुद सही कर देता है।
3. टूल बिल्कुल मुफ़्त है और कभी भी https://thesmartwork.onrender.com पर जाकर यूज़ किया जा सकता है।

[ग्राहक का नाम]: {clean_name} ({dist}, {state})
[कस्टमर का मैसेज]: "{incoming_msg}"

[नियम]
- केवल 2 से 3 पंक्तियाँ में सटीक जवाब दें।
- अगर कीमत पूछे: बताएं कि टूल बिल्कुल फ्री है, जब चाहें इस्तेमाल कर सकते हैं।
- अगर डेमो मांगे: लिंक https://thesmartwork.onrender.com दें या बोलें कि Google में thesmartwork.onrender.com सर्च करें।
- हर जवाब के अंत में एक छोटा सवाल पूछें: "क्या आप अभी 1 पेज का मुफ़्त टेस्ट करके देखना चाहेंगे?"
"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    payload = {
        "contents": [{"parts": [{"text": system_prompt}]}],
        "generationConfig": {"temperature": 0.35, "maxOutputTokens": 250}
    }
    
    try:
        req_data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=req_data,
            headers={
                "Content-Type": "application/json",
                "x-goog-api-key": GEMINI_API_KEY
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=10) as res:
            res_json = json.loads(res.read().decode("utf-8"))
            ai_text = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
            return ai_text, is_hot
    except Exception as e:
        # Category-Specific Intelligent Knowledge Engine (Fallback & Instant Speed)
        if "price" in lower_msg or "kitna" in lower_msg or "charge" in lower_msg or "paisa" in lower_msg or "cost" in lower_msg:
            return f"नमस्ते {clean_name} जी! अभी हम सभी कोर्ट टाइपिस्ट और सीएससी भाइयों को 7 दिन का बिल्कुल 100% मुफ़्त ट्रायल दे रहे हैं, कोई चार्ज नहीं है। आप https://thesmartwork.onrender.com पर जाकर खुद टेस्ट करके देख सकते हैं। क्या आप 1 पेज चेक करना चाहेंगे?", True
            
        elif "link" in lower_msg or "demo" in lower_msg or "trial" in lower_msg or "bhejo" in lower_msg or "de do" in lower_msg:
            return f"नमस्ते {clean_name} जी! यह रहा SmartWork AI का डायरेक्ट लिंक: https://thesmartwork.onrender.com - यहाँ आप हाथ से लिखे किसी भी कागज की फोटो डालकर तुरंत MS Word (.docx) फ़ाइल बना सकते हैं। टेस्ट करके जरूर बताएं सर!", True
            
        elif "krutidev" in lower_msg or "mangal" in lower_msg or "font" in lower_msg:
            return f"नमस्ते {clean_name} जी! जी बिल्कुल, SmartWork AI कृतिदेव (KrutiDev) और मंगल (Mangal) दोनों फॉन्ट में 100% सही MS Word फाइल देता है। आप https://thesmartwork.onrender.com पर 2 मिनट में टेस्ट कर सकते हैं।", True

        elif cat_key == "typist":
            return f"नमस्ते {clean_name} जी! स्मार्टवर्क एआई (https://thesmartwork.onrender.com) पर आप वकीलों के हाथ से लिखे कठिन नोट्स की फोटो डालकर 5 सेकंड में कृतिदेव/मंगल MS Word फाइल बना सकते हैं। क्या मैं आपको इसका 2 मिनट का फ्री डेमो दिखाऊँ?", is_hot

        elif cat_key == "stamp":
            return f"नमस्ते {clean_name} जी! स्मार्टवर्क एआई में किरायानामा, शपथ पत्र और बैनामा के सभी रेडीमेड लीगल फॉर्मेट्स उपलब्ध हैं। आप https://thesmartwork.onrender.com पर 1 ड्राफ्ट बनाकर देख सकते हैं। क्या आप टेस्ट करना चाहेंगे?", is_hot

        elif cat_key == "csc":
            return f"नमस्ते {clean_name} जी! ग्राहक के व्हाट्सएप फोटो या आवेदन पत्र को 5 सेकंड में सरकारी फॉर्मेट की Word फ़ाइल में बदलें। डेमो लिंक: https://thesmartwork.onrender.com - क्या आप 1 आवेदन पत्र टेस्ट करना चाहेंगे?", is_hot

        elif cat_key == "advocate":
            return f"सादर प्रणाम वकील साहब! आप केस के मुख्य तथ्य बोलकर या नोट्स की फोटो से धारा (Sections) सहित जमानत याचिका व लीगल नोटिस तुरंत ड्राफ्ट कर सकते हैं (https://thesmartwork.onrender.com)। क्या आप 1 ड्राफ्ट देखना चाहेंगे?", is_hot

        return f"नमस्ते {clean_name} जी! स्मार्टवर्क एआई (https://thesmartwork.onrender.com) पर आप हाथ से लिखे कागज की फोटो या बोलकर 5 सेकंड में MS Word (.docx) फाइल बना सकते हैं। क्या आप मुफ़्त में 1 पेज टेस्ट करना चाहेंगे?", is_hot



# ============================================================================
# AUTOMATED 2-MINUTE OUTREACH BROADCAST ENGINE & SCHEDULER
# ============================================================================

AUTO_REPLY_ENABLED = False  # Set to False by default as requested

BROADCAST_STATE = {
    "active": True,           # Broadcast starts active
    "interval_sec": 120,      # 1 lead every 2 minutes
    "last_sent_phone": "",
    "last_sent_name": "",
    "last_sent_time": 0,
    "last_sent_iso": "",
    "sent_today": 0,
    "status": "RUNNING",
    "quiet_hours": "10:00 PM - 07:00 AM IST"
}

def is_quiet_hours() -> bool:
    """Returns True if current time is between 10:00 PM (22:00) and 07:00 AM IST."""
    ist_zone = timezone(timedelta(hours=5, minutes=30))
    now_ist = datetime.now(ist_zone)
    return now_ist.hour >= 22 or now_ist.hour < 7

def outreach_broadcast_worker():
    """
    STRICT REAL-TIME BROADCAST WORKER:
    1. Pauses completely if WhatsApp is not linked via QR code.
    2. Verifies if number is registered on WhatsApp via socket.
    3. If number is NOT on WhatsApp: marks as 'NO_WHATSAPP' and IMMEDIATELY switches to next lead (0 delay).
    4. ONLY when message is CONFIRMED SENT on real WhatsApp: marks as 'SENT' and starts 2-minute countdown.
    5. Honors quiet hours (10:00 PM - 07:00 AM IST).
    """
    import time
    while True:
        try:
            if BROADCAST_STATE["active"]:
                if is_quiet_hours():
                    BROADCAST_STATE["status"] = "QUIET_HOURS (10 PM - 7 AM IST)"
                    time.sleep(30)
                    continue

                # 1. Check if WhatsApp Gateway is Connected
                wa_st = get_whatsapp_status()
                if wa_st.get("status") != "CONNECTED":
                    BROADCAST_STATE["status"] = "WHATSAPP_NOT_LINKED (Please Scan QR Code)"
                    time.sleep(10)
                    continue

                now_ts = time.time()
                elapsed = now_ts - BROADCAST_STATE["last_sent_time"]
                
                if elapsed >= BROADCAST_STATE["interval_sec"]:
                    # Find next queued lead
                    df_master, _, _, _, _ = get_master_data()
                    if not df_master.empty:
                        queued_mask = (df_master["OutreachStatus"].str.upper() == "QUEUED") & (df_master["Phone"].str.len() >= 10)
                        queued_indices = df_master[queued_mask].index
                        
                        if len(queued_indices) > 0:
                            idx = queued_indices[0]
                            lead = df_master.iloc[idx].to_dict()
                            phone = re.sub(r"\D", "", str(lead.get("Phone", "")))
                            if len(phone) >= 10:
                                phone = phone[-10:]
                            
                            name = lead.get("Name", "bhai")
                            clean_name = re.sub(r"\[.*?\]", "", name).strip() or "bhai"
                            if "/" in clean_name:
                                clean_name = clean_name.split("/")[0].strip()
                            if len(clean_name) > 25:
                                clean_name = "bhai"

                            pitch_text = f"""Hello {clean_name}, हमने तुम्हारे काम का एक AI tool बनाया है जो दो काम करता है:

1. Handwritten paper की photo लेकर उसे editable MS Word document में बदल देता है। (matlab koi customer hath se likha hua paper leke aaya, aap website se unka photo khinchna or document ms word me teyar ho jayega, kuch v likhne ki jarurat nhi)
2. Voice से application/document तैयार कर देता है। (matlab koi customer aata hai kuch letter likhwane ko to unko website ke mic me bolne ko kahe, website khud pura letter teyar kr degi, agr wo bolte time glti v krta hai fir v ai usko sahi kr dega)
यानी typing का काफी काम bahut aasan हो सकता है।
tool free hai jb chahe website se kaam krwa skte ho. Abhi try v kr skte ho link ye raha: https://thesmartwork.onrender.com
ya google me ye search kro: thesmartwork.onrender.com"""

                            payload = {"number": f"91{phone}", "text": pitch_text}
                            gw_url = f"{EVOLUTION_API_URL}/message/sendText/smartwork_outreach"
                            
                            try:
                                req_data = json.dumps(payload).encode("utf-8")
                                req = urllib.request.Request(gw_url, data=req_data, headers={"Content-Type": "application/json", "User-Agent": "SmartWork-Broadcast"}, method="POST")
                                with urllib.request.urlopen(req, timeout=15) as res:
                                    res_data = json.loads(res.read().decode("utf-8"))
                                    send_status = res_data.get("status")

                                    if send_status == "not_on_whatsapp":
                                        # Number has no WhatsApp -> mark NO_WHATSAPP and switch to next immediately (0 sec delay)
                                        print(f"[NO_WA_SKIP] {clean_name} ({phone}) is NOT on WhatsApp. Skipping immediately to next lead.")
                                        df_master.at[idx, "OutreachStatus"] = "NO_WHATSAPP"
                                        df_master.at[idx, "SentAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        save_master_data(df_master)
                                        BROADCAST_STATE["status"] = f"SKIPPED_NO_WA ({phone})"
                                        # Do not reset last_sent_time so next loop immediately tries the next lead!
                                        time.sleep(2)
                                        continue

                                    elif send_status == "sent":
                                        # Real WhatsApp Message Sent!
                                        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        df_master.at[idx, "OutreachStatus"] = "SENT"
                                        df_master.at[idx, "SentAt"] = now_iso
                                        save_master_data(df_master)

                                        BROADCAST_STATE["last_sent_time"] = time.time()
                                        BROADCAST_STATE["last_sent_phone"] = phone
                                        BROADCAST_STATE["last_sent_name"] = clean_name
                                        BROADCAST_STATE["last_sent_iso"] = now_iso
                                        BROADCAST_STATE["sent_today"] += 1
                                        BROADCAST_STATE["status"] = f"SENT TO {clean_name} ({phone})"
                                        print(f"[REAL_WHATSAPP_DISPATCH] Confirmed message sent to {clean_name} ({phone})!")
                                    else:
                                        print(f"[SEND_UNCONFIRMED] Gateway returned status: {send_status}. Keeping lead as QUEUED.")
                                        time.sleep(10)
                            except urllib.error.HTTPError as http_err:
                                err_body = http_err.read().decode("utf-8")
                                print(f"[GATEWAY_HTTP_ERR] Code {http_err.code}: {err_body}")
                                BROADCAST_STATE["status"] = "WHATSAPP_NOT_LINKED (Scan QR First)" if http_err.code == 503 else f"GATEWAY_ERR_{http_err.code}"
                                time.sleep(10)
                            except Exception as err:
                                print(f"[SCHEDULER_CONN_WARN] Could not reach WhatsApp gateway: {err}")
                                BROADCAST_STATE["status"] = "GATEWAY_UNREACHABLE"
                                time.sleep(10)
                        else:
                            BROADCAST_STATE["status"] = "ALL_LEADS_CONTACTED"
                            time.sleep(30)
                else:
                    remaining = int(BROADCAST_STATE["interval_sec"] - elapsed)
                    BROADCAST_STATE["status"] = f"NEXT SEND IN {remaining}s"
            else:
                BROADCAST_STATE["status"] = "PAUSED"
        except Exception as e:
            print(f"[SCHEDULER_LOOP_ERROR] {e}")
        
        time.sleep(5)

# Start Background Scheduler Thread on server startup
import threading
threading.Thread(target=outreach_broadcast_worker, daemon=True).start()

# Scheduler Control APIs
@app.get("/api/outreach/broadcast/status")
def get_broadcast_status():
    ist_zone = timezone(timedelta(hours=5, minutes=30))
    now_ist = datetime.now(ist_zone).strftime("%I:%M:%S %p IST")
    return {
        **BROADCAST_STATE,
        "is_quiet_hours": is_quiet_hours(),
        "auto_reply_enabled": AUTO_REPLY_ENABLED,
        "current_time_ist": now_ist
    }

@app.post("/api/outreach/broadcast/start")
def start_broadcast():
    BROADCAST_STATE["active"] = True
    BROADCAST_STATE["status"] = "RUNNING"
    return {"status": "success", "message": "Broadcast engine activated with 2-minute pacing."}

@app.post("/api/outreach/broadcast/pause")
def pause_broadcast():
    BROADCAST_STATE["active"] = False
    BROADCAST_STATE["status"] = "PAUSED"
    return {"status": "success", "message": "Broadcast engine paused."}

@app.post("/api/outreach/toggle-auto-reply")
def toggle_auto_reply(enable: Optional[bool] = Query(None)):
    global AUTO_REPLY_ENABLED
    if enable is not None:
        AUTO_REPLY_ENABLED = enable
    else:
        AUTO_REPLY_ENABLED = not AUTO_REPLY_ENABLED
    return {"status": "success", "auto_reply_enabled": AUTO_REPLY_ENABLED}


@app.post("/api/outreach/webhook")
async def incoming_reply_webhook(request: Request):
    """
    STRICT FILTERING WEBHOOK:
    1. ONLY processes incoming messages from verified SmartWork leads in FINAL_COMMON_TYPISTS.xlsx (~3,586 numbers).
    2. IGNORES all personal chats, friends, family, or unknown numbers to protect operator privacy.
    3. Triggers Gemini 3.6 Flash AI to converse naturally in Hindi based on customer category.
    4. Automatically sends AI response back to customer on WhatsApp via Render Evolution Gateway.
    5. Escalates interested leads to HOT_LEAD with real-time timestamps.
    """
    try:
        body = await request.json()
        
        # Support Evolution API v2 (messages.upsert) and flat payload formats
        sender = ""
        message = ""
        
        if "data" in body and isinstance(body["data"], dict):
            key_data = body["data"].get("key", {})
            if key_data.get("fromMe", False):
                return {"status": "ignored_self_sent"}
            sender = str(key_data.get("remoteJid", ""))
            msg_obj = body["data"].get("message", {})
            message = msg_obj.get("conversation") or msg_obj.get("extendedTextMessage", {}).get("text") or msg_obj.get("imageMessage", {}).get("caption") or ""
        else:
            sender = str(body.get("sender", body.get("phone", body.get("from", ""))))
            message = str(body.get("message", body.get("text", body.get("body", ""))))

        if not message.strip():
            return {"status": "ignored_empty_message"}

        clean_ph = re.sub(r"\D", "", sender)
        if len(clean_ph) >= 10:
            clean_ph = clean_ph[-10:]

        df_master, _, _, _, _ = get_master_data()
        if df_master.empty:
            return JSONResponse(status_code=503, content={"error": "Master database not loaded"})

        # ====================================================================
        # STRICT PRIVACY RULE: Match Lead in Database
        # ====================================================================
        match_idx = df_master[df_master["Phone"].str.endswith(clean_ph)].index
        
        if len(match_idx) == 0:
            print(f"[PRIVACY_GUARD] Ignored message from non-database contact {clean_ph} ('{message[:30]}...'). Personal chats protected.")
            return {
                "status": "ignored_non_lead",
                "phone": clean_ph,
                "message": "Sender is not in SmartWork customer database. Ignored to protect personal chats."
            }

        # Lead found in SmartWork database
        idx = match_idx[0]
        lead_record = df_master.iloc[idx].to_dict()
        cust_name = lead_record.get("Name", "Customer")
        now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        print(f"[SMARTWORK_AI] Verified Lead: {cust_name} ({clean_ph}) sent: '{message}'")

        # 1. Update Database with customer reply & classify
        lower_msg = message.lower()
        hot_keywords = ["demo", "link", "trial", "price", "kharidna", "call", "bhejo", "de do", "batao", "free", "try", "use", "karein", "kaise", "haan", "ha", "yes", "ok", "intrested", "interested", "accha", "sahi"]
        is_hot = any(k in lower_msg for k in hot_keywords)
        new_status = "HOT_LEAD" if is_hot else "REPLIED"
        
        df_master.at[idx, "OutreachStatus"] = new_status
        df_master.at[idx, "ReplyText"] = message
        df_master.at[idx, "ReplyAt"] = now_iso
        save_master_data(df_master)

        ai_reply = ""
        send_success = False
        send_error = ""

        # 2. Dispatch AI Response ONLY if AUTO_REPLY_ENABLED is True
        if AUTO_REPLY_ENABLED:
            ai_reply, _ = generate_gemini_sales_reply(lead_record, message)
            try:
                out_payload = {"number": f"91{clean_ph}", "text": ai_reply}
                req_data = json.dumps(out_payload).encode("utf-8")
                gw_url = f"{EVOLUTION_API_URL}/message/sendText/smartwork_outreach"
                req = urllib.request.Request(gw_url, data=req_data, headers={"Content-Type": "application/json", "User-Agent": "SmartWork-Hub"}, method="POST")
                with urllib.request.urlopen(req, timeout=12) as res:
                    send_success = True
                    print(f"[AI_OUTREACH_DISPATCHED] To {cust_name} ({clean_ph}): '{ai_reply}'")
            except Exception as send_err:
                send_error = str(send_err)
                print(f"[AI_DISPATCH_FAILED] To {clean_ph}: {send_error}")
        else:
            print(f"[AUTO_REPLY_PAUSED] Reply from {cust_name} ({clean_ph}) recorded. Auto-reply is OFF.")

        return {
            "status": "success_lead_replied",
            "phone": clean_ph,
            "name": cust_name,
            "category": lead_record.get("Category"),
            "customer_message": message,
            "ai_response": ai_reply,
            "new_status": new_status,
            "auto_dispatched": send_success,
            "dispatch_error": send_error if not send_success else None
        }

    except Exception as e:
        print(f"[WEBHOOK_EXCEPTION] {e}")
        return JSONResponse(status_code=400, content={"error": str(e)})


@app.get("/download")
def download_excel():
    if os.path.exists(EXCEL_MASTER):
        return FileResponse(
            EXCEL_MASTER,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="FINAL_COMMON_TYPISTS.xlsx"
        )
    return JSONResponse(status_code=404, content={"error": "File not found."})


# ============================================================================
# EVOLUTION / WHATSAPP GATEWAY INTEGRATION ENDPOINTS
# ============================================================================

EVOLUTION_API_URL = os.environ.get("EVOLUTION_API_URL", "https://smartwork-wa-evolution.onrender.com")

@app.get("/api/whatsapp/status")
def get_whatsapp_status():
    """Checks WhatsApp connection status from Render Evolution Gateway."""
    try:
        req = urllib.request.Request(f"{EVOLUTION_API_URL}/instance/connectionState/smartwork_outreach", headers={"User-Agent": "SmartWork-Hub"})
        with urllib.request.urlopen(req, timeout=8) as res:
            data = json.loads(res.read().decode("utf-8"))
            inst = data.get("instance", {})
            state = inst.get("state", "close")
            user = inst.get("user")
            phone = user.get("id") if user else None
            return {
                "status": "CONNECTED" if state == "open" else ("CONNECTING" if state == "connecting" else "DISCONNECTED"),
                "state": state,
                "phone": phone,
                "gateway_url": EVOLUTION_API_URL
            }
    except Exception as e:
        return {"status": "DISCONNECTED", "state": "close", "phone": None, "error": str(e), "gateway_url": EVOLUTION_API_URL}

@app.get("/api/whatsapp/qr")
def get_whatsapp_qr():
    """Fetches live base64 QR code from Render Evolution Gateway."""
    try:
        req = urllib.request.Request(f"{EVOLUTION_API_URL}/instance/connect/smartwork_outreach", headers={"User-Agent": "SmartWork-Hub"})
        with urllib.request.urlopen(req, timeout=10) as res:
            data = json.loads(res.read().decode("utf-8"))
            qr_b64 = data.get("base64") or data.get("qrcode", {}).get("base64")
            state = data.get("state", "connecting")
            is_connected = state == "open" or data.get("connected", False)
            phone = data.get("phone")
            return {
                "connected": is_connected,
                "qr": qr_b64,
                "state": state,
                "phone": phone
            }
    except Exception as e:
        return {"connected": False, "qr": None, "error": str(e)}

@app.post("/api/whatsapp/send")
def send_whatsapp_direct(
    phone: str = Query(...),
    message: str = Query(...)
):
    """Sends outbound WhatsApp message directly through Render Evolution Gateway."""
    clean_ph = re.sub(r"\D", "", str(phone))
    if len(clean_ph) >= 10:
        clean_ph = clean_ph[-10:]

    payload = {
        "number": f"91{clean_ph}",
        "text": message.strip()
    }
    
    try:
        req_data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{EVOLUTION_API_URL}/message/sendText/smartwork_outreach",
            data=req_data,
            headers={"Content-Type": "application/json", "User-Agent": "SmartWork-Hub"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=15) as res:
            res_data = json.loads(res.read().decode("utf-8"))
            
            # Update Lead in Excel as SENT
            df_master, _, _, _, _ = get_master_data()
            if not df_master.empty:
                match_idx = df_master[df_master["Phone"].str.endswith(clean_ph)].index
                if len(match_idx) > 0:
                    idx = match_idx[0]
                    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df_master.at[idx, "OutreachStatus"] = "SENT"
                    if not df_master.at[idx, "SentAt"]:
                        df_master.at[idx, "SentAt"] = now_iso
                    save_master_data(df_master)
                    
            return {"status": "success", "phone": clean_ph, "response": res_data}
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8")
        return JSONResponse(status_code=e.code, content={"error": err_msg})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.post("/api/whatsapp/logout")
def logout_whatsapp():
    """Logs out WhatsApp session on Render Evolution Gateway."""
    try:
        req = urllib.request.Request(
            f"{EVOLUTION_API_URL}/instance/logout/smartwork_outreach",
            headers={"User-Agent": "SmartWork-Hub"},
            method="DELETE"
        )
        with urllib.request.urlopen(req, timeout=10) as res:
            return json.loads(res.read().decode("utf-8"))
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
