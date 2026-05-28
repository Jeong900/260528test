from __future__ import annotations

import calendar
import json
import re
import threading
import uuid
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from xml.etree.ElementTree import iterparse


ROOT = Path(r"C:\Users\User\Desktop\0528")
CACHE_DIR = ROOT / ".dashboard_cache"
PORT = 8765
FOLDER_RE = re.compile(r"^\d{4}$")
FILE_RE = re.compile(r"^(\d{6})_데이터\(전체\)(?:_dummy)?\.xlsx$")
ALL_UNITS = "DS부문 전체"
CACHE_VERSION = 7

JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


def parse_year_month(folder_name: str) -> tuple[int, int]:
    if not FOLDER_RE.match(folder_name):
        raise ValueError("폴더명은 YYMM 형식이어야 합니다.")
    year = 2000 + int(folder_name[:2])
    month = int(folder_name[2:])
    if month < 1 or month > 12:
        raise ValueError("폴더명의 월 정보가 올바르지 않습니다.")
    return year, month


def money_to_int(value) -> int:
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    try:
        return int(float(str(value).replace(",", "").strip()))
    except ValueError:
        return 0


def empty_metric() -> dict[str, int]:
    return {"participants": 0, "amount": 0}


def folder_signature(folder: Path) -> list[dict]:
    return [
        {"name": file.name, "size": file.stat().st_size, "mtime": file.stat().st_mtime}
        for file in sorted(folder.glob("*.xlsx"))
    ]


def xml_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def cell_col(ref: str) -> str:
    return "".join(ch for ch in ref if ch.isalpha())


def read_shared_strings(book: zipfile.ZipFile) -> list[str]:
    try:
        source = book.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    strings: list[str] = []
    parts: list[str] = []
    in_si = False
    for event, elem in iterparse(source, events=("start", "end")):
        name = xml_name(elem.tag)
        if event == "start" and name == "si":
            in_si = True
            parts = []
        elif event == "end" and in_si and name == "t":
            parts.append(elem.text or "")
        elif event == "end" and name == "si":
            strings.append("".join(parts))
            in_si = False
            elem.clear()
    source.close()
    return strings


def sheet_path(book: zipfile.ZipFile) -> str:
    if "xl/worksheets/sheet1.xml" in book.namelist():
        return "xl/worksheets/sheet1.xml"
    for name in book.namelist():
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            return name
    raise ValueError("엑셀 시트를 찾을 수 없습니다.")


def cell_value(cell, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = ""
    inline_text: list[str] = []
    for child in cell:
        name = xml_name(child.tag)
        if name == "v":
            value = child.text or ""
        elif name == "is":
            for sub in child.iter():
                if xml_name(sub.tag) == "t":
                    inline_text.append(sub.text or "")
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return ""
    if cell_type == "inlineStr":
        return "".join(inline_text)
    return value


def read_workbook(path: Path) -> tuple[dict[str, int], dict[str, dict[str, int]]]:
    total = empty_metric()
    by_unit: dict[str, dict[str, int]] = {}

    with zipfile.ZipFile(path) as book:
        shared_strings = read_shared_strings(book)
        with book.open(sheet_path(book)) as source:
            for _event, row in iterparse(source, events=("end",)):
                if xml_name(row.tag) != "row":
                    continue
                row_number = int(row.attrib.get("r", "0") or 0)
                if row_number < 3:
                    row.clear()
                    continue

                values = {"A": "", "B": "", "I": ""}
                for cell in row:
                    if xml_name(cell.tag) != "c":
                        continue
                    col = cell_col(cell.attrib.get("r", ""))
                    if col in values:
                        values[col] = cell_value(cell, shared_strings)

                if values["A"] == "DS(Korea)":
                    unit = values["B"] or "미지정"
                    amount = money_to_int(values["I"])
                    total["participants"] += 1
                    total["amount"] += amount
                    bucket = by_unit.setdefault(unit, empty_metric())
                    bucket["participants"] += 1
                    bucket["amount"] += amount
                row.clear()
    return total, by_unit


def set_job(job_id: str, **updates) -> None:
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(updates)


def analyze_folder(folder_name: str, job_id: str | None = None) -> dict:
    year, month = parse_year_month(folder_name)
    folder = ROOT / folder_name
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder}")

    CACHE_DIR.mkdir(exist_ok=True)
    signature = folder_signature(folder)
    cache_path = CACHE_DIR / f"{folder_name}_ds_metrics.json"
    if cache_path.exists():
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
        if cached.get("version") == CACHE_VERSION and cached.get("signature") == signature:
            data = cached["data"]
            data["cached"] = True
            if job_id:
                set_job(job_id, processed=len(signature), total=len(signature), currentFile="캐시 사용")
            return data

    candidates: list[tuple[str, Path]] = []
    ignored_files: list[str] = []
    for file in sorted(folder.glob("*.xlsx")):
        match = FILE_RE.match(file.name)
        if not match:
            ignored_files.append(file.name)
            continue
        date_code = match.group(1)
        if not date_code.startswith(folder_name):
            ignored_files.append(file.name)
            continue
        try:
            date_obj = datetime.strptime(date_code, "%y%m%d")
        except ValueError:
            ignored_files.append(file.name)
            continue
        if date_obj.year != year or date_obj.month != month:
            ignored_files.append(file.name)
            continue
        candidates.append((date_code, file))

    if job_id:
        set_job(job_id, total=len(candidates), processed=0, currentFile="파일 목록 확인 완료")

    by_date: dict[str, dict] = {}
    unit_names: set[str] = set()
    for index, (date_code, file) in enumerate(candidates, start=1):
        if job_id:
            set_job(job_id, processed=index - 1, currentFile=file.name)
        total, by_unit = read_workbook(file)
        unit_names.update(by_unit.keys())
        by_date[date_code] = {"fileName": file.name, "total": total, "businessUnits": by_unit}
        if job_id:
            set_job(job_id, processed=index, currentFile=file.name)

    _, max_day = calendar.monthrange(year, month)
    days = []
    for day in range(1, max_day + 1):
        date_code = f"{year % 100:02d}{month:02d}{day:02d}"
        date_obj = datetime(year, month, day)
        day_data = by_date.get(date_code)
        days.append(
            {
                "dateCode": date_code,
                "day": day,
                "weekday": (date_obj.weekday() + 1) % 7,
                "fileName": day_data["fileName"] if day_data else None,
                "total": day_data["total"] if day_data else empty_metric(),
                "businessUnits": day_data["businessUnits"] if day_data else {},
            }
        )

    data = {
        "folder": folder_name,
        "year": year,
        "month": month,
        "monthLabel": f"{year}년 {month}월",
        "daysInMonth": max_day,
        "fileCount": len(by_date),
        "businessUnits": [ALL_UNITS] + sorted(unit_names),
        "calendar": days,
        "ignoredFiles": ignored_files,
        "cached": False,
    }
    cache_path.write_text(
        json.dumps({"version": CACHE_VERSION, "signature": signature, "data": data}, ensure_ascii=False),
        encoding="utf-8",
    )
    return data


def start_analysis_job(folder_name: str) -> str:
    job_id = uuid.uuid4().hex
    with JOBS_LOCK:
        JOBS[job_id] = {
            "id": job_id,
            "status": "running",
            "folder": folder_name,
            "processed": 0,
            "total": 0,
            "currentFile": "분석 준비 중",
            "data": None,
            "error": None,
        }

    def run() -> None:
        try:
            data = analyze_folder(folder_name, job_id)
            set_job(job_id, status="completed", data=data, currentFile="완료")
        except Exception as exc:
            set_job(job_id, status="error", error=str(exc), currentFile="오류")

    threading.Thread(target=run, daemon=True).start()
    return job_id


HTML = r"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>DS부문 기부 대시보드</title>
  <style>
    :root{--ink:#151922;--muted:#687386;--line:#d9dee7;--panel:#fff;--page:#eef1f5;--blue:#1428a0;--green:#0b755d;--red:#b42318;--shadow:0 12px 30px rgba(21,25,34,.08)}
    *{box-sizing:border-box}body{margin:0;background:var(--page);color:var(--ink);font-family:"Segoe UI","Malgun Gothic",Arial,sans-serif;letter-spacing:0}
    header{background:var(--panel);border-bottom:1px solid var(--line)}.topbar{max-width:1320px;margin:0 auto;padding:22px 28px;display:grid;grid-template-columns:1fr auto;gap:20px;align-items:end}
    h1{margin:0;font-size:25px;font-weight:760}.sub{margin-top:7px;color:var(--muted);font-size:13px}main{max-width:1320px;margin:0 auto;padding:24px 28px 36px}
    .intro{min-height:430px;background:var(--panel);border:1px solid var(--line);border-radius:8px;box-shadow:var(--shadow);display:grid;place-items:center;padding:48px 24px}.intro-inner{width:min(560px,100%);text-align:center}.intro h2{margin:0 0 12px;font-size:24px;font-weight:760}.intro p{margin:0 0 24px;color:var(--muted);font-size:14px;line-height:1.6}
    button,select{height:40px;border-radius:6px;border:1px solid var(--line);background:var(--panel);color:var(--ink);padding:0 12px;font:inherit;font-size:13px}button{border-color:var(--blue);background:var(--blue);color:#fff;font-weight:700;cursor:pointer;padding:0 18px}button.secondary{background:var(--panel);color:var(--ink);border-color:var(--line)}button:disabled{opacity:.55;cursor:wait}
    .toolbar{display:none;gap:10px;align-items:end;justify-content:flex-end;flex-wrap:wrap}.control{display:grid;gap:5px}.control label{color:var(--muted);font-size:11px;font-weight:700}.control select{min-width:190px}
    .tabs{display:none;gap:6px;margin-bottom:14px}.tab-btn{height:36px;background:#fff;color:var(--ink);border:1px solid var(--line);padding:0 14px}.tab-btn.active{background:var(--blue);border-color:var(--blue);color:#fff}.page{display:none}.page.active{display:block}
    .calendar-wrap,.board-card{background:var(--panel);border:1px solid var(--line);border-radius:8px;box-shadow:var(--shadow);overflow:hidden}.section-head{min-height:58px;padding:14px 18px;border-bottom:1px solid var(--line);display:grid;grid-template-columns:1fr auto;gap:14px;align-items:center}.section-head h2{margin:0;font-size:16px;font-weight:740}.status{color:var(--muted);font-size:12px;white-space:nowrap}.selection{color:var(--ink);font-size:13px;font-weight:700;white-space:nowrap}
    .weekdays,.calendar{display:grid;grid-template-columns:repeat(7,minmax(0,1fr))}.weekdays{background:#fbfcfe;border-bottom:1px solid var(--line)}.weekdays div{padding:10px 12px;color:var(--muted);font-size:12px;font-weight:700;text-align:center}.calendar{padding:10px;gap:8px}.day{min-height:132px;border:1px solid var(--line);border-radius:8px;background:#fff;padding:11px;display:flex;flex-direction:column;gap:8px}.day.empty{border-color:transparent;background:transparent}.day.has-data{border-color:#e7c94f;background:linear-gradient(180deg,#fff 0%,var(--heat,#fff8cf) 100%)}
    .day-top{display:flex;align-items:center;justify-content:space-between;gap:8px}.day-no{font-size:15px;font-weight:760}.badge{border-radius:99px;padding:3px 8px;font-size:11px;font-weight:760}.badge.ok{color:#7a5a00;background:rgba(255,255,255,.62)}.badge.no{color:var(--red);background:#fff1f0}
    .value{flex:1;min-width:0;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center}.value span{display:block;color:var(--muted);font-size:12px;margin-bottom:6px}.value strong{display:block;font-size:clamp(15px,1.55vw,24px);font-weight:780;font-variant-numeric:tabular-nums;line-height:1.12;white-space:nowrap;overflow:hidden;max-width:100%}
    .board{display:none;gap:16px}.board.active{display:grid}.chart-wrap{padding:14px 16px 18px}.chart{width:100%;height:280px;display:block}.compare-layout{display:grid;grid-template-columns:minmax(0,1fr) 420px;gap:0;border-top:1px solid var(--line)}.compare-layout .chart-wrap{border-right:1px solid var(--line)}
    table{width:100%;border-collapse:collapse;font-size:12px}th,td{padding:10px 8px;border-bottom:1px solid #eef1f5;text-align:left}th{color:var(--muted);font-weight:700;background:#fbfcfe}td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}.legend{display:flex;gap:12px;flex-wrap:wrap;color:var(--muted);font-size:12px;margin-top:10px}.legend i{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:5px}
    .error{display:none;margin-top:16px;border:1px solid #f3b2b2;background:#fff4f4;color:#9d1c1c;border-radius:8px;padding:12px 14px;font-size:13px}.loading{position:fixed;inset:0;background:rgba(238,241,245,.88);display:none;place-items:center;z-index:10}.loading-card{width:min(440px,calc(100vw - 32px));background:#fff;border:1px solid var(--line);border-radius:8px;box-shadow:var(--shadow);padding:24px}.loading-title{font-size:18px;font-weight:760;margin-bottom:8px}.loading-note{color:var(--muted);font-size:13px;line-height:1.5;margin-bottom:18px}.progress{height:8px;background:#edf0f5;border-radius:999px;overflow:hidden}.progress-bar{height:100%;width:0%;background:var(--blue);transition:width .25s ease}.progress-meta{display:flex;justify-content:space-between;gap:12px;color:var(--muted);font-size:12px;margin-top:10px}.current-file{margin-top:12px;color:var(--ink);font-size:12px;overflow-wrap:anywhere}
    input[type=file]{position:fixed;left:-9999px;width:1px;height:1px;opacity:0}@media(max-width:980px){.topbar,.section-head,.compare-layout{grid-template-columns:1fr}.toolbar{justify-content:flex-start}.compare-layout .chart-wrap{border-right:0;border-bottom:1px solid var(--line)}}@media(max-width:720px){main,.topbar{padding-left:14px;padding-right:14px}.weekdays{display:none}.calendar{grid-template-columns:1fr}.control,.control select,button{width:100%}}
  </style>
</head>
<body>
  <header><div class="topbar"><div><h1>DS부문 기부 대시보드</h1><div class="sub" id="subtitle">분석하고자 하는 월별 폴더를 선택해 주세요.</div></div><div class="toolbar" id="toolbar"><div class="control"><label for="unitSelect">사업부</label><select id="unitSelect"></select></div><div class="control"><label for="metricSelect">표시 지표</label><select id="metricSelect"><option value="participants">참여인원</option><option value="amount">기부금액</option><option value="average">평균 기부액</option></select></div><button class="secondary" id="changeBtn">Open</button></div></div></header>
  <main>
    <section class="intro" id="intro"><div class="intro-inner"><h2>분석하고자 하는 폴더를 선택해 주세요.</h2><p>대용량 파일도 고려해 로컬 폴더를 서버가 직접 읽고, 처리 중에는 진행 상태를 표시합니다. 한 번 분석한 폴더는 캐시되어 다음 조회가 빠릅니다.</p><button id="openBtn">Open</button></div></section>
    <nav class="tabs" id="pageTabs"><button class="tab-btn active" data-page="calendar">달력</button><button class="tab-btn" data-page="trend">추이/비교</button></nav>
    <section class="calendar-wrap page active" id="calendarWrap"><div class="section-head"><div><h2 id="calendarTitle">달력</h2><div class="status" id="status">대기</div></div><div class="selection" id="selectionLabel">DS부문 전체 · 참여인원</div></div><div class="weekdays"><div>일</div><div>월</div><div>화</div><div>수</div><div>목</div><div>금</div><div>토</div></div><div class="calendar" id="calendar"></div></section>
    <section class="page board" id="trendPage"><div class="board-card"><div class="section-head"><div><h2>날짜별 추이</h2><div class="status">선택 기준의 일별 변화</div></div><div class="selection" id="trendLabel">DS부문 전체 · 참여인원</div></div><div class="chart-wrap"><svg class="chart" id="trendChart" role="img" aria-label="날짜별 추이 선그래프"></svg></div></div><div class="board-card"><div class="section-head"><div><h2>사업부별 비교</h2><div class="status">상위 사업부 기준 선그래프와 월간 요약</div></div></div><div class="compare-layout"><div class="chart-wrap"><svg class="chart" id="compareChart" role="img" aria-label="사업부별 비교 선그래프"></svg><div class="legend" id="compareLegend"></div></div><div id="compareTable"></div></div></div></section>
    <div class="error" id="errorBox"></div><input id="folderInput" type="file" webkitdirectory directory multiple />
  </main>
  <div class="loading" id="loading"><div class="loading-card"><div class="loading-title">로딩중입니다</div><div class="loading-note">선택한 월별 데이터를 읽고 있습니다. 대용량 파일은 잠시 시간이 걸릴 수 있습니다.</div><div class="progress"><div class="progress-bar" id="progressBar"></div></div><div class="progress-meta"><span id="progressText">준비 중</span><span id="progressPct">0%</span></div><div class="current-file" id="currentFile">분석 준비 중</div></div></div>
  <script>
    const openBtn=document.getElementById("openBtn"),changeBtn=document.getElementById("changeBtn"),folderInput=document.getElementById("folderInput"),intro=document.getElementById("intro"),toolbar=document.getElementById("toolbar"),unitSelect=document.getElementById("unitSelect"),metricSelect=document.getElementById("metricSelect"),calendarWrap=document.getElementById("calendarWrap"),calendarEl=document.getElementById("calendar"),statusEl=document.getElementById("status"),errorBox=document.getElementById("errorBox"),selectionLabel=document.getElementById("selectionLabel"),loading=document.getElementById("loading"),progressBar=document.getElementById("progressBar"),progressText=document.getElementById("progressText"),progressPct=document.getElementById("progressPct"),currentFile=document.getElementById("currentFile"),pageTabs=document.getElementById("pageTabs"),trendPage=document.getElementById("trendPage"),trendChart=document.getElementById("trendChart"),compareChart=document.getElementById("compareChart"),compareLegend=document.getElementById("compareLegend"),compareTable=document.getElementById("compareTable"),trendLabel=document.getElementById("trendLabel");
    const fmt=new Intl.NumberFormat("ko-KR"),folderRe=/^\d{4}$/,allUnits="DS부문 전체",colors=["#1428a0","#0b755d","#9c6b00","#7f56d9","#b42318","#0077b6"];let currentData=null;
    const metricMeta={participants:{label:"참여인원",suffix:"명"},amount:{label:"기부금액",suffix:"원"},average:{label:"평균 기부액",suffix:"원"}};
    function showError(message){errorBox.textContent=message;errorBox.style.display="block"}function clearError(){errorBox.textContent="";errorBox.style.display="none"}function showLoading(){loading.style.display="grid";setProgress({processed:0,total:0,currentFile:"분석 준비 중"})}function hideLoading(){loading.style.display="none"}
    function setProgress(job){const total=job.total||0,processed=job.processed||0,pct=total?Math.round(processed/total*100):8;progressBar.style.width=`${Math.max(8,Math.min(100,pct))}%`;progressText.textContent=total?`${processed} / ${total}개 파일 처리`:"파일 목록 확인 중";progressPct.textContent=total?`${pct}%`:"";currentFile.textContent=job.currentFile||"분석 중"}
    function getFolderName(files){if(!files.length)throw new Error("선택된 파일이 없습니다.");const folderName=(files[0].webkitRelativePath||files[0].name).split("/")[0];if(!folderRe.test(folderName))throw new Error("폴더명은 YYMM 형식이어야 합니다. 예: 2603");return folderName}
    async function analyze(files){const folderName=getFolderName(files);showLoading();openBtn.disabled=true;changeBtn.disabled=true;const start=await fetch(`/api/start-analysis?folder=${encodeURIComponent(folderName)}`);const started=await start.json();if(!start.ok)throw new Error(started.error||"분석을 시작하지 못했습니다.");return await pollJob(started.jobId)}
    async function pollJob(jobId){for(;;){const res=await fetch(`/api/job?id=${encodeURIComponent(jobId)}`);const job=await res.json();if(!res.ok)throw new Error(job.error||"작업 상태를 확인하지 못했습니다.");setProgress(job);if(job.status==="completed")return job.data;if(job.status==="error")throw new Error(job.error||"분석 중 오류가 발생했습니다.");await new Promise(resolve=>setTimeout(resolve,650))}}
    function metricSource(day,unit){if(!day.fileName)return null;if(unit===allUnits)return day.total;return day.businessUnits[unit]||{participants:0,amount:0}}function valueFor(source,metric){if(!source)return null;if(metric==="participants")return source.participants;if(metric==="amount")return source.amount;return source.participants?Math.round(source.amount/source.participants):0}
    function formatValue(value){return value===null?"-":fmt.format(value)}function fitCalendarValues(){document.querySelectorAll(".value strong").forEach(node=>{node.style.fontSize="";let size=parseFloat(getComputedStyle(node).fontSize);while(node.scrollWidth>node.clientWidth&&size>12){size-=1;node.style.fontSize=`${size}px`}})}
    function heatColor(value,min,max){if(value===null)return"#fff";const range=max-min;const normalized=range>0?(value-min)/range:.45;const lightness=97-Math.round(normalized*25);const saturation=82+Math.round(normalized*13);return`hsl(48 ${saturation}% ${lightness}%)`}
    function renderControls(data){unitSelect.innerHTML=data.businessUnits.map(unit=>`<option value="${unit}">${unit}</option>`).join("");unitSelect.value=allUnits;metricSelect.value="participants";document.getElementById("subtitle").textContent=`${data.monthLabel} · DS부문`;document.getElementById("calendarTitle").textContent=`${data.monthLabel} 달력`;statusEl.textContent=`${data.fileCount}개 파일 반영${data.cached?" · 캐시 사용":""}`;selectionLabel.textContent=`${allUnits} · 참여인원`}
    function renderCalendar(data){const unit=unitSelect.value||allUnits,metric=metricSelect.value||"participants",meta=metricMeta[metric];selectionLabel.textContent=`${unit} · ${meta.label}`;calendarEl.innerHTML="";const values=data.calendar.map(day=>valueFor(metricSource(day,unit),metric)).filter(value=>value!==null);const min=values.length?Math.min(...values):0,max=values.length?Math.max(...values):0;const first=data.calendar[0];for(let i=0;i<first.weekday;i++){const empty=document.createElement("div");empty.className="day empty";calendarEl.appendChild(empty)}data.calendar.forEach(day=>{const value=valueFor(metricSource(day,unit),metric),el=document.createElement("div");el.className=value===null?"day":"day has-data";if(value!==null)el.style.setProperty("--heat",heatColor(value,min,max));const badge=value===null?`<span class="badge no">파일 없음</span>`:`<span class="badge ok">반영</span>`;const body=value===null?`<div class="value"><span>${meta.label}</span><strong>-</strong></div>`:`<div class="value"><span>${meta.label}</span><strong>${formatValue(value)}</strong></div>`;el.innerHTML=`<div class="day-top"><div class="day-no">${day.day}</div>${badge}</div>${body}`;calendarEl.appendChild(el)});requestAnimationFrame(fitCalendarValues)}
    function lineChart(svg,series){const w=900,h=280,p={l:54,r:18,t:18,b:34};svg.setAttribute("viewBox",`0 0 ${w} ${h}`);svg.innerHTML="";const vals=series.flatMap(s=>s.values.filter(v=>v!==null));const max=vals.length?Math.max(...vals):1,min=vals.length?Math.min(...vals):0,range=max-min||1,days=currentData.calendar.map(d=>d.day),n=days.length;const sx=i=>p.l+(n<=1?0:i*(w-p.l-p.r)/(n-1)),sy=v=>h-p.b-((v-min)/range)*(h-p.t-p.b);for(let i=0;i<5;i++){const y=p.t+i*(h-p.t-p.b)/4;svg.insertAdjacentHTML("beforeend",`<line x1="${p.l}" y1="${y}" x2="${w-p.r}" y2="${y}" stroke="#e8ecf2"/><text x="8" y="${y+4}" font-size="11" fill="#687386">${fmt.format(Math.round(max-(range*i/4)))}</text>`)}days.forEach((d,i)=>{if(i%5===0||i===n-1)svg.insertAdjacentHTML("beforeend",`<text x="${sx(i)}" y="${h-10}" text-anchor="middle" font-size="11" fill="#687386">${d}</text>`)});series.forEach((s,si)=>{const pts=s.values.map((v,i)=>v===null?null:[sx(i),sy(v)]).filter(Boolean);if(pts.length>1)svg.insertAdjacentHTML("beforeend",`<polyline points="${pts.map(p=>p.join(",")).join(" ")}" fill="none" stroke="${s.color||colors[si%colors.length]}" stroke-width="2.4"/>`);pts.forEach(pt=>svg.insertAdjacentHTML("beforeend",`<circle cx="${pt[0]}" cy="${pt[1]}" r="2.6" fill="${s.color||colors[si%colors.length]}"/>`))})}
    function dayValues(unit,metric){return currentData.calendar.map(day=>valueFor(metricSource(day,unit),metric))}
    function renderTrend(data){const unit=unitSelect.value||allUnits,metric=metricSelect.value||"participants",meta=metricMeta[metric];trendLabel.textContent=`${unit} · ${meta.label}`;lineChart(trendChart,[{name:unit,color:"#1428a0",values:dayValues(unit,metric)}]);const units=data.businessUnits.filter(u=>u!==allUnits);const ranked=units.map(unit=>{const vals=dayValues(unit,metric).filter(v=>v!==null);const latest=[...dayValues(unit,metric)].reverse().find(v=>v!==null)??0;const avg=vals.length?Math.round(vals.reduce((a,b)=>a+b,0)/vals.length):0;return{unit,latest,avg,min:vals.length?Math.min(...vals):0,max:vals.length?Math.max(...vals):0,values:dayValues(unit,metric)}}).sort((a,b)=>b.latest-a.latest).slice(0,6);lineChart(compareChart,ranked.map((r,i)=>({name:r.unit,color:colors[i%colors.length],values:r.values})));compareLegend.innerHTML=ranked.map((r,i)=>`<span><i style="background:${colors[i%colors.length]}"></i>${r.unit}</span>`).join("");compareTable.innerHTML=`<table><thead><tr><th>사업부</th><th class="num">최근</th><th class="num">평균</th><th class="num">최소</th><th class="num">최대</th></tr></thead><tbody>${ranked.map(r=>`<tr><td>${r.unit}</td><td class="num">${formatValue(r.latest)}</td><td class="num">${formatValue(r.avg)}</td><td class="num">${formatValue(r.min)}</td><td class="num">${formatValue(r.max)}</td></tr>`).join("")}</tbody></table>`}
    function renderAll(){if(!currentData)return;renderCalendar(currentData);renderTrend(currentData)}function showDashboard(data){currentData=data;renderControls(data);renderAll();intro.style.display="none";toolbar.style.display="flex";pageTabs.style.display="flex";calendarWrap.classList.add("active")}
    function openPicker(){folderInput.value="";folderInput.click()}openBtn.addEventListener("click",openPicker);changeBtn.addEventListener("click",openPicker);unitSelect.addEventListener("change",renderAll);metricSelect.addEventListener("change",renderAll);pageTabs.addEventListener("click",event=>{const btn=event.target.closest(".tab-btn");if(!btn)return;document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));btn.classList.add("active");document.querySelectorAll(".page").forEach(p=>p.classList.remove("active"));(btn.dataset.page==="calendar"?calendarWrap:trendPage).classList.add("active");if(btn.dataset.page==="trend")renderTrend(currentData)});
    folderInput.addEventListener("change",async()=>{clearError();try{showDashboard(await analyze([...folderInput.files]))}catch(err){showError(err.message)}finally{hideLoading();openBtn.disabled=false;changeBtn.disabled=false}});
  </script>
</body>
</html>
"""


class DashboardHandler(BaseHTTPRequestHandler):
    def send_json(self, payload, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path in ("/", "/index.html"):
                html_path = ROOT / "dashboard.html"
                body = html_path.read_bytes() if html_path.exists() else HTML.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            elif parsed.path == "/api/start-analysis":
                folder = parse_qs(parsed.query).get("folder", [""])[0]
                self.send_json({"jobId": start_analysis_job(folder)})
            elif parsed.path == "/api/job":
                job_id = parse_qs(parsed.query).get("id", [""])[0]
                with JOBS_LOCK:
                    job = JOBS.get(job_id)
                    payload = dict(job) if job else None
                if not payload:
                    self.send_json({"error": "작업을 찾을 수 없습니다."}, status=404)
                else:
                    self.send_json(payload)
            elif parsed.path == "/api/analyze-folder":
                folder = parse_qs(parsed.query).get("folder", [""])[0]
                self.send_json(analyze_folder(folder))
            else:
                self.send_error(404)
        except Exception as exc:
            self.send_json({"error": str(exc)}, status=400)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    server = ThreadingHTTPServer(("127.0.0.1", PORT), DashboardHandler)
    print(f"http://127.0.0.1:{PORT}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
