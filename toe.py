# file: toe.py
import os
import json
import argparse
import datetime as dt
import win32com.client
from flask import Flask, jsonify, render_template_string, request

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

CONFIG = load_config()

def safe_getattr(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default

def get_week_range(offset_weeks=0):
    today = dt.datetime.now() - dt.timedelta(weeks=offset_weeks)
    start_of_week = today - dt.timedelta(days=today.weekday())
    start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_week = start_of_week + dt.timedelta(days=4, hours=23, minutes=59, seconds=59)
    week_num = int(start_of_week.strftime("%V"))
    return start_of_week, end_of_week, week_num

def get_week_label(offset_weeks=0):
    s, e, w = get_week_range(offset_weeks)
    return f"Week {w}: {s.strftime('%b %d')} → {e.strftime('%b %d')}", w

def export_week_events(offset_weeks=0):
    start_of_week, end_of_week, week_num = get_week_range(offset_weeks)
    print(f"📅 Exporting Outlook events for Week {week_num} ({start_of_week:%b %d} → {end_of_week:%b %d})")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    items = calendar.Items
    items.IncludeRecurrences = True
    try:
        items.Sort("[Start]")
    except Exception:
        pass

    filter_str = (
        f"[Start] >= '{start_of_week:%m/%d/%Y %I:%M %p}' AND "
        f"[End] <= '{end_of_week:%m/%d/%Y %I:%M %p}'"
    )
    try:
        restricted_items = items.Restrict(filter_str)
    except Exception:
        restricted_items = items

    events = []
    for appt in restricted_items:
        try:
            start_time = safe_getattr(appt, "Start")
            end_time   = safe_getattr(appt, "End")
            if not start_time or not end_time:
                continue
            duration = int((end_time - start_time).total_seconds() / 60)
            events.append({
                "subject": safe_getattr(appt, "Subject", "") or "",
                "start": start_time.strftime("%a, %b %d %H:%M"),  # "Mon, Nov 03 09:00"
                "end": end_time.strftime("%H:%M"),
                "duration_minutes": duration,
                "categories": safe_getattr(appt, "Categories", "") or "",
                "date": start_time.strftime("%d/%b/%y")
            })
        except Exception:
            continue

    os.makedirs(DATA_DIR, exist_ok=True)
    file_path = os.path.join(DATA_DIR, f"events_W{week_num}.json")
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump({
            "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "week_number": week_num,
            "event_count": len(events),
            "events": events
        }, f, indent=2, ensure_ascii=False)
    print(f"✅ Saved {len(events)} events → {file_path}")
    return file_path

app = Flask(__name__)

# raw string to keep JS escapes intact
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en" x-data="app()">
<head>
<meta charset="UTF-8">
<title>TOE — Time Optimization Engine</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js" defer></script>
<style>[x-cloak]{display:none !important;}</style>
</head>
<body class="bg-gray-50 font-sans">

<div class="max-w-5xl mx-auto p-6" x-init="init()" x-cloak>
  <h1 class="text-3xl font-bold mb-4 text-gray-800">🧭 TOE — Time Optimization Engine</h1>

  <!-- Week selector + actions -->
  <div class="flex flex-wrap items-center justify-between mb-6">
    <div>
      <label class="font-semibold mr-2">Select Week:</label>
      <select x-model.number="selectedWeek" @change="loadWeek()" class="border rounded p-2">
        <template x-for="(week, i) in weeks" :key="'w_'+week.week+'_'+i">
          <option :value="week.week" x-text="week.label"></option>
        </template>
      </select>
    </div>
    <div class="text-right mt-2 sm:mt-0">
      <p class="text-gray-700 font-semibold" x-text="'Entries needing Jira: ' + categorizedCount"></p>
      <button @click="openPtoModal()" class="ml-3 bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700">Generate JSON</button>
      <p x-show="message" class="text-green-600 mt-2" x-text="message"></p>
    </div>
  </div>

  <!-- Category Utilization -->
  <div class="mb-6">
    <div class="flex items-center justify-between mb-2">
      <h2 class="font-semibold text-gray-800">Category Utilization</h2>
      <span class="text-sm text-gray-600" x-text="totalHours().toFixed(1) + ' h total'"></span>
    </div>

    <div class="relative" x-ref="utilBarWrap">
      <div class="w-full h-8 rounded bg-gray-200 overflow-hidden flex" x-ref="utilBar">
        <template x-for="(seg, i) in segments" :key="'seg_'+seg.name+'_'+i">
          <div class="h-full cursor-pointer"
               :style="`width:${seg.percent}%; background:${seg.color};`"
               @mousemove="moveTooltip($event)"
               @mouseenter="showSeg(seg, $event)"
               @mouseleave="hoverSeg=null">
          </div>
        </template>
      </div>
      <template x-if="hoverSeg">
        <div class="absolute -top-10 px-2 py-1 text-xs rounded bg-black text-white whitespace-nowrap pointer-events-none"
             :style="tooltipStyle">
          <span x-text="hoverSeg.name"></span> —
          <span x-text="hoverSeg.hours.toFixed(1)"></span>h
          (<span x-text="hoverSeg.percent.toFixed(1)"></span>%)
        </div>
      </template>
      <div class="mt-3 flex flex-wrap gap-3">
        <template x-for="(seg, i) in segments" :key="'legend_'+seg.name+'_'+i">
          <div class="flex items-center gap-2">
            <span class="w-3 h-3 inline-block rounded" :style="`background:${seg.color};`"></span>
            <span class="text-sm text-gray-700" x-text="seg.name + ' (' + seg.hours.toFixed(1) + 'h, ' + seg.percent.toFixed(0) + '%)'"></span>
          </div>
        </template>
      </div>
    </div>
  </div>

  <!-- Day Selector + Filter -->
  <div class="flex flex-wrap items-center justify-between mb-6">
    <div class="flex flex-wrap gap-2">
      <template x-for="(day, i) in weekdays" :key="'d_'+day+'_'+i">
        <button
          @click="selectedDay = day"
          :class="selectedDay === day ? 'bg-blue-600 text-white' : 'bg-white text-gray-700 border'"
          class="px-3 py-2 rounded shadow hover:bg-blue-100 transition text-sm"
          :title="day"
        >
          <div class="flex flex-col leading-tight">
            <span class="font-semibold" x-text="day"></span>
            <span class="text-[11px] opacity-80" x-text="dayDateLabel(day)"></span>
          </div>
        </button>
      </template>
    </div>
    <div class="flex items-center space-x-2 mt-3 sm:mt-0">
      <input type="checkbox" id="showCategorized" x-model="showOnlyCategorized" class="w-4 h-4">
      <label for="showCategorized" class="text-gray-700 text-sm">Show only categorized events</label>
    </div>
  </div>

  <!-- Daily Events List -->
  <div class="max-w-3xl">
    <template x-for="(ev) in filteredEvents()" :key="ev.__key">
      <div class="bg-white p-3 rounded-lg shadow mb-3 hover:shadow-md transition flex items-start space-x-3">
        <input type="checkbox" class="mt-1 w-5 h-5 text-blue-600"
               :id="ev.event_id + '_done'"
               x-model="doneMap[ev.event_id]"
               @change="saveDone()">
        <div class="flex-1">
          <!-- INLINE EDITABLE TITLE + 'Edited' badge -->
          <template x-if="!ev.editing">
            <div class="flex items-center gap-2">
              <h3 class="font-semibold text-blue-600 text-lg truncate cursor-text hover:underline"
                  @click="startEdit(ev)"
                  :title="'Click to edit title'">
                <span x-text="ev.subject || 'Untitled'"></span>
              </h3>
              <span x-show="isEdited(ev)"
                    class="text-[10px] uppercase tracking-wide bg-yellow-100 text-yellow-800 px-1.5 py-0.5 rounded"
                    title="Title differs from Outlook original">Edited</span>
            </div>
          </template>
          <template x-if="ev.editing">
            <input type="text"
                   x-model="ev.subject_draft"
                   @keydown.enter.prevent="commitEdit(ev)"
                   @keydown.escape.prevent="cancelEdit(ev)"
                   @blur="commitEdit(ev)"
                   class="w-full border rounded px-2 py-1 font-semibold text-blue-700"
                   :placeholder="ev.subject || 'Untitled'"
                   autofocus>
          </template>

          <!-- TIME/DATE -->
          <p class="text-sm text-gray-700 mb-1" x-text="(ev.start || '') + (ev.end ? ' → ' + ev.end : '')"></p>
          <p class="text-xs text-gray-500">📅 <span x-text="ev.date || ''"></span></p>
          <p class="text-xs text-gray-600 mb-1">⏱ <span x-text="(ev.duration_minutes || 0) + 'm'"></span></p>

          <!-- Outlook Categories -->
          <template x-if="(ev.categories || '').trim().length">
            <div class="mt-1 flex flex-wrap items-center gap-2">
              <span class="text-xs text-gray-500">Outlook Categories:</span>
              <template x-for="(cat, i) in (ev.categories || '').split(',').map(c => c.trim()).filter(Boolean)" :key="ev.event_id + '_cat_' + i">
                <span class="text-xs font-medium px-2 py-1 rounded bg-gray-100 text-gray-700">
                  <span x-text="cat"></span>
                </span>
              </template>
            </div>
          </template>

          <!-- JIRA -->
          <template x-if="jiraMapFor(ev).length">
            <div class="mt-1 flex flex-wrap items-center gap-2">
              <span class="text-xs text-gray-500">JIRA TIMECODE:</span>
              <template x-for="(jira, i) in jiraMapFor(ev)" :key="ev.event_id + '_jira_' + i">
                <span
                  :style="jiraTagStyle(jira)"
                  class="text-xs font-bold px-2 py-1 rounded shadow-sm">
                  <span x-text="jira"></span>
                </span>
              </template>
            </div>
          </template>
        </div>
      </div>
    </template>
    <p x-show="filteredEvents().length === 0" class="text-gray-500 italic">No events for this day.</p>
  </div>

  <!-- PTO Modal (unchanged from previous fix) -->
  <div x-show="ptoModalOpen" class="fixed inset-0 z-50 flex items-center justify-center">
    <div class="absolute inset-0 bg-black/40" @click="ptoModalOpen=false"></div>
    <div class="relative bg-white w-full max-w-2xl rounded-lg shadow-lg p-5">
      <div class="flex items-center justify-between mb-3">
        <h3 class="text-lg font-semibold">Mark PTO / Sick for this week</h3>
        <button @click="ptoModalOpen=false" class="text-gray-500 hover:text-gray-700">✕</button>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-sm">
          <thead class="text-left text-gray-600">
            <tr>
              <th class="py-2">Day</th>
              <th class="py-2">Status</th>
              <th class="py-2">Hours (Half)</th>
              <th class="py-2">Reason</th>
            </tr>
          </thead>
          <tbody>
            <template x-for="(row) in ptoRows" :key="'ptorow_'+row.day">
              <tr class="border-t">
                <td class="py-2">
                  <div class="flex flex-col">
                    <span class="font-medium" x-text="row.day"></span>
                    <span class="text-xs text-gray-500" x-text="row.date || '-'"></span>
                  </div>
                </td>
                <td class="py-2">
                  <select x-model="row.status" @change="onPtoStatusChange(row)" class="border rounded px-2 py-1">
                    <option value="none">None</option>
                    <option value="half">Half</option>
                    <option value="full">Full</option>
                  </select>
                </td>
                <td class="py-2">
                  <input type="number"
                         min="0.5" max="7.5" step="0.5"
                         x-model.number="row.hours"
                         :disabled="row.status!=='half'"
                         @blur="row.status==='half' && (row.hours = normalizeHalfHours(row.hours))"
                         class="border rounded px-2 py-1 w-24"
                         :placeholder="row.status==='half' ? '4' : '—'">
                  <div class="text-[11px] text-gray-500" x-show="row.status==='half'">Default 4h • Editable</div>
                </td>
                <td class="py-2">
                  <select x-model="row.reason" class="border rounded px-2 py-1">
                    <option value="PTO">PTO</option>
                    <option value="Sick">Sick</option>
                  </select>
                </td>
              </tr>
            </template>
          </tbody>
        </table>
      </div>
      <div class="mt-4 flex justify-end gap-3">
        <button @click="ptoModalOpen=false" class="px-4 py-2 rounded border">Cancel</button>
        <button @click="submitPto()" class="px-4 py-2 rounded bg-blue-600 text-white hover:bg-blue-700">Continue & Generate</button>
      </div>
    </div>
  </div>

  <!-- Toast -->
  <div x-show="toastVisible" x-transition class="fixed bottom-6 right-6 bg-green-600 text-white px-4 py-2 rounded shadow-lg text-sm">
    <span x-text="toastMessage"></span>
  </div>
</div>

<script>
function app() {
  return {
    // state
    events: [],
    weeks: [],
    cfg: { categories: {}, palette: [] },
    edits: {},                 // event_id -> {subject}
    selectedWeek: 0,
    weekdays: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
    selectedDay: 'Mon',
    dayLabelMap: {},
    dayDateMap: {},
    summaryMinutes: {},        // cfgKey -> minutes
    segments: [],
    categorizedCount: 0,       // now = entries needing Jira (no mapped category)
    message: '',
    doneMap: {},
    showOnlyCategorized: false,
    toastVisible: false,
    toastMessage: '',
    hoverSeg: null,
    tooltipStyle: '',

    // PTO
    ptoModalOpen: false,
    ptoRows: [],

    // derived
    outlookToCfg: {},
    cfgToJira: {},
    jiraKeyToColor: {},

    palette: ['#93c5fd','#86efac','#fcd34d','#fca5a5','#a5b4fc','#f9a8d4','#fbbf24','#f87171','#34d399','#60a5fa'],

    async init() {
      const cfgRes = await fetch('/config');
      this.cfg = await cfgRes.json();
      if (Array.isArray(this.cfg.palette) && this.cfg.palette.length) this.palette = this.cfg.palette.slice();

      // build maps
      this.outlookToCfg = {};
      this.cfgToJira = {};
      this.jiraKeyToColor = {};
      Object.entries(this.cfg.categories || {}).forEach(([cfgKey, obj]) => {
        const outlookName = (obj.outlook_category || cfgKey).trim();
        this.outlookToCfg[outlookName] = cfgKey;
        this.cfgToJira[cfgKey] = Array.isArray(obj.jira_timecodes) ? obj.jira_timecodes.slice() : [];
        const hex = (obj.color || '#999999');
        (this.cfgToJira[cfgKey]).forEach(tc => {
          const key = (tc.split(' - ')[0] || tc).trim();
          this.jiraKeyToColor[key] = hex;
        });
      });

      const res = await fetch('/weeks');
      this.weeks = await res.json();
      this.selectedWeek = this.weeks[0]?.week || 0;

      this.loadDone();
      await this.loadWeek();
    },

    async loadWeek() {
      // load events
      const res = await fetch('/data/' + this.selectedWeek);
      const data = await res.json();
      const raw = Array.isArray(data.events) ? data.events : [];

      // normalized events
      this.events = raw.map((e, i) => {
        const start = typeof e.start === 'string' ? e.start : '';
        const end = typeof e.end === 'string' ? e.end : '';
        const date = typeof e.date === 'string' ? e.date : '';
        const event_id = `${start}|${end}|${date}`;
        const subject = typeof e.subject === 'string' ? e.subject : '';
        return {
          original_subject: subject,
          subject,
          subject_draft: subject,
          editing: false,
          start, end, date,
          categories: typeof e.categories === 'string' ? e.categories : '',
          duration_minutes: Number.isFinite(+e.duration_minutes) ? Number(e.duration_minutes) : 0,
          event_id,
          __key: `${event_id}|${i}`,
        };
      });

      // apply saved edits
      await this.loadEdits();
      this.applyEdits();

      this.rebuildDayLabels();
      this.computeSummary();

      // default to first day with events
      if (this.filteredEvents().length === 0) {
        const daysWithEvents = this.weekdays.filter(d => this.events.some(e => this.weekdayOf(e) === d));
        if (daysWithEvents.length) this.selectedDay = daysWithEvents[0];
      }
      this.message = '';
      this.buildPtoRows();
    },

    async loadEdits() {
      try {
        const res = await fetch('/edits/' + this.selectedWeek);
        const data = await res.json();
        this.edits = data.edits || {};
      } catch { this.edits = {}; }
    },
    applyEdits() {
      this.events = this.events.map(e => {
        const patch = this.edits[e.event_id];
        if (patch && typeof patch.subject === 'string' && patch.subject.trim().length) {
          e.subject = patch.subject;
          e.subject_draft = patch.subject;
        }
        return e;
      });
    },

    // inline edit
    startEdit(ev) { ev.subject_draft = ev.subject || ''; ev.editing = true; },
    cancelEdit(ev) { ev.subject_draft = ev.subject; ev.editing = false; },
    async commitEdit(ev) {
      const newTitle = (ev.subject_draft || '').trim();
      if (!newTitle || newTitle === ev.subject) { ev.editing = false; return; }
      ev.subject = newTitle; ev.editing = false;
      try {
        await fetch('/edits/' + this.selectedWeek, {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({ event_id: ev.event_id, subject: newTitle })
        });
        this.edits[ev.event_id] = { subject: newTitle };
      } catch {
        this.toastMessage = '❌ Failed to save edit';
        this.toastVisible = true;
        setTimeout(() => this.toastVisible = false, 1500);
      }
    },
    isEdited(ev) { return (ev.subject || '') !== (ev.original_subject || ''); },

    // PTO (unchanged)
    openPtoModal() { this.buildPtoRows(); this.ptoModalOpen = true; },
    buildPtoRows() {
      this.ptoRows = this.weekdays.map(day => ({
        day,
        date: this.dayDateMap[day] || '',
        status: 'none',
        hours: 0,
        reason: 'PTO',
      }));
    },
    onPtoStatusChange(row) {
      if (row.status === 'none') row.hours = 0;
      else if (row.status === 'half') row.hours = this.normalizeHalfHours(row.hours || 4);
      else if (row.status === 'full') row.hours = 8;
    },
    normalizeHalfHours(h) {
      let val = Number(h || 0);
      if (!Number.isFinite(val)) val = 4;
      val = Math.max(0.5, Math.min(7.5, val));
      return Math.round(val * 2) / 2;
    },
    async submitPto() {
      const payload = {
        pto: this.ptoRows
          .filter(r => r.status !== 'none' && r.date)
          .map(r => ({
            date: r.date,
            status: r.status,
            hours: r.status === 'half' ? this.normalizeHalfHours(r.hours) : 8,
            reason: r.reason
          }))
      };
      try {
        const res = await fetch('/generate_json/' + this.selectedWeek, {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        this.message = data.message || '✅ Jira JSON generated';
        this.toastMessage = this.message; this.toastVisible = true;
        setTimeout(() => this.toastVisible = false, 1500);
      } catch {
        this.message = '❌ Failed generating JSON';
      } finally { this.ptoModalOpen = false; }
    },

    // mapping helpers
    mapTokenToCfg(token) {
      const t = (token || '').trim();
      if (!t) return null;
      if (this.outlookToCfg[t]) return this.outlookToCfg[t];
      if ((this.cfg.categories || {})[t]) return t;
      return null;
    },

    // >>> FIXED: compute segments from mapped cats; compute "needs Jira" count
    computeSummary() {
      this.summaryMinutes = {};
      let needs = 0;

      for (const e of this.events) {
        const dur = Number(e.duration_minutes || 0);
        const tokens = (e.categories || '').split(',').map(c => c.trim()).filter(Boolean);
        const mapped = Array.from(new Set(tokens.map(this.mapTokenToCfg.bind(this)).filter(Boolean)));

        if (mapped.length === 0) {
          needs += 1;                      // no mapped category → needs Jira
          continue;
        }
        for (const cfgKey of mapped) {
          this.summaryMinutes[cfgKey] = (this.summaryMinutes[cfgKey] || 0) + dur;
        }
      }

      this.categorizedCount = needs;

      const total = this.totalMinutes();
      const names = Object.keys(this.summaryMinutes).sort((a,b) => this.summaryMinutes[b] - this.summaryMinutes[a]);

      this.segments = names.map((cfgKey, idx) => {
        const minutes = this.summaryMinutes[cfgKey];
        const percent = total > 0 ? (minutes / total) * 100 : 0;
        const cfgColor = (this.cfg.categories?.[cfgKey]?.color) || this.palette[idx % this.palette.length];
        return { name: cfgKey, minutes, hours: minutes / 60, percent, color: cfgColor };
      });
    },

    totalMinutes() {
      const vals = Object.values(this.summaryMinutes);
      if (!vals.length) return 0;
      return vals.reduce((a,b) => a + b, 0);
    },
    totalHours() { return this.totalMinutes() / 60; },

    // day helpers
    weekdayOf(e) {
      const token = (e && typeof e.start === 'string') ? e.start.split(',')[0] : '';
      return (token || '').trim().slice(0,3);
    },
    monthDayOf(e) {
      if (!e || typeof e.start !== 'string') return '';
      const parts = e.start.split(',');
      if (parts.length < 2) return '';
      const rest = parts[1].trim(); // "Nov 03 09:00"
      const chunks = rest.split(/\s+/);
      if (chunks.length < 2) return '';
      return `${chunks[0]} ${chunks[1]}`;
    },
    rebuildDayLabels() {
      const label = {}, dates = {};
      this.events.forEach(e => {
        const wd = this.weekdayOf(e);
        const md = this.monthDayOf(e);
        if (wd && md && !label[wd]) label[wd] = md;
        if (wd && (e.date || '') && !dates[wd]) dates[wd] = e.date;
      });
      this.dayLabelMap = label;
      this.dayDateMap = dates;
    },
    dayDateLabel(day) { return this.dayLabelMap[day] || ''; },

    filteredEvents() {
      try {
        return this.events.filter(e => {
          const dayMatch = this.weekdayOf(e) === this.selectedDay;
          const hasCat = (e.categories || '').trim().length > 0;
          return dayMatch && (!this.showOnlyCategorized || hasCat);
        });
      } catch { return []; }
    },

    jiraMapFor(ev) {
      const tokens = (ev.categories || '').split(',').map(c => c.trim()).filter(Boolean);
      const list = [];
      tokens.forEach(tok => {
        const cfgKey = this.mapTokenToCfg(tok);
        if (!cfgKey) return;
        (this.cfgToJira[cfgKey] || []).forEach(tc => list.push(tc));
      });
      return Array.from(new Set(list));
    },
    jiraTagStyle(jira) {
      const key = (jira.split(' - ')[0] || jira).trim();
      const bg = this.jiraKeyToColor[key] || '#374151';
      return `background:${bg};color:#ffffff`;
    },

    saveDone() { localStorage.setItem('doneMap', JSON.stringify(this.doneMap)); },
    loadDone() { try { this.doneMap = JSON.parse(localStorage.getItem('doneMap')) || {}; } catch { this.doneMap = {}; } },

    // Tooltip
    showSeg(seg, ev) { this.hoverSeg = { name: seg.name, hours: seg.hours, percent: seg.percent }; this.moveTooltip(ev); },
    moveTooltip(ev) {
      const wrap = this.$refs.utilBar;
      if (!wrap || !this.hoverSeg) return;
      const rect = wrap.getBoundingClientRect();
      const pointerX = ev.clientX - rect.left;
      const tipWidth = 160, half = tipWidth/2, minX = 6+half, maxX = rect.width-6-half;
      const center = Math.max(minX, Math.min(maxX, pointerX));
      this.tooltipStyle = `left:${center-half}px; top:-36px;`;
    },
  }
}
</script>

</body>
</html>
"""

# ---------------------------- Routes ---------------------------------
@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/config")
def serve_config():
    return jsonify({
        "categories": CONFIG.get("categories", {}),
        "palette": CONFIG.get("palette", [])
    })

@app.route("/weeks")
def list_weeks():
    weeks = []
    for i in range(0, 6):
        label, wnum = get_week_label(i)
        weeks.append({"week": wnum, "label": label})
    return jsonify(weeks)

def _edits_path(week_num: int) -> str:
    os.makedirs(DATA_DIR, exist_ok=True)
    return os.path.join(DATA_DIR, f"event_edits_W{week_num}.json")

def _load_edits(week_num: int):
    path = _edits_path(week_num)
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
        return data.get("edits", {}) if isinstance(data, dict) else {}

def _save_edits(week_num: int, edits: dict):
    path = _edits_path(week_num)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"edits": edits}, f, indent=2)

@app.route("/edits/<int:week_num>", methods=["GET", "POST"])
def edits_api(week_num):
    if request.method == "GET":
        return jsonify({"edits": _load_edits(week_num)})
    data = request.get_json(silent=True) or {}
    event_id = data.get("event_id")
    subject = (data.get("subject") or "").strip()
    if not event_id or not subject:
        return jsonify({"message": "event_id and subject required"}), 400
    edits = _load_edits(week_num)
    edits[event_id] = {"subject": subject}
    _save_edits(week_num, edits)
    return jsonify({"message": "Saved", "edits": edits})

@app.route("/data/<int:week_num>")
def serve_data(week_num):
    file_path = os.path.join(DATA_DIR, f"events_W{week_num}.json")
    if not os.path.exists(file_path):
        return jsonify({"events": []})
    with open(file_path, "r", encoding="utf-8") as f:
        return jsonify(json.load(f))

@app.route("/generate_json/<int:week_num>", methods=["POST"])
def generate_json(week_num):
    file_path = os.path.join(DATA_DIR, f"events_W{week_num}.json")
    if not os.path.exists(file_path):
        return jsonify({"message": f"No events file found for week {week_num}"}), 404

    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # mappings
    cats = CONFIG.get("categories", {})
    outlook_to_cfg, cfg_to_jira = {}, {}
    for cfg_key, obj in cats.items():
        outlook_name = (obj.get("outlook_category") or cfg_key).strip()
        outlook_to_cfg[outlook_name] = cfg_key
        cfg_to_jira[cfg_key] = list(obj.get("jira_timecodes", []))

    # PTO payload
    pto_payload = request.get_json(silent=True) or {}
    pto_rows = pto_payload.get("pto") or []
    pto_full_dates = set([r["date"] for r in pto_rows if r.get("status") == "full" and r.get("date")])
    pto_partial_hours = {r["date"]: float(r.get("hours", 0)) for r in pto_rows if r.get("status") == "half" and r.get("date")}
    pto_reason_by_date = {r["date"]: (r.get("reason") or "PTO") for r in pto_rows if r.get("date")}

    # PTO timecode from config
    pto_cfg = cats.get("PTO or Sick") or {}
    pto_tc_full = (pto_cfg.get("jira_timecodes") or ["DXGLV-10 - Sick or PTO"])[0]

    # Load edits
    edits = _load_edits(week_num)

    grouped = {}
    included_count = 0

    # process events (skip full PTO days)
    for e in data.get("events", []):
        date_str = e.get("date", "")
        if date_str in pto_full_dates:
            continue

        # override subject if edited
        start = e.get("start", "") or ""
        end = e.get("end", "") or ""
        event_id = f"{start}|{end}|{date_str}"
        subject = (edits.get(event_id, {}) or {}).get("subject", e.get("subject", ""))

        tokens = [c.strip() for c in (e.get("categories") or "").split(",") if c.strip()]
        mapped_timecodes = []
        for tok in tokens:
            cfg_key = outlook_to_cfg.get(tok) or (tok if tok in cfg_to_jira else None)
            if not cfg_key:
                continue
            mapped_timecodes.extend(cfg_to_jira.get(cfg_key, []))
        if not mapped_timecodes:
            continue

        grouped.setdefault(date_str, []).append({
            "jira_categories": mapped_timecodes,
            "subject": subject,
            "duration_minutes": e.get("duration_minutes", 0),
            "start": start,
            "end": end,
            "date": date_str
        })
        included_count += 1

    # PTO entries
    for date_str in pto_full_dates:
        grouped.setdefault(date_str, []).append({
            "jira_categories": [pto_tc_full],
            "subject": f"{pto_reason_by_date.get(date_str, 'PTO')} - Full Day",
            "duration_minutes": 8 * 60,
            "start": "",
            "end": "",
            "date": date_str
        })
        included_count += 1

    for date_str, hours in pto_partial_hours.items():
        if hours <= 0:
            continue
        grouped.setdefault(date_str, []).append({
            "jira_categories": [pto_tc_full],
            "subject": f"{pto_reason_by_date.get(date_str, 'PTO')} - Partial ({hours:g}h)",
            "duration_minutes": int(hours * 60),
            "start": "",
            "end": "",
            "date": date_str
        })
        included_count += 1

    out_path = os.path.join(DATA_DIR, f"jira_export_W{week_num}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(grouped, f, indent=2)

    return jsonify({"message": f"✅ Jira JSON saved to {out_path} ({included_count} entries incl. edits & PTO)"})


def main():
    parser = argparse.ArgumentParser(description="TOE — Time Optimization Engine")
    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("load", help="Load current + previous 5 weeks Outlook events")
    subparsers.add_parser("display", help="Run local web server")
    args = parser.parse_args()

    if args.command == "load":
        for w in range(0, 6):
            export_week_events(w)
    elif args.command == "display":
        print("🌐 Starting TOE at http://127.0.0.1:5000/")
        app.run(debug=False, port=5000)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
