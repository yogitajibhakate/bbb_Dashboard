import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

const CATS = {
  master: 'Master Member List',
  kyt: 'KYT Done',
  attendance: 'Attendance',
  referrals: 'Referrals Given',
  business: 'Business Given',
  inductions: 'Inductions',
  visitors: 'Visitors / VIP / Observer',
  testimonials: 'Testimonials',
  bbp: 'BBP Score'
};

const TOTAL_MEETINGS = 24;

// NAME MATCHING & CONSOLIDATION UTILITIES
function nn(n) {
  return (n || '')
    .toString()
    .toLowerCase()
    .trim()
    .replace(/^(dr|mr|mrs|ms)\.?\s+/g, '') // strip leading titles
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function jw(s1, s2) {
  s1 = nn(s1);
  s2 = nn(s2);
  if (s1 === s2) return 1;
  if (!s1.length || !s2.length) return 0;
  const mw = Math.floor(Math.max(s1.length, s2.length) / 2) - 1;
  const m1 = new Array(s1.length).fill(false);
  const m2 = new Array(s2.length).fill(false);
  let mt = 0, tr = 0;
  for (let i = 0; i < s1.length; i++) {
    const st = Math.max(0, i - mw);
    const en = Math.min(s2.length - 1, i + mw);
    for (let j = st; j <= en; j++) {
      if (!m2[j] && s1[i] === s2[j]) {
        m1[i] = m2[j] = true;
        mt++;
        break;
      }
    }
  }
  if (!mt) return 0;
  let k = 0;
  for (let i = 0; i < s1.length; i++) {
    if (m1[i]) {
      while (!m2[k]) k++;
      if (s1[i] !== s2[k]) tr++;
      k++;
    }
  }
  const jar = (mt / s1.length + mt / s2.length + (mt - tr / 2) / mt) / 3;
  let pf = 0;
  for (let i = 0; i < Math.min(4, s1.length, s2.length); i++) {
    if (s1[i] === s2[i]) pf++;
    else break;
  }
  return jar + pf * 0.1 * (1 - jar);
}

function bestMatch(raw, names) {
  if (!raw) return null;
  const nr = nn(raw);
  if (!nr) return null;
  const ex = names.find(n => nn(n) === nr);
  if (ex) return ex;
  let bst = null, hi = 0;
  names.forEach(n => {
    const s = jw(raw, n);
    if (s > hi) {
      hi = s;
      bst = n;
    }
  });
  if (hi >= 0.95) {
    // Safety check: verify that last words/initials are not conflicting
    const w1 = nr.split(/\s+/).filter(t => t.length > 0);
    const w2 = nn(bst).split(/\s+/).filter(t => t.length > 0);
    if (w1.length > 1 && w2.length > 1) {
      const last1 = w1[w1.length - 1];
      const last2 = w2[w2.length - 1];
      if (last1 !== last2 && jw(last1, last2) < 0.85) {
        return null; // different last names or initials
      }
    }
    return bst;
  }
  return null;
}

function consolidate(files) {
  if (!files || !files.length) return [];
  const mf = files.filter(f => f.category === 'master');
  let all = [];
  if (mf.length) {
    const ns = new Set();
    mf.forEach(f => f.rows.forEach(r => ns.add(r.rawName)));
    all = Array.from(ns);
  } else {
    const ns = new Set();
    files.forEach(f => f.rows.forEach(r => ns.add(r.rawName)));
    all = Array.from(ns);
  }
  const uniq = [];
  all.forEach(n => {
    if (!bestMatch(n, uniq)) uniq.push(n);
  });
  const mg = uniq.map(name => ({
    name,
    kyt: 0,
    attendance: '',
    referrals: 0,
    business: 0,
    inductions: 0,
    visitors: 0,
    testimonials: 0,
    bbp: 0,
    score: 0
  }));

  files.forEach(file => {
    file.rows.forEach(row => {
      const mn = bestMatch(row.rawName, uniq);
      if (!mn) return;
      const rec = mg.find(m => m.name === mn);
      if (!rec) return;
      const vk = Object.keys(row).filter(k => k !== 'rawName');
      let matchedAny = false;

      vk.forEach(k => {
        const lowerK = k.toLowerCase();
        const val = row[k];
        if (val === undefined || val === '') return;

        let matched = false;

        if (/attendance|status|present/.test(lowerK)) {
          const s = String(val).trim().toUpperCase();
          if (s.startsWith('P') || s === 'PRESENT') rec.attendance = 'P';
          else if (s.startsWith('L') || s === 'LATE') rec.attendance = 'L';
          else if (s.startsWith('S') || s === 'SUBSTITUTE' || s === 'SUB') rec.attendance = 'S';
          else if (s.startsWith('A') || s === 'ABSENT') rec.attendance = 'A';
          matched = true;
        }

        if (/kyt/.test(lowerK)) {
          const n = parseFloat(val);
          rec.kyt = isNaN(n) ? val : n;
          matched = true;
        }

        if (/visitor|vip|observer/.test(lowerK)) {
          rec.visitors += parseFloat(val) || 0;
          matched = true;
        }

        if (/referral/.test(lowerK)) {
          rec.referrals += parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
          matched = true;
        }

        if (/business|revenue/.test(lowerK)) {
          rec.business += parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
          matched = true;
        }

        if (/induction/.test(lowerK)) {
          rec.inductions += parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
          matched = true;
        }

        if (/testimonial/.test(lowerK)) {
          rec.testimonials += parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
          matched = true;
        }

        if (/bbp/.test(lowerK)) {
          const v = parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
          rec.bbp = Math.max(rec.bbp, v);
          matched = true;
        }

        if (matched) {
          matchedAny = true;
        }
      });

      // Fallback: if no column matched standard BNI metric headers, use the sheet's top-level category on the main value column
      if (!matchedAny && vk.length > 0) {
        const k = vk.find(x => /score|points|count|value|total|given/.test(x.toLowerCase())) || vk[0];
        if (k) {
          const val = row[k];
          if (val !== undefined && val !== '') {
            if (file.category === 'attendance') {
              const s = String(val).trim().toUpperCase();
              if (s.startsWith('P')) rec.attendance = 'P';
              else if (s.startsWith('L')) rec.attendance = 'L';
              else if (s.startsWith('S')) rec.attendance = 'S';
              else if (s.startsWith('A')) rec.attendance = 'A';
            } else if (file.category === 'kyt') {
              const n = parseFloat(val);
              rec.kyt = isNaN(n) ? val : n;
            } else if (file.category === 'visitors') {
              rec.visitors += parseFloat(val) || 0;
            } else {
              const v = parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
              if (file.category === 'referrals') rec.referrals += v;
              else if (file.category === 'business') rec.business += v;
              else if (file.category === 'inductions') rec.inductions += v;
              else if (file.category === 'testimonials') rec.testimonials += v;
              else if (file.category === 'bbp') rec.bbp = Math.max(rec.bbp, v);
            }
          }
        }
      }
    });
  });

  mg.forEach(m => {
    let sc = 0;
    const ks = String(m.kyt).toLowerCase();
    if ((typeof m.kyt === 'number' && m.kyt > 0) || /yes|done|^1$/.test(ks)) sc += 10;
    if (m.attendance === 'P') sc += 10;
    else if (m.attendance === 'L') sc -= 5;
    else if (m.attendance === 'S') sc += 5;
    if (m.referrals > 0) sc += 10;
    if (m.business > 0) sc += 20;
    if (m.inductions > 0) sc += 20;
    if (m.visitors > 0) sc += 10;
    if (m.testimonials > 0) sc += 10;
    sc += m.bbp * 2;
    m.score = sc;
  });
  return mg;
}

function computeOverall(data) {
  const meetingResults = {};
  Object.entries(data).forEach(([mk, files]) => {
    if (Array.isArray(files) && files.length) {
      meetingResults[mk] = consolidate(files);
    }
  });

  if (!Object.keys(meetingResults).length) return [];

  const allNamesSet = new Set();
  Object.values(meetingResults).forEach(members => {
    members.forEach(m => allNamesSet.add(m.name));
  });
  const allNames = Array.from(allNamesSet);

  const uniqueNames = [];
  allNames.forEach(name => {
    if (!bestMatch(name, uniqueNames)) uniqueNames.push(name);
  });

  const merged = uniqueNames.map(name => ({
    name,
    kyt: 0,
    attendance: '',
    referrals: 0,
    business: 0,
    inductions: 0,
    visitors: 0,
    testimonials: 0,
    bbp: 0,
    score: 0,
    _meetingCount: 0
  }));

  Object.values(meetingResults).forEach(meetingMembers => {
    meetingMembers.forEach(mMember => {
      const canonName = bestMatch(mMember.name, uniqueNames);
      if (!canonName) return;
      const rec = merged.find(m => m.name === canonName);
      if (!rec) return;

      rec.referrals += mMember.referrals;
      rec.business += mMember.business;
      rec.inductions += mMember.inductions;
      rec.visitors += mMember.visitors;
      rec.testimonials += mMember.testimonials;
      rec.bbp += mMember.bbp;

      const ks = String(mMember.kyt).toLowerCase();
      if ((typeof mMember.kyt === 'number' && mMember.kyt > 0) || /yes|done/.test(ks) || ks === '1') {
        rec.kyt = (typeof rec.kyt === 'number' ? rec.kyt : 0) + 1;
      }

      const rank = { P: 4, S: 3, L: 2, A: 1, '': 0 };
      if ((rank[mMember.attendance] || 0) > (rank[rec.attendance] || 0)) {
        rec.attendance = mMember.attendance;
      }

      rec.score += mMember.score;
      rec._meetingCount++;
    });
  });

  merged.sort((a, b) => b.score - a.score);
  return merged;
}

export default function App() {
  const [data, setData] = useState({});
  const [cur, setCur] = useState('overall');
  const [theme, setTheme] = useState('light');
  const [toasts, setToasts] = useState([]);
  const [activeTab, setActiveTab] = useState('lbTab');
  const [isDragOver, setIsDragOver] = useState(false);

  // Table status states
  const [tsrch, setTsrch] = useState('');
  const [ovSearch, setOvSearch] = useState('');
  const [sb, setSb] = useState('score');
  const [so, setSo] = useState('desc');
  const [ps, setPs] = useState(25);
  const [cp, setCp] = useState(1);
  const [cf, setCf] = useState({});
  const [cw, setCw] = useState({});

  const fileInputRef = useRef(null);

  // Fetch state from server on mount
  useEffect(() => {
    fetchData();
    const storedTheme = localStorage.getItem('bni_theme') || 'light';
    setTheme(storedTheme);
    document.documentElement.setAttribute('data-theme', storedTheme);
  }, []);

  const fetchData = async () => {
    try {
      const res = await fetch('/api/dashboard');
      if (res.ok) {
        const state = await res.json();
        setData(state.data || {});
        setCur(state.cur || 'overall');
      }
    } catch (err) {
      console.error('Error fetching dashboard state:', err);
      toast('Failed to load data from server.', 'error');
    }
  };

  const toast = (msg, type = 'info') => {
    const id = Date.now() + Math.random().toString(36).substr(2, 5);
    setToasts(prev => [...prev, { id, msg, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 3500);
  };

  const toggleTheme = () => {
    const nextTheme = theme === 'light' ? 'dark' : 'light';
    setTheme(nextTheme);
    localStorage.setItem('bni_theme', nextTheme);
    document.documentElement.setAttribute('data-theme', nextTheme);
  };

  // CALCULATE MEMBERS LIST DYNAMICALLY BASED ON ACTIVE SELECTION
  const members = cur === 'overall' ? computeOverall(data) : consolidate(data[cur] || []);

  // RENDER DYNAMIC CHARTS
  useEffect(() => {
    if (!window.Plotly) return;

    const dark = theme === 'dark';
    const txtC = dark ? '#f5eedc' : '#3c321d';
    const gridC = dark ? 'rgba(255,255,255,.08)' : 'rgba(146,108,21,.12)';
    const baseLayout = {
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Inter,sans-serif', color: txtC },
      showlegend: true,
      legend: { orientation: 'h', y: -0.18, font: { size: 11 } }
    };

    if (cur === 'overall') {
      let present = 0, absent = 0, late = 0, sub2 = 0;
      let green = 0, orange = 0, red = 0;

      members.forEach(m => {
        if (m.attendance === 'P') present++;
        else if (m.attendance === 'A') absent++;
        else if (m.attendance === 'L') late++;
        else if (m.attendance === 'S') sub2++;

        if (m.score >= 60) green++;
        else if (m.score >= 30) orange++;
        else red++;
      });

      // Overall Pie Chart
      const pieEl = document.getElementById('ov-ch-pie');
      if (pieEl) {
        window.Plotly.newPlot(pieEl, [{
          values: [green, orange, red],
          labels: ['Green (60+)', 'Orange (30–59)', 'Red (<30)'],
          type: 'pie',
          hole: 0.5,
          marker: { colors: ['#10b981', '#f97316', '#ef4444'] },
          textinfo: 'value+percent',
          hovertemplate: '%{label}<br>%{value} members (%{percent})<extra></extra>'
        }], { ...baseLayout, margin: { t: 10, b: 30, l: 10, r: 10 } }, { responsive: true });
      }

      // Attendance Pie Chart
      const attEl = document.getElementById('ov-ch-att');
      if (attEl) {
        window.Plotly.newPlot(attEl, [{
          values: [present, late, sub2, absent],
          labels: ['Present', 'Late', 'Substitute', 'Absent'],
          type: 'pie',
          hole: 0.5,
          marker: { colors: ['#10b981', '#f59e0b', '#38bdf8', '#ef4444'] },
          textinfo: 'value+percent',
          hovertemplate: '%{label}<br>%{value} members<extra></extra>'
        }], { ...baseLayout, margin: { t: 10, b: 30, l: 10, r: 10 } }, { responsive: true });
      }

      // Top 15 Bar Chart
      const topEl = document.getElementById('ov-ch-top');
      if (topEl && members.length) {
        const top15 = [...members].sort((a, b) => b.score - a.score).slice(0, 15);
        window.Plotly.newPlot(topEl, [{
          x: top15.map(m => m.name),
          y: top15.map(m => m.score),
          type: 'bar',
          marker: { color: top15.map(m => m.score >= 60 ? '#10b981' : m.score >= 30 ? '#f97316' : '#ef4444') },
          text: top15.map(m => m.score + 'pts'),
          textposition: 'outside',
          hovertemplate: '%{x}<br>Score: %{y}pts<extra></extra>'
        }], {
          ...baseLayout,
          showlegend: false,
          margin: { t: 30, b: 90, l: 40, r: 20 },
          xaxis: { tickangle: -40, gridcolor: gridC, tickfont: { size: 11 }, color: txtC },
          yaxis: { gridcolor: gridC, title: 'Score', color: txtC }
        }, { responsive: true });
      }
    } else {
      // Meeting View Charts (Only render if active tab is chTab)
      if (activeTab === 'chTab') {
        const tc = {
          text: dark ? '#f1f5f9' : '#2d2514',
          grid: dark ? '#2d3748' : 'rgba(146,108,21,.15)',
          pri: dark ? '#d5b263' : '#926C15',
          sec: dark ? '#38bdf8' : '#c79830'
        };

        const meetingLayout = {
          paper_bgcolor: 'rgba(0,0,0,0)',
          plot_bgcolor: 'rgba(0,0,0,0)',
          font: { family: 'Inter,sans-serif', color: tc.text },
          margin: { t: 40, b: 40, l: 40, r: 20 },
          xaxis: { gridcolor: tc.grid, zerolinecolor: tc.grid },
          yaxis: { gridcolor: tc.grid, zerolinecolor: tc.grid }
        };

        const chAttEl = document.getElementById('chAtt');
        const chScoreEl = document.getElementById('chScore');
        const chTopEl = document.getElementById('chTop');

        if (!members.length) {
          const emptyLayout = {
            ...meetingLayout,
            annotations: [{ text: 'No data. Upload spreadsheets.', showarrow: false, font: { size: 16 } }]
          };
          if (chAttEl) window.Plotly.newPlot(chAttEl, [], emptyLayout, { responsive: true });
          if (chScoreEl) window.Plotly.newPlot(chScoreEl, [], emptyLayout, { responsive: true });
          if (chTopEl) window.Plotly.newPlot(chTopEl, [], emptyLayout, { responsive: true });
          return;
        }

        // 1. Attendance Split
        const att = { P: 0, L: 0, S: 0, A: 0, U: 0 };
        members.forEach(x => {
          const k = att[x.attendance] !== undefined ? x.attendance : 'U';
          att[k] = (att[k] || 0) + 1;
        });
        if (chAttEl) {
          window.Plotly.newPlot(chAttEl, [{
            values: [att.P, att.L, att.S, att.A, att.U],
            labels: ['Present', 'Late', 'Substitute', 'Absent', 'Not Reported'],
            type: 'pie',
            marker: { colors: ['#10b981', '#fbbf24', '#38bdf8', '#ef4444', '#64748b'] },
            hole: 0.4
          }], { ...meetingLayout, showlegend: true }, { responsive: true });
        }

        // 2. Score Distribution
        if (chScoreEl) {
          window.Plotly.newPlot(chScoreEl, [{
            x: members.map(x => x.score),
            type: 'histogram',
            nbinsx: 10,
            marker: { color: tc.pri }
          }], {
            ...meetingLayout,
            xaxis: { ...meetingLayout.xaxis, title: 'Score' },
            yaxis: { ...meetingLayout.yaxis, title: 'Count' }
          }, { responsive: true });
        }

        // 3. Top 10 Contributors
        if (chTopEl) {
          const top10 = [...members].sort((a, b) => b.score - a.score).slice(0, 10);
          window.Plotly.newPlot(chTopEl, [{
            x: top10.map(x => x.name),
            y: top10.map(x => x.score),
            type: 'bar',
            marker: { color: tc.sec }
          }], meetingLayout, { responsive: true });
        }
      }
    }
  }, [cur, theme, activeTab, members.length]);

  // FILE UPLOAD HANDLING
  const triggerUpload = () => {
    if (cur === 'overall') {
      toast('Select a specific meeting (1–24) before uploading.', 'warning');
      return;
    }
    fileInputRef.current.click();
  };

  const handleFiles = (files) => {
    if (!files.length || cur === 'overall') return;
    toast(`Processing ${files.length} file(s)...`, 'info');

    Array.from(files).forEach(f => {
      const ext = f.name.split('.').pop().toLowerCase();
      const reader = new FileReader();

      if (ext === 'csv') {
        Papa.parse(f, {
          header: false,
          skipEmptyLines: true,
          complete: (results) => {
            uploadRawRows(f.name, results.data);
          },
          error: () => toast('Error reading CSV: ' + f.name, 'error')
        });
      } else if (ext === 'xlsx' || ext === 'xls') {
        reader.onload = (e) => {
          try {
            const d = new Uint8Array(e.target.result);
            const wb = XLSX.read(d, { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const jsonRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            uploadRawRows(f.name, jsonRows);
          } catch (err) {
            toast('Excel parse error: ' + f.name, 'error');
          }
        };
        reader.readAsArrayBuffer(f);
      } else {
        toast('Unsupported file type: ' + f.name, 'warning');
      }
    });

    fileInputRef.current.value = '';
  };

  const uploadRawRows = async (fileName, rawRows) => {
    try {
      const res = await fetch('/api/upload-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ meetingId: cur, fileName, rawRows })
      });
      if (res.ok) {
        const body = await res.json();
        setData(body.state.data || {});
        toast(`Successfully loaded "${fileName}" via AI Parser`, 'success');
      } else {
        const errData = await res.json();
        toast(errData.error || 'Server upload error.', 'error');
      }
    } catch (err) {
      console.error('Error uploading sheet to server:', err);
      toast('Network error uploading file.', 'error');
    }
  };

  // CATEGORY UPDATES AND DELETES
  const updateCategory = async (sheetId, newCat) => {
    try {
      const res = await fetch('/api/update-sheet-category', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ meetingId: cur, sheetId, category: newCat })
      });
      if (res.ok) {
        const body = await res.json();
        setData(body.state.data || {});
        toast('Category updated successfully', 'success');
      }
    } catch (err) {
      toast('Failed to update category', 'error');
    }
  };

  const deleteSheet = async (sheetId, sheetMeetingId) => {
    const targetMeeting = cur === 'overall' ? sheetMeetingId : cur;
    try {
      const res = await fetch('/api/delete-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ meetingId: targetMeeting, sheetId })
      });
      if (res.ok) {
        const body = await res.json();
        setData(body.state.data || {});
        toast('Sheet deleted', 'info');
      }
    } catch (err) {
      toast('Failed to delete sheet', 'error');
    }
  };

  const resetCurrent = async () => {
    if (cur === 'overall') return;
    if (window.confirm(`Clear all data for Meeting ${cur.replace('meeting_', '')}?`)) {
      try {
        const res = await fetch('/api/reset-meeting', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ meetingId: cur })
        });
        if (res.ok) {
          const body = await res.json();
          setData(body.state.data || {});
          toast('Meeting cleared', 'warning');
        }
      } catch (err) {
        toast('Failed to reset meeting', 'error');
      }
    }
  };

  const resetAll = async () => {
    if (window.confirm('Clear ALL data? This cannot be undone.')) {
      try {
        const res = await fetch('/api/reset-all', { method: 'POST' });
        if (res.ok) {
          const body = await res.json();
          setData(body.state.data || {});
          toast('All data reset', 'warning');
        }
      } catch (err) {
        toast('Failed to reset all data', 'error');
      }
    }
  };

  // EXPORT FUNCTIONS
  const exportOverallData = (fmt) => {
    if (!members.length) {
      toast('No data to export.', 'warning');
      return;
    }
    const hdrs = ['Rank', 'Member', 'Attendance', 'KYT', 'Referrals', 'Business', 'Inductions', 'Visitors', 'Testimonials', 'Score', 'Tier'];
    const sorted = [...members].sort((a, b) => b.score - a.score);
    const rows = sorted.map((m, i) => [
      i + 1,
      m.name,
      m.attendance || '-',
      m.kyt,
      m.referrals,
      m.business,
      m.inductions,
      m.visitors,
      m.testimonials,
      m.score,
      m.score >= 60 ? 'Green' : m.score >= 30 ? 'Orange' : 'Red'
    ]);

    if (fmt === 'csv') {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([Papa.unparse({ fields: hdrs, data: rows })], { type: 'text/csv' }));
      a.download = 'bni_overall.csv';
      a.click();
      toast('CSV exported', 'success');
    } else {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([hdrs, ...rows]), 'Overall');
      XLSX.writeFile(wb, 'bni_overall.xlsx');
      toast('Excel exported', 'success');
    }
  };

  const exportMeetingData = (fmt) => {
    if (!members.length) {
      toast('No data to export.', 'warning');
      return;
    }
    const hdrs = ['Member', 'KYT', 'Attendance', 'Referrals', 'Business', 'Inductions', 'Visitors', 'Testimonials', 'Score'];
    const rows = members.map(m => [
      m.name,
      m.kyt,
      m.attendance || '-',
      m.referrals,
      m.business,
      m.inductions,
      m.visitors,
      m.testimonials,
      m.score
    ]);

    if (fmt === 'csv') {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([Papa.unparse({ fields: hdrs, data: rows })], { type: 'text/csv' }));
      a.download = 'bni_meeting.csv';
      a.click();
      toast('CSV exported', 'success');
    } else {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([hdrs, ...rows]), 'Performance');
      XLSX.writeFile(wb, 'bni_meeting.xlsx');
      toast('Excel exported', 'success');
    }
  };

  // TABLE MANIPULATION (SORT & FILTER)
  const handleSort = (key) => {
    if (sb === key) {
      setSo(prev => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSb(key);
      setSo('desc');
    }
  };

  const clearFilters = () => {
    setTsrch('');
    setCf({});
    toast('Filters cleared', 'info');
  };

  // Process the Master Table rows based on Search and Column Filters
  let filteredMembers = [...members];
  if (tsrch) {
    filteredMembers = filteredMembers.filter(m => m.name.toLowerCase().includes(tsrch.toLowerCase()));
  }
  Object.keys(cf).forEach(k => {
    const v = cf[k].toLowerCase().trim();
    if (v) {
      filteredMembers = filteredMembers.filter(m =>
        String(m[k] !== undefined ? m[k] : '')
          .toLowerCase()
          .includes(v)
      );
    }
  });

  // Sort
  filteredMembers.sort((a, b) => {
    let va = a[sb], vb = b[sb];
    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();
    return va < vb ? (so === 'asc' ? -1 : 1) : va > vb ? (so === 'asc' ? 1 : -1) : 0;
  });

  const totalRecords = filteredMembers.length;
  const totalPages = Math.ceil(totalRecords / ps) || 1;
  const paginatedMembers = filteredMembers.slice((cp - 1) * ps, cp * ps);

  // Column resize implementation in React
  const startResize = (e, key) => {
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = cw[key] || e.target.parentElement.getBoundingClientRect().width;

    const onMouseMove = (moveEvent) => {
      const nw = Math.max(80, startWidth + (moveEvent.clientX - startX));
      setCw(prev => ({ ...prev, [key]: nw }));
    };

    const onMouseUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  };

  // ACTIVE DATA SHEETS LISTING
  let activeFiles = [];
  if (cur === 'overall') {
    Object.entries(data).forEach(([mk, mf]) => {
      if (mf && mf.length) {
        activeFiles.push(...mf.map(f => ({ ...f, _mk: mk })));
      }
    });
  } else {
    activeFiles = data[cur] || [];
  }

  // AGGREGATE SUMMARY NUMBERS FOR OVERALL HERO BANNER
  let overallPresent = 0, overallAbsent = 0, overallLate = 0, overallSub = 0;
  let overallRef = 0, overallBiz = 0, overallInd = 0, overallVis = 0, overallTest = 0, overallKyt = 0;
  let tierGreen = 0, tierOrange = 0, tierRed = 0;

  members.forEach(m => {
    if (m.attendance === 'P') overallPresent++;
    else if (m.attendance === 'A') overallAbsent++;
    else if (m.attendance === 'L') overallLate++;
    else if (m.attendance === 'S') overallSub++;

    overallRef += m.referrals;
    overallBiz += m.business;
    overallInd += m.inductions;
    overallVis += m.visitors;
    overallTest += m.testimonials;

    const ks = String(m.kyt).toLowerCase();
    if ((typeof m.kyt === 'number' && m.kyt > 0) || /yes|done/.test(ks) || ks === '1') overallKyt++;

    if (m.score >= 60) tierGreen++;
    else if (m.score >= 30) tierOrange++;
    else tierRed++;
  });

  const avgScore = members.length
    ? Math.round(members.reduce((a, m) => a + m.score, 0) / members.length)
    : 0;
  const topMember = members.length
    ? [...members].sort((a, b) => b.score - a.score)[0]
    : null;
  const attRate = members.length
    ? Math.round(overallPresent / members.length * 100)
    : 0;

  // MEMBER MEETINGS MAP FOR CONSOLIDATED OVERALL TABLE
  const memberMeetings = {};
  if (cur === 'overall') {
    Object.keys(data).forEach(mk => {
      (data[mk] || []).forEach(file => {
        file.rows.forEach(row => {
          const mn = bestMatch(row.rawName, members.map(m => m.name));
          if (mn) {
            if (!memberMeetings[mn]) memberMeetings[mn] = new Set();
            memberMeetings[mn].add(mk);
          }
        });
      });
    });
  }

  // Filter Overall Table
  let filteredOverallMembers = [...members];
  if (ovSearch) {
    filteredOverallMembers = filteredOverallMembers.filter(m =>
      m.name.toLowerCase().includes(ovSearch.toLowerCase())
    );
  }

  const meetingKeysWithData = Object.keys(data).filter(k => data[k] && data[k].length);
  const totalMeetingsCompleted = meetingKeysWithData.length;

  return (
    <>
      {/* Navbar */}
      <header className="navbar">
        <div className="nav-brand">
          <div className="nav-logo" style={{ background: 'none', width: 'auto', height: '38px' }}>
            <img 
              src="https://bbb-india.com/wp-content/uploads/elementor/thumbs/logo-pwxbt60i7amqn8dr90567800axjmjj76cretv1pzpg.jpg" 
              alt="BNI Logo" 
              style={{ height: '100%', objectFit: 'contain', borderRadius: 'var(--rsm)' }} 
            />
          </div>
          <div className="nav-title">Member Performance Dashboard</div>
        </div>
        <div className="nav-actions">
          {cur !== 'overall' && (
            <button className="btn btn-dng btn-sm" onClick={resetCurrent}>
              <i className="fa-solid fa-eraser"></i> Clear Meeting
            </button>
          )}
          <button className="btn btn-dng btn-sm" onClick={resetAll}>
            <i className="fa-solid fa-trash-can"></i> Reset All
          </button>
          <button className="ibtn" onClick={toggleTheme} title="Toggle theme">
            <i className={theme === 'dark' ? 'fa-solid fa-sun' : 'fa-solid fa-moon'}></i>
          </button>
        </div>
      </header>

      {/* Meeting Selector Bar */}
      <div className="mbar">
        <div className="mbar-label">
          <i className="fa-solid fa-calendar-days" style={{ color: 'var(--pri)' }}></i> Select View:
        </div>
        <select className="msel" value={cur} onChange={(e) => { setCur(e.target.value); setCp(1); }}>
          <option value="overall">⭐ Overall Meeting Summary (Till Date)</option>
          <optgroup label="Financial Year Meetings">
            {Array.from({ length: TOTAL_MEETINGS }).map((_, idx) => (
              <option key={idx} value={`meeting_${idx + 1}`}>
                Meeting {idx + 1}
              </option>
            ))}
          </optgroup>
        </select>
        <div className="mbadge">
          <i className="fa-solid fa-circle-dot"></i>
          <span>{cur === 'overall' ? 'Overall View' : `Meeting ${cur.replace('meeting_', '')}`}</span>
        </div>
        <div className="minfo">
          <span className={activeFiles.length ? 'mdot' : 'mdot empty'}></span>
          <span>
            {activeFiles.length
              ? `${activeFiles.length} sheet(s) loaded`
              : cur === 'overall'
              ? 'Select a meeting to upload'
              : 'No data — upload files'}
          </span>
        </div>
      </div>

      {/* Dashboard Main Grid */}
      <main className="dash">
        {/* Sidebar */}
        <aside className="sidebar">
          {/* Upload Center */}
          <section className="gc" style={{ padding: '20px' }}>
            <h2 className="stitle">
              <i className="fa-solid fa-cloud-arrow-up" style={{ color: 'var(--pri)' }}></i> Upload Center
            </h2>
            <div
              className={`dz ${isDragOver ? 'over' : ''}`}
              onClick={triggerUpload}
              onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
              onDragLeave={() => setIsDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setIsDragOver(false);
                if (cur === 'overall') {
                  toast('Select a meeting first.', 'warning');
                  return;
                }
                if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
              }}
            >
              <i className="fa-solid fa-file-arrow-up dz-icon"></i>
              <span style={{ fontSize: '.88rem', fontWeight: 500, color: 'var(--muted)' }}>
                Drag & Drop or click to browse
              </span>
              <span style={{ fontSize: '.75rem', color: 'var(--muted)' }}>CSV, XLS, XLSX</span>
              <span className="dz-tag">
                <i className="fa-solid fa-arrow-right"></i> {cur === 'overall' ? 'Overall View' : `Meeting ${cur.replace('meeting_', '')}`}
              </span>
              <input
                type="file"
                ref={fileInputRef}
                className="file-input"
                accept=".csv,.xls,.xlsx"
                multiple
                onChange={(e) => handleFiles(e.target.files)}
              />
            </div>
            <p style={{ fontSize: '.73rem', color: 'var(--muted)', marginTop: '10px', textAlign: 'center' }}>
              Files go into the <strong>currently selected meeting</strong>.
            </p>
          </section>

          {/* Active Data Sheets */}
          <section className="gc" style={{ padding: '20px' }}>
            <h2 className="stitle">
              <i className="fa-solid fa-folder-open" style={{ color: 'var(--sec)' }}></i> Active Data Sheets
            </h2>
            <div className="flist">
              {activeFiles.length === 0 ? (
                <div style={{ textAlign: 'center', color: 'var(--muted)', fontSize: '.8rem', padding: '12px 0' }}>
                  No sheets uploaded yet.
                </div>
              ) : (
                activeFiles.map(file => (
                  <div key={file.id} className="fc">
                    <div className="fchdr">
                      <span className="fcname" title={file.name}>{file.name}</span>
                      {cur === 'overall' && (
                        <span className="fcm">M{file._mk.replace('meeting_', '')}</span>
                      )}
                      <span className="fccnt">{file.count} rows</span>
                      <button
                        className="ibtn btn-sm"
                        style={{ border: 'none', color: 'var(--danger)', background: 'transparent' }}
                        onClick={() => deleteSheet(file.id, file._mk)}
                        title="Delete Sheet"
                      >
                        <i className="fa-solid fa-trash-can"></i>
                      </button>
                    </div>
                    <select
                      className="catsel"
                      value={file.category}
                      onChange={(e) => updateCategory(file.id, e.target.value)}
                    >
                      {Object.entries(CATS).map(([k, label]) => (
                        <option key={k} value={k}>{label}</option>
                      ))}
                    </select>
                  </div>
                ))
              )}
            </div>
          </section>
        </aside>

        {/* Main Content Area */}
        <div className="maincontent">
          {cur === 'overall' ? (
            /* ===== OVERALL VIEW PAGE ===== */
            <div id="overallPage" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              {/* Hero Banner */}
              <div className="ov-banner">
                <div className="ov-banner-top">
                  <div className="ov-banner-title">
                    <i className="fa-solid fa-layer-group"></i> Overall Meeting Summary
                  </div>
                  <div className="ov-banner-sub">
                    {totalMeetingsCompleted === 0
                      ? 'No meetings have data yet. Select a meeting (1–24) and upload files.'
                      : `Financial year summary · ${totalMeetingsCompleted} meeting${totalMeetingsCompleted > 1 ? 's' : ''} completed · ${members.length} unique member${members.length !== 1 ? 's' : ''}`}
                  </div>
                  {totalMeetingsCompleted > 0 && (
                    <div className="ov-banner-chips">
                      <span className="ov-chip"><i className="fa-solid fa-circle-check"></i> {tierGreen} green</span>
                      <span className="ov-chip"><i className="fa-solid fa-circle-exclamation"></i> {tierOrange} orange</span>
                      <span className="ov-chip"><i className="fa-solid fa-circle-xmark"></i> {tierRed} red</span>
                      <span className="ov-chip">
                        <i className="fa-solid fa-trophy"></i> Top: {topMember ? `${topMember.name.split(' ')[0]} (${topMember.score}pts)` : '—'}
                      </span>
                      <span className="ov-chip"><i className="fa-solid fa-percent"></i> {attRate}% attendance rate</span>
                    </div>
                  )}
                </div>
                <div className="ov-stats-grid">
                  {[
                    { icon: 'fa-calendar-days', val: totalMeetingsCompleted, lbl: 'Meetings done', sub: `of ${TOTAL_MEETINGS} planned`, color: 'var(--pri)', pct: Math.round((totalMeetingsCompleted / TOTAL_MEETINGS) * 100) },
                    { icon: 'fa-users', val: members.length, lbl: 'Unique members', sub: 'across all meetings', color: 'var(--sec)', pct: 100 },
                    { icon: 'fa-user-check', val: overallPresent, lbl: 'Total present', sub: `${attRate}% rate · ${overallAbsent} absent`, color: 'var(--ok)', pct: attRate },
                    { icon: 'fa-clock-rotate-left', val: overallLate, lbl: 'Late arrivals', sub: `${overallSub} substitutes`, color: 'var(--warn)', pct: members.length ? Math.round((overallLate / members.length) * 100) : 0 },
                    { icon: 'fa-handshake', val: overallRef, lbl: 'Total referrals', sub: 'given across meetings', color: 'var(--sec)', pct: Math.min(100, overallRef * 5) },
                    { icon: 'fa-sack-dollar', val: `Rs. ${overallBiz.toLocaleString()}`, lbl: 'Business given', sub: 'cumulative value', color: 'var(--gold)', pct: 100 },
                    { icon: 'fa-user-plus', val: overallInd, lbl: 'Inductions', sub: 'new members inducted', color: 'var(--orange)', pct: Math.min(100, overallInd * 10) },
                    { icon: 'fa-eye', val: overallVis, lbl: 'Visitors', sub: 'brought to meetings', color: 'var(--warn)', pct: Math.min(100, overallVis * 5) },
                    { icon: 'fa-star', val: overallTest, lbl: 'Testimonials', sub: 'total given', color: 'var(--danger)', pct: Math.min(100, overallTest * 5) },
                    { icon: 'fa-brain', val: overallKyt, lbl: 'KYT done', sub: members.length ? `${Math.round((overallKyt / members.length) * 100)}% completion` : '', color: 'var(--ok)', pct: members.length ? Math.round((overallKyt / members.length) * 100) : 0 },
                    { icon: 'fa-chart-line', val: `${avgScore}pts`, lbl: 'Avg score', sub: topMember ? `Best: ${topMember.score}pts` : '', color: 'var(--pri)', pct: Math.min(100, avgScore) },
                  ].map((s, idx) => (
                    <div key={idx} className="ov-scard">
                      <i className={`fa-solid ${s.icon} ov-scard-icon`} style={{ color: s.color }}></i>
                      <div className="ov-scard-val" style={{ color: s.color }}>{s.val}</div>
                      <div className="ov-scard-lbl">{s.lbl}</div>
                      <div className="ov-scard-sub">{s.sub}</div>
                      <div className="ov-scard-bar">
                        <div className="ov-scard-fill" style={{ width: `${Math.max(0, Math.min(100, s.pct))}%`, background: s.color }}></div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Tiers summary */}
              <div className="ov-section">
                <div className="ov-sec-hdr">
                  <div className="ov-sec-title">
                    <i className="fa-solid fa-chart-simple" style={{ color: 'var(--pri)' }}></i> Member score tiers
                  </div>
                  <span style={{ fontSize: '.78rem', color: 'var(--muted)' }}>Green ≥ 60 · Orange 30–59 · Red &lt; 30</span>
                </div>
                <div className="ov-tier-grid">
                  {[
                    { title: 'Green tier', range: 'Score ≥ 60 pts', color: 'var(--ok)', icon: 'fa-circle-check', class: 'green', members: [...members].filter(m => m.score >= 60) },
                    { title: 'Orange tier', range: 'Score 30–59 pts', color: 'var(--orange)', icon: 'fa-circle-exclamation', class: 'orange', members: [...members].filter(m => m.score >= 30 && m.score < 60) },
                    { title: 'Red tier', range: 'Score < 30 pts', color: 'var(--danger)', icon: 'fa-circle-xmark', class: 'red', members: [...members].filter(m => m.score < 30) }
                  ].map((tier, tidx) => {
                    const maxScore = members.length ? Math.max(...members.map(m => m.score), 1) : 1;
                    return (
                      <div key={tidx} className={`ov-tier ${tier.class}`}>
                        <div className="ov-tier-head">
                          <div>
                            <div className="ov-tier-label">{tier.title}</div>
                            <div className="ov-tier-range" style={{ color: tier.color }}>{tier.range}</div>
                          </div>
                          <div className="ov-tier-meta">
                            <div className="ov-tier-count">{tier.members.length}</div>
                            <div className="ov-tier-label">members</div>
                          </div>
                          <div className="ov-tier-icon"><i className={`fa-solid ${tier.icon}`}></i></div>
                        </div>
                        <div className="ov-tier-body">
                          <div className="ov-tier-list">
                            {tier.members.length === 0 ? (
                              <div style={{ fontSize: '.82rem', color: 'var(--muted)', padding: '16px 0', textAlign: 'center' }}>
                                No members in this tier
                              </div>
                            ) : (
                              tier.members.map((m, midx) => (
                                <React.Fragment key={midx}>
                                  <div className="ov-tier-row2">
                                    <span className="ov-tier-rank">{midx + 1}</span>
                                    <span className="ov-tier-mname" title={m.name}>{m.name}</span>
                                    <span className="ov-tier-mscore">{m.score}pts</span>
                                  </div>
                                  <div className="ov-tier-progress">
                                    <div className="ov-tier-prog-fill" style={{ width: `${Math.round((m.score / maxScore) * 100)}%` }}></div>
                                  </div>
                                </React.Fragment>
                              ))
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Visual analytics */}
              <div className="ov-section">
                <div className="ov-sec-hdr">
                  <div className="ov-sec-title">
                    <i className="fa-solid fa-chart-pie" style={{ color: 'var(--pri)' }}></i> Visual analytics
                  </div>
                </div>
                <div className="ov-chart-grid">
                  <div className="gc ov-ch">
                    <div className="ov-ch-title"><i className="fa-solid fa-circle-half-stroke"></i> Score tier split</div>
                    <div className="ov-ch-body" id="ov-ch-pie"></div>
                  </div>
                  <div className="gc ov-ch">
                    <div className="ov-ch-title"><i className="fa-solid fa-bars-progress"></i> Attendance breakdown</div>
                    <div className="ov-ch-body" id="ov-ch-att"></div>
                  </div>
                  <div className="gc ov-ch ov-ch-full">
                    <div className="ov-ch-title"><i className="fa-solid fa-ranking-star"></i> Top 15 members — cumulative score</div>
                    <div className="ov-ch-body" id="ov-ch-top"></div>
                  </div>
                </div>
              </div>

              {/* Consolidated table */}
              <div className="ov-section">
                <div className="ov-sec-hdr" style={{ marginBottom: '14px' }}>
                  <div className="ov-sec-title">
                    <i className="fa-solid fa-users" style={{ color: 'var(--pri)' }}></i> All members — consolidated
                  </div>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <div className="srchwrap" style={{ minWidth: '220px' }}>
                      <i className="fa-solid fa-magnifying-glass srchico"></i>
                      <input
                        type="text"
                        className="srchinp"
                        placeholder="Search member…"
                        value={ovSearch}
                        onChange={(e) => setOvSearch(e.target.value)}
                      />
                    </div>
                    <button className="btn btn-sec btn-sm" onClick={() => exportOverallData('csv')}>
                      <i className="fa-solid fa-file-csv"></i> Export CSV
                    </button>
                    <button className="btn btn-sec btn-sm" onClick={() => exportOverallData('excel')}>
                      <i className="fa-solid fa-file-excel"></i> Excel
                    </button>
                  </div>
                </div>
                <div className="ov-tbl-wrap">
                  <table className="ov-tbl">
                    <thead>
                      <tr>
                        <th style={{ width: '40px' }}>#</th>
                        <th>Member name</th>
                        <th>Meetings attended</th>
                        <th>Attendance</th>
                        <th>KYT</th>
                        <th>Referrals</th>
                        <th>Business given</th>
                        <th>Inductions</th>
                        <th>Visitors</th>
                        <th>Testimonials</th>
                        <th>Total score</th>
                        <th>Tier</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredOverallMembers.length === 0 ? (
                        <tr>
                          <td colSpan={12} style={{ textAlign: 'center', color: 'var(--muted)', padding: '24px' }}>
                            No members found.
                          </td>
                        </tr>
                      ) : (
                        filteredOverallMembers.map((m, idx) => {
                          const tierPill = m.score >= 60
                            ? <span className="pill pill-green">Green</span>
                            : m.score >= 30
                            ? <span className="pill pill-orange">Orange</span>
                            : <span className="pill pill-red">Red</span>;

                          const scoreClass = m.score >= 60 ? 'shi' : m.score >= 30 ? 'smd' : 'slo';
                          const initials = m.name.split(' ').map(w => w[0] || '').slice(0, 2).join('').toUpperCase();
                          const meetCount = memberMeetings[m.name] ? memberMeetings[m.name].size : 0;

                          return (
                            <tr key={idx}>
                              <td style={{ color: 'var(--muted)', fontWeight: 700, fontSize: '.8rem' }}>{idx + 1}</td>
                              <td>
                                <div className="ov-name-cell">
                                  <span className="ov-avatar">{initials}</span>
                                  <strong>{m.name}</strong>
                                </div>
                              </td>
                              <td>
                                <span style={{ fontSize: '.8rem', fontWeight: 600 }}>
                                  {meetCount} meeting{meetCount !== 1 ? 's' : ''}
                                </span>
                              </td>
                              <td>{m.attendance || '—'}</td>
                              <td>{m.kyt}</td>
                              <td>{m.referrals}</td>
                              <td style={{ fontWeight: 600 }}>Rs. {m.business.toLocaleString()}</td>
                              <td>{m.inductions}</td>
                              <td>{m.visitors}</td>
                              <td>{m.testimonials}</td>
                              <td><span className={`sbadge ${scoreClass}`}>{m.score}</span></td>
                              <td>{tierPill}</td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Meetings at a glance card grid */}
              <div className="ov-section">
                <div className="ov-sec-hdr" style={{ marginBottom: '14px' }}>
                  <div className="ov-sec-title">
                    <i className="fa-solid fa-calendar-check" style={{ color: 'var(--pri)' }}></i> Meetings at a glance
                  </div>
                  <span style={{ fontSize: '.78rem', color: 'var(--muted)' }}>Click any meeting to view its details</span>
                </div>
                <div className="ov-mgrid">
                  {Array.from({ length: TOTAL_MEETINGS }).map((_, idx) => {
                    const mk = `meeting_${idx + 1}`;
                    const meetingFiles = data[mk] || [];
                    const hasData = meetingFiles.length > 0;
                    const meetingMembers = hasData ? consolidate(meetingFiles) : [];
                    const greenCount = meetingMembers.filter(m => m.score >= 60).length;
                    const orangeCount = meetingMembers.filter(m => m.score >= 30 && m.score < 60).length;
                    const redCount = meetingMembers.filter(m => m.score < 30).length;
                    const topSc = meetingMembers.length ? Math.max(...meetingMembers.map(m => m.score)) : 0;
                    const greenPct = hasData && meetingMembers.length ? Math.round((greenCount / meetingMembers.length) * 100) : 0;

                    return (
                      <div
                        key={idx}
                        className={`ov-mc ${hasData ? '' : 'empty'}`}
                        onClick={hasData ? () => { setCur(mk); setCp(1); } : undefined}
                        style={{ cursor: hasData ? 'pointer' : 'default' }}
                      >
                        <span className="ov-mc-dot"></span>
                        <div className="ov-mc-num">M{idx + 1}</div>
                        <div className="ov-mc-sub">Meeting {idx + 1}</div>
                        <div className="ov-mc-cnt">
                          {hasData ? `${meetingMembers.length} member${meetingMembers.length !== 1 ? 's' : ''}` : 'No data yet'}
                        </div>
                        {hasData && (
                          <>
                            <div className="ov-mc-pills">
                              <span className="pill pill-green">{greenCount} <i className="fa-solid fa-circle" style={{ fontSize: '.45rem' }}></i></span>
                              <span className="pill pill-orange">{orangeCount} <i className="fa-solid fa-circle" style={{ fontSize: '.45rem' }}></i></span>
                              <span className="pill pill-red">{redCount} <i className="fa-solid fa-circle" style={{ fontSize: '.45rem' }}></i></span>
                            </div>
                            <div style={{ fontSize: '.7rem', color: 'var(--muted)' }}>Top: {topSc}pts</div>
                            <div className="ov-mc-bar">
                              <div className="ov-mc-barfill" style={{ width: `${greenPct}%` }}></div>
                            </div>
                          </>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          ) : (
            /* ===== MEETING PAGE (non-overall) ===== */
            <div id="meetingPage" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              {/* KPI cards row */}
              <div className="kpirow">
                {(() => {
                  if (!members.length) {
                    return (
                      <>
                        <div className="gc kpicard"><span className="kpitit">Total Members</span><span className="kpival">0</span></div>
                        <div className="gc kpicard"><span className="kpitit">Green (60+)</span><span className="kpival">0</span></div>
                        <div className="gc kpicard"><span className="kpitit">Attendance</span><span className="kpival">0</span></div>
                        <div className="gc kpicard"><span className="kpitit">Red (&lt;30)</span><span className="kpival">0</span></div>
                      </>
                    );
                  }
                  let pres = 0, abs = 0, grn = 0, rd = 0;
                  members.forEach(x => {
                    if (x.score >= 60) grn++;
                    if (x.score < 30) rd++;
                    if (x.attendance === 'P') pres++;
                    if (x.attendance === 'A') abs++;
                  });
                  return (
                    <>
                      <div className="gc kpicard">
                        <span className="kpitit">Total Members</span>
                        <span className="kpival">{members.length}</span>
                        <span className="kpifot"><i className="fa-solid fa-users"></i> Master list</span>
                      </div>
                      <div className="gc kpicard">
                        <span className="kpitit">Green (60+)</span>
                        <span className="kpival" style={{ color: 'var(--ok)' }}>{grn}</span>
                        <span className="kpifot"><i className="fa-solid fa-circle" style={{ color: 'var(--ok)' }}></i> Score ≥ 60</span>
                      </div>
                      <div className="gc kpicard">
                        <span className="kpitit">Attendance</span>
                        <span className="kpival">{pres} <span style={{ fontSize: '1rem', color: 'var(--muted)' }}>/ {abs} abs</span></span>
                        <span className="kpifot"><i className="fa-solid fa-user-check"></i> Present / Absent</span>
                      </div>
                      <div className="gc kpicard">
                        <span className="kpitit">Red (&lt;30)</span>
                        <span className="kpival" style={{ color: 'var(--danger)' }}>{rd}</span>
                        <span className="kpifot"><i className="fa-solid fa-circle" style={{ color: 'var(--danger)' }}></i> Score &lt; 30</span>
                      </div>
                    </>
                  );
                })()}
              </div>

              {/* Tab Navigation */}
              <nav className="tabbar">
                <div
                  className={`tabit ${activeTab === 'lbTab' ? 'active' : ''}`}
                  onClick={() => setActiveTab('lbTab')}
                >
                  <i className="fa-solid fa-ranking-star"></i> Leaderboard
                </div>
                <div
                  className={`tabit ${activeTab === 'tblTab' ? 'active' : ''}`}
                  onClick={() => setActiveTab('tblTab')}
                >
                  <i className="fa-solid fa-table-list"></i> Master Table
                </div>
                <div
                  className={`tabit ${activeTab === 'chTab' ? 'active' : ''}`}
                  onClick={() => setActiveTab('chTab')}
                >
                  <i className="fa-solid fa-chart-pie"></i> Visual Analytics
                </div>
              </nav>

              {/* Leaderboard Panel */}
              {activeTab === 'lbTab' && (
                <div className="tabpanel active">
                  <div className="lbgrid">
                    <div className="gc lbcard">
                      <div className="lbhdr">
                        <h3 className="stitle" style={{ marginBottom: 0 }}>
                          <i className="fa-solid fa-medal" style={{ color: '#fcd34d' }}></i> Rankings
                        </h3>
                        <span style={{ fontSize: '.75rem', color: 'var(--muted)' }}>By total score</span>
                      </div>
                      <div className="lblist">
                        {members.length === 0 ? (
                          <div style={{ textAlign: 'center', color: 'var(--muted)', padding: '32px 0' }}>
                            Upload spreadsheets to populate.
                          </div>
                        ) : (
                          [...members]
                            .sort((a, b) => b.score - a.score)
                            .map((x, i) => {
                              const rank = i + 1;
                              const rankClass = rank === 1 ? 'rk1' : rank === 2 ? 'rk2' : rank === 3 ? 'rk3' : 'rkn';
                              const rankIcon = rank === 1 ? <i className="fa-solid fa-trophy"></i>
                                : rank === 2 ? <i className="fa-solid fa-award"></i>
                                : rank === 3 ? <i className="fa-solid fa-medal"></i>
                                : rank;

                              return (
                                <div key={i} className="lbit">
                                  <div className={`rankbadge ${rankClass}`}>{rankIcon}</div>
                                  <div className="lbdet">
                                    <span className="lbname">{x.name}</span>
                                    <span className="lbsc">{x.score} pts</span>
                                  </div>
                                </div>
                              );
                            })
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Master Table Panel */}
              {activeTab === 'tblTab' && (
                <div className="tabpanel active">
                  <div className="tctrl">
                    <div className="srchwrap">
                      <i className="fa-solid fa-magnifying-glass srchico"></i>
                      <input
                        type="text"
                        className="srchinp"
                        placeholder="Search members..."
                        value={tsrch}
                        onChange={(e) => { setTsrch(e.target.value); setCp(1); }}
                      />
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button className="btn btn-sec" onClick={clearFilters}>
                        <i className="fa-solid fa-filter-circle-xmark"></i> Clear
                      </button>
                      <button className="btn btn-sec" onClick={() => exportMeetingData('csv')}>
                        <i className="fa-solid fa-file-csv"></i> CSV
                      </button>
                      <button className="btn btn-sec" onClick={() => exportMeetingData('excel')}>
                        <i className="fa-solid fa-file-excel"></i> Excel
                      </button>
                      <button className="btn btn-sec" onClick={() => window.print()}>
                        <i className="fa-solid fa-print"></i> Print
                      </button>
                    </div>
                  </div>

                  <div className="tblwrap">
                    <table className="tbl">
                      <thead>
                        <tr>
                          {[
                            { key: 'name', label: 'Member' },
                            { key: 'kyt', label: 'KYT' },
                            { key: 'attendance', label: 'Attendance' },
                            { key: 'referrals', label: 'Referrals' },
                            { key: 'business', label: 'Business Given' },
                            { key: 'inductions', label: 'Inductions' },
                            { key: 'visitors', label: 'Visitors' },
                            { key: 'testimonials', label: 'Testimonials' },
                            { key: 'score', label: 'Total Score' }
                          ].map(col => {
                            const isSorted = sb === col.key;
                            const sortIcon = isSorted
                              ? (so === 'asc' ? 'fa-sort-up' : 'fa-sort-down')
                              : 'fa-sort';

                            return (
                              <th
                                key={col.key}
                                style={{ width: cw[col.key] ? `${cw[col.key]}px` : undefined }}
                              >
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                  <span
                                    onClick={() => handleSort(col.key)}
                                    style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px' }}
                                  >
                                    {col.label}
                                    <i className={`fa-solid ${sortIcon} sico`}></i>
                                  </span>
                                </div>
                                <div className="cfilt">
                                  <input
                                    type="text"
                                    className="cfinp"
                                    value={cf[col.key] || ''}
                                    placeholder="Filter..."
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setCf(prev => {
                                        const next = { ...prev };
                                        if (!val.trim()) delete next[col.key];
                                        else next[col.key] = val;
                                        return next;
                                      });
                                      setCp(1);
                                    }}
                                  />
                                </div>
                                <div className="resizer" onMouseDown={(e) => startResize(e, col.key)}></div>
                              </th>
                            );
                          })}
                        </tr>
                      </thead>
                      <tbody>
                        {paginatedMembers.length === 0 ? (
                          <tr>
                            <td colSpan={9} style={{ textAlign: 'center', color: 'var(--muted)', padding: '32px' }}>
                              No records found.
                            </td>
                          </tr>
                        ) : (
                          paginatedMembers.map((m, idx) => {
                            const scoreClass = m.score >= 60 ? 'shi' : m.score >= 30 ? 'smd' : 'slo';
                            return (
                              <tr key={idx}>
                                <td><strong>{m.name}</strong></td>
                                <td>{m.kyt}</td>
                                <td>{m.attendance || '-'}</td>
                                <td>{m.referrals}</td>
                                <td>Rs. {m.business.toLocaleString()}</td>
                                <td>{m.inductions}</td>
                                <td>{m.visitors}</td>
                                <td>{m.testimonials}</td>
                                <td><span className={`sbadge ${scoreClass}`}>{m.score}</span></td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>

                    <div className="pgbar">
                      <div>
                        Show{' '}
                        <select
                          value={ps}
                          onChange={(e) => { setPs(parseInt(e.target.value)); setCp(1); }}
                          style={{
                            minWidth: '60px',
                            padding: '4px 24px 4px 8px',
                            border: '1px solid var(--bdr)',
                            borderRadius: 'var(--rxs)',
                            background: 'var(--ibg)',
                            color: 'var(--txt)'
                          }}
                        >
                          <option value="10">10</option>
                          <option value="25">25</option>
                          <option value="50">50</option>
                          <option value="100">100</option>
                        </select>{' '}
                        per page
                      </div>
                      <div className="pgpgs">
                        <button className="ibtn" onClick={() => setCp(prev => Math.max(1, prev - 1))} disabled={cp === 1}>
                          <i className="fa-solid fa-chevron-left"></i>
                        </button>
                        <span>Page {cp} of {totalPages}</span>
                        <button className="ibtn" onClick={() => setCp(prev => Math.min(totalPages, prev + 1))} disabled={cp === totalPages}>
                          <i className="fa-solid fa-chevron-right"></i>
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Visual Analytics Panel */}
              {activeTab === 'chTab' && (
                <div className="tabpanel active">
                  <div className="chgrid">
                    <div className="gc chcard">
                      <h3 style={{ marginBottom: '12px' }}>Attendance breakdown</h3>
                      <div id="chAtt" className="chcont"></div>
                    </div>
                    <div className="gc chcard">
                      <h3 style={{ marginBottom: '12px' }}>Score distribution</h3>
                      <div id="chScore" className="chcont"></div>
                    </div>
                    <div className="gc chcard chfull">
                      <h3 style={{ marginBottom: '12px' }}>Top 10 contributors</h3>
                      <div id="chTop" className="chcont"></div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </main>

      {/* Footer */}
      <footer style={{
        textAlign: 'center',
        padding: '24px',
        borderTop: '1px solid var(--bdr)',
        background: 'var(--surf)',
        fontSize: '0.9rem',
        color: 'var(--muted)',
        marginTop: 'auto'
      }}>
        Powered by{' '}
        <a 
          href="https://automationlabs.online" 
          target="_blank" 
          rel="noopener noreferrer" 
          style={{ 
            color: 'var(--pri)', 
            textDecoration: 'none', 
            fontWeight: '600' 
          }}
          onMouseEnter={(e) => e.target.style.textDecoration = 'underline'}
          onMouseLeave={(e) => e.target.style.textDecoration = 'none'}
        >
          AI Automation Labs
        </a>
      </footer>

      {/* Toast Notifications container */}
      <div className="toasts" id="toasts">
        {toasts.map(t => {
          const typeClass = t.type === 'success' ? 'ok' : t.type === 'error' ? 'err' : t.type === 'warning' ? 'warn' : 'info';
          const iconClass = t.type === 'success' ? 'fa-circle-check'
            : t.type === 'error' ? 'fa-circle-exclamation'
            : t.type === 'warning' ? 'fa-triangle-exclamation'
            : 'fa-info-circle';

          return (
            <div key={t.id} className={`toast t-${typeClass}`}>
              <i className={`fa-solid ${iconClass}`}></i>
              <span>{t.msg}</span>
            </div>
          );
        })}
      </div>
    </>
  );
}
