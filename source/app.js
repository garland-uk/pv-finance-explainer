/* Garland UK | Solar PV asset finance explainer
   Reads a PV*SOL report (.docx or .pdf) entirely in the browser, takes two monthly
   repayments and lays out a print-ready A4 landscape explainer.
   Nothing is uploaded, stored or sent anywhere. */
(function () {
  'use strict';

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const $ = (s) => document.querySelector(s);

  // ------------------------------------------------------------------
  // 1. Text extraction. Both formats become a list of lines, each line a
  //    list of cells (table rows keep their columns).
  // ------------------------------------------------------------------

  function parseXml(str) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('The Word file could not be read.');
    return doc;
  }

  // Text boxes and legacy fallbacks sit inside runs; keep them out of the
  // paragraph's own text (they are read separately as their own lines).
  function insideSkipped(n, stop) {
    for (let p = n.parentNode; p && p !== stop; p = p.parentNode) {
      if (p.localName === 'txbxContent' || p.localName === 'Fallback') return true;
    }
    return false;
  }
  function textBoxLines(el) {
    const out = [];
    for (const tb of Array.from(el.getElementsByTagNameNS(W_NS, 'txbxContent'))) {
      if (insideSkipped(tb, el)) continue;
      for (const p of Array.from(tb.getElementsByTagNameNS(W_NS, 'p'))) {
        const t = runText(p).trim();
        if (t) out.push([t]);
      }
    }
    return out;
  }

  function runText(el) {
    let out = '';
    const all = el.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
      const n = all[i];
      if (n.namespaceURI !== W_NS) continue;
      if (insideSkipped(n, el)) continue;
      if (n.localName === 't') out += n.textContent;
      else if (n.localName === 'tab') out += ' ';
      else if (n.localName === 'br' || n.localName === 'cr') out += ' ';
    }
    return out;
  }

  async function extractDocx(buf) {
    const zip = await window.JSZip.loadAsync(buf);
    // PV*SOL names its main part document2.xml, so follow the package relationship
    let mainPath = 'word/document.xml';
    const relsFile = zip.file('_rels/.rels');
    if (relsFile) {
      const rels = parseXml(await relsFile.async('string')).getElementsByTagName('Relationship');
      for (const r of rels) {
        if (/\/officeDocument$/.test(r.getAttribute('Type') || '')) mainPath = (r.getAttribute('Target') || mainPath).replace(/^\//, '');
      }
    }
    const main = zip.file(mainPath);
    if (!main) throw new Error('This .docx does not contain a readable document body.');
    const doc = parseXml(await main.async('string'));
    const body = doc.getElementsByTagNameNS(W_NS, 'body')[0];
    const lines = [];
    const walk = (node) => {
      for (const el of Array.from(node.children)) {
        const n = el.localName;
        if (n === 'p') {
          textBoxLines(el).forEach((l) => lines.push(l));
          const t = runText(el).trim();
          if (t) lines.push([t]);
        } else if (n === 'tbl') {
          for (const tr of Array.from(el.children)) {
            if (tr.localName !== 'tr') continue;
            const cells = [];
            for (const tc of Array.from(tr.children)) {
              if (tc.localName !== 'tc') continue;
              const nested = tc.getElementsByTagNameNS(W_NS, 'tbl');
              if (nested.length) { walk(tc); continue; }
              const paras = Array.from(tc.getElementsByTagNameNS(W_NS, 'p')).filter((p) => !insideSkipped(p, tc)).map(runText).map((s) => s.trim()).filter(Boolean);
              cells.push(paras.join(' '));
            }
            if (cells.some(Boolean)) lines.push(cells);
          }
        } else if (['sdt', 'sdtContent', 'customXml', 'smartTag', 'ins'].includes(n)) {
          walk(el);
        }
      }
    };
    walk(body);
    return lines;
  }

  let pdfjsPromise = null;
  function b64ToBlobUrl(id, type) {
    const b64 = document.getElementById(id).textContent.trim();
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return URL.createObjectURL(new Blob([bytes], { type }));
  }
  function loadPdfjs() {
    if (!pdfjsPromise) {
      pdfjsPromise = import(b64ToBlobUrl('lib-pdfjs', 'text/javascript')).then(async (lib) => {
        const workerUrl = b64ToBlobUrl('lib-pdfworker', 'text/javascript');
        // A background worker keeps the page responsive. When the page is opened
        // straight from disk (file://) browsers block it, so fall back to reading
        // the PDF on the main thread instead.
        const workerOk = await new Promise((resolve) => {
          let w;
          try { w = new Worker(workerUrl, { type: 'module' }); } catch (e) { resolve(null); return; }
          const done = (ok) => { clearTimeout(t); w.onmessage = null; w.onerror = null; resolve(ok ? w : (w.terminate(), null)); };
          const t = setTimeout(() => done(false), 5000);
          w.onmessage = () => done(true);
          w.onerror = () => done(false);
        });
        if (workerOk) {
          lib.GlobalWorkerOptions.workerPort = workerOk;
        } else {
          globalThis.pdfjsWorker = await import(workerUrl);
          lib.GlobalWorkerOptions.workerSrc = workerUrl;
        }
        return lib;
      });
    }
    return pdfjsPromise;
  }

  async function extractPdf(buf) {
    const pdfjs = await loadPdfjs();
    const pdf = await pdfjs.getDocument({
      data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true, useSystemFonts: false, stopAtErrors: false,
    }).promise;
    const lines = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      const items = tc.items
        .filter((i) => i.str && i.str.trim())
        .map((i) => ({ s: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: Math.abs(i.transform[3]) || i.height || 8 }));
      const rows = [];
      items.sort((a, b) => b.y - a.y || a.x - b.x);
      for (const it of items) {
        let r = rows.find((row) => Math.abs(row.y - it.y) < Math.max(2, it.h * 0.45));
        if (!r) { r = { y: it.y, items: [] }; rows.push(r); }
        r.items.push(it);
      }
      rows.sort((a, b) => b.y - a.y);
      let prev = null;
      for (const r of rows) {
        r.items.sort((a, b) => a.x - b.x);
        const cells = [];
        let cur = null; let end = null;
        for (const it of r.items) {
          const gap = end == null ? 0 : it.x - end;
          if (cur !== null && gap < Math.max(4, it.h * 1.0)) cur.s += (gap > it.h * 0.12 ? ' ' : '') + it.s;
          else { if (cur !== null) cells.push(cur); cur = { s: it.s, x: it.x, h: it.h }; }
          end = it.x + it.w;
          cur.r = end;
        }
        if (cur !== null) cells.push(cur);
        const clean = cells.filter((c) => c.s.trim());
        if (!clean.length) continue;
        // A value that wrapped onto a second line in a two column table: a lone
        // cell sitting directly under the previous row's value column.
        if (prev && clean.length === 1 && prev.cells.length >= 2 && prev.y - r.y < clean[0].h * 1.8) {
          const target = prev.cells.find((c, i) => i > 0 && (Math.abs(c.x - clean[0].x) < 3 || Math.abs(c.r - clean[0].r) < 3));
          if (target && !/^\d+\.\s/.test(clean[0].s)) {
            target.s += ' ' + clean[0].s;
            lines[lines.length - 1] = prev.cells.map((c) => c.s.trim());
            prev.y = r.y;
            continue;
          }
        }
        prev = { y: r.y, cells: clean };
        lines.push(clean.map((c) => c.s.trim()));
      }
    }
    return lines;
  }

  // ------------------------------------------------------------------
  // 2. Field matching. Labels are matched at the start of a line, so the
  //    same rules work on Word tables and on PDF text rows.
  // ------------------------------------------------------------------

  const norm = (s) => String(s).replace(/ /g, ' ').replace(/CO₂/g, 'CO2').replace(/\s+/g, ' ').trim().toLowerCase();
  const NUM_RE = /[-−]?\s?£?\s?[-−]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|[-−]?\s?£?\s?[-−]?\d+(?:\.\d+)?/;

  function toNum(s) {
    if (s == null) return null;
    const m = String(s).match(NUM_RE);
    if (!m) return null;
    const neg = /[-−]/.test(m[0]);
    const v = parseFloat(m[0].replace(/[^\d.]/g, ''));
    return Number.isFinite(v) ? (neg ? -v : v) : null;
  }

  function buildIndex(rawLines) {
    return rawLines.map((cells) => {
      const c = cells.map((x) => String(x).replace(/\s+/g, ' ').trim());
      const nonEmpty = c.filter(Boolean);
      return { cells: c, text: nonEmpty.join(' | '), n: norm(nonEmpty.join(' ')) };
    }).filter((l) => l.text);
  }

  function makeFinder(lines) {
    // value text after a label on the same line
    function after(line, label) {
      const L = norm(label);
      // Prefer cell structure: label in first non-empty cell
      const ne = line.cells.filter(Boolean);
      if (ne.length > 1 && norm(ne[0]) === L) return ne.slice(1).join(' ');
      if (ne.length > 1 && norm(ne[0]).startsWith(L)) {
        const restOfCell = ne[0].replace(/\s+/g, ' ').slice(label.length).trim();
        return [restOfCell].concat(ne.slice(1)).filter(Boolean).join(' ');
      }
      const flat = ne.join(' ');
      const idx = norm(flat).indexOf(L);
      return flat.slice(idx + label.length).replace(/^[\s|:]+/, '').trim();
    }
    const starts = (line, label) => line.n.startsWith(norm(label));

    return {
      num(labels, opts = {}) {
        labels = [].concat(labels);
        for (const label of labels) {
          for (let i = 0; i < lines.length; i++) {
            if (!starts(lines[i], label)) continue;
            const rest = after(lines[i], label);
            if (opts.unit && !new RegExp(opts.unit, 'i').test(rest) && !new RegExp(opts.unit, 'i').test(lines[i + 1] ? lines[i + 1].text : '')) continue;
            const v = toNum(rest);
            if (v != null) return v;
          }
        }
        // PDF fallback: label on its own line, value on the next line
        for (const label of labels) {
          for (let i = 0; i < lines.length - 1; i++) {
            if (lines[i].n !== norm(label)) continue;
            const nx = lines[i + 1].text;
            if (/^[-−£\d.,\s]+(kwp|kwh|%|£|m²|kg|years?)?/i.test(nx) && /^\s*[-−£\d]/.test(nx)) {
              const v = toNum(nx);
              if (v != null) return v;
            }
          }
        }
        return null;
      },
      text(labels) {
        labels = [].concat(labels);
        for (const label of labels) {
          for (const line of lines) {
            if (!starts(line, label)) continue;
            const rest = after(line, label);
            if (rest) return rest;
          }
        }
        return '';
      },
      all(label) { return lines.filter((l) => starts(l, label)).map((l) => after(l, label)); },
      lines,
    };
  }

  function extractFields(rawLines) {
    const lines = buildIndex(rawLines);
    const f = makeFinder(lines);
    const d = {};

    const firstDate = lines.slice(0, 12).map((l) => l.text.match(/\b(\d{2}\/\d{2}\/\d{4})\b/)).find(Boolean);
    d.reportDate = firstDate ? firstDate[1] : '';
    d.company = cleanText(f.text('Company'));
    d.contact = cleanText(f.text('Contact person'));
    d.projectName = cleanText(f.text('Project Name'));
    d.address = cleanText(f.text('Address'));
    d.offerNo = cleanText(f.text(['Offer no.', 'Offer No', 'Offer number'])).split(' ')[0] || '';
    if (!d.offerNo) { const m = lines.slice(0, 6).map((l) => l.text.match(/Documentation\s*[-|]\s*(\S+)/i)).find(Boolean); if (m) d.offerNo = m[1]; }
    d.designer = cleanText(f.text('Project Designer'));
    d.location = cleanText(f.text(['Climate Data', 'Location']));

    d.kwp = f.num('PV Generator Output', { unit: 'kwp' }) ?? f.num('PV Generator Output');
    d.surface = f.num('PV Generator Surface');
    d.modules = f.num('Number of PV Modules');
    d.inverters = f.num('Number of Inverters');
    d.specYield = f.num('Spec. Annual Yield');
    d.pr = f.num('Performance Ratio (PR)');
    d.shading = f.num('Yield Reduction due to Shading');
    d.generation = f.num('PV Generator Energy (AC grid)', { unit: 'kwh' });
    d.ownConsumption = f.num('Own Consumption');
    d.gridExport = f.num('Grid Export');
    d.ownUsePct = f.num('Own Power Consumption');
    d.co2 = f.num(['CO2 Emissions avoided', 'CO2 emissions avoided']);
    d.selfSufficiency = f.num('Level of Self-sufficiency');
    d.consumption = f.num('Total Consumption');
    d.loadPeak = f.num('Load Peak');

    d.investment = f.num(['Total investment costs', 'Investment Costs']);
    d.irr = f.num('Internal Rate of Return');
    d.amortisation = cleanText(f.text('Amortization Period')) || cleanText(f.text('Amortisation Period'));
    d.lcoe = f.num('Electricity Production Costs');
    d.firstYearSavings = f.num('First year savings');
    d.energyPrice = f.num('Energy Price');
    d.inflation = f.num(['Inflation Rate for Energy Price', 'Price Increase Rate']);
    d.period = f.num(['Assessment Period', 'Remuneration period']);
    d.exportIncome = f.num('Total Payment from Utility in First Year');
    d.startOperation = cleanText(f.text('Start of Operation'));

    // Module areas
    const areas = [];
    let cur = null;
    for (const l of lines) {
      const h = l.text.match(/^(\d+)\.\s*Module Area\s*[-–]\s*(.+)$/i);
      if (h) {
        const idx = +h[1];
        cur = areas.find((a) => a.idx === idx);
        if (!cur) { cur = { idx, name: h[2].replace(/\s*\|.*$/, '').trim() }; areas.push(cur); }
        continue;
      }
      if (/^(Inverter configuration|Configuration \d|AC Mains|Simulation Results)/i.test(l.text)) cur = null;
      if (!cur) continue;
      if (cur.count == null && /^PV Modules\b/i.test(l.text)) {
        const v = l.text.replace(/^PV Modules\s*\|?\s*/i, '');
        const m = v.match(/^(\d+)\s*x\s*(.+)$/i);
        if (m) { cur.count = +m[1]; cur.type = m[2].trim(); }
      }
      if (cur.tilt == null && /^Inclination\b/i.test(l.text)) cur.tilt = toNum(l.text.replace(/^Inclination/i, ''));
      if (cur.orient == null && /^Orientation\b/i.test(l.text)) cur.orient = l.text.replace(/^Orientation\s*\|?\s*/i, '').replace(/\s*\|?\s*°\s*$/, '').trim();
    }
    d.areas = areas;
    d.moduleType = uniq(areas.map((a) => a.type).filter(Boolean)).join(', ') || cleanText(f.text('PV Module:'));

    // Inverters: model and the manufacturer line that follows it
    const models = [];
    lines.forEach((l, i) => {
      if (/^Model\b/i.test(l.text)) {
        const model = l.text.replace(/^Model\s*\|?\s*/i, '').trim();
        const man = lines[i + 1] && /^Manufacturer\b/i.test(lines[i + 1].text) ? lines[i + 1].text.replace(/^Manufacturer\s*\|?\s*/i, '').trim() : '';
        if (model) models.push((man ? man + ' ' : '') + model);
      }
    });
    d.inverterModel = uniq(models).join(', ');

    // 25 year electricity savings from the PV*SOL cash flow table
    const series = [];
    for (const l of lines) {
      if (!/^electricity savings\b/i.test(l.n)) continue;
      const rest = l.text.replace(/^Electricity Savings/i, '');
      const vals = rest.match(/[-−]?£\s?[-−]?[\d,]+(?:\.\d+)?|[-−]?[\d,]{4,}\.\d{2}/g) || [];
      vals.forEach((v) => { const n = toNum(v); if (n != null) series.push(n); });
    }
    d.savingsSeries = series;
    return d;
  }

  function cleanText(s) { return String(s || '').replace(/\s*\|\s*/g, ' ').replace(/\s+/g, ' ').trim(); }
  function uniq(a) { return Array.from(new Set(a)); }

  // ------------------------------------------------------------------
  // 3. Review fields shown in the side panel (editable overrides)
  // ------------------------------------------------------------------

  const FIELDS = [
    ['Project', [
      ['client', 'Client name', 'text', true],
      ['address', 'Site address', 'text', false],
      ['offerNo', 'PV*SOL offer number', 'text', true],
      ['designer', 'Project designer', 'text', false],
      ['reportDate', 'PV*SOL report date', 'text', false],
      ['prepared', 'Explainer date', 'text', true],
    ]],
    ['System', [
      ['kwp', 'System size (kWp)', 'num', true],
      ['modules', 'Number of PV modules', 'num', false],
      ['moduleType', 'Module type', 'text', false],
      ['inverters', 'Number of inverters', 'num', false],
      ['inverterModel', 'Inverter model', 'text', false],
      ['areasText', 'Roof areas', 'text', false],
      ['generation', 'Annual PV generation (kWh)', 'num', true],
      ['ownUsePct', 'PV used on site (%)', 'num', false],
      ['selfSufficiency', 'Self sufficiency (%)', 'num', false],
      ['co2', 'CO2 avoided (kg a year)', 'num', false],
      ['consumption', 'Site consumption (kWh a year)', 'num', false],
    ]],
    ['Finance model', [
      ['investment', 'Investment, excluding VAT (£)', 'num', true],
      ['firstYearSavings', 'First year savings (£)', 'num', true],
      ['energyPrice', 'Energy price (£/kWh)', 'num', false],
      ['inflation', 'Energy price inflation (% a year)', 'num', true],
      ['period', 'Appraisal period (years)', 'num', true],
    ]],
  ];

  function areasToText(areas) {
    return (areas || []).map((a) => {
      const dir = a.orient ? a.orient.replace(/\s*\d+(\.\d+)?\s*°?$/, '').trim() : '';
      return [a.name, a.count ? a.count + ' modules' : '', dir].filter(Boolean).join(', ');
    }).join('; ');
  }

  // ------------------------------------------------------------------
  // 4. Finance model
  // ------------------------------------------------------------------

  function buildSeries(d) {
    const N = Math.max(1, Math.round(d.period || 25));
    const s = (d.savingsSeries || []).slice(0, N);
    const ok = s.length === N && d.firstYearSavings != null && Math.abs(s[0] - d.firstYearSavings) < 1;
    if (ok) return { series: s, fromReport: true };
    const g = (d.inflation || 0) / 100;
    const out = [];
    for (let y = 1; y <= N; y++) out.push(d.firstYearSavings * Math.pow(1 + g, y - 1));
    return { series: out, fromReport: false };
  }

  function model(d, pay) {
    const { series, fromReport } = buildSeries(d);
    const N = series.length;
    const opts = [
      { key: '5', years: 5, months: 60, monthly: pay[5] },
      { key: '7', years: 7, months: 84, monthly: pay[7] },
    ].map((o) => {
      o.annual = o.monthly * 12;
      o.total = o.monthly * o.months;
      o.financeCost = o.total - d.investment;
      o.rows = [];
      let cum = 0;
      for (let y = 1; y <= N; y++) {
        const p = y <= o.years ? o.annual : 0;
        const net = series[y - 1] - p;
        cum += net;
        o.rows.push({ y, pay: p, net, cum });
      }
      o.y1net = o.rows[0].net;
      o.firstPosCum = (o.rows.find((r) => r.cum > 0) || {}).y || null;
      o.firstPosAnnual = (o.rows.find((r) => r.net > 0) || {}).y || null;
      o.cumAtTermEnd = o.rows[Math.min(o.years, N) - 1].cum;
      o.endNet = o.rows[N - 1].cum;
      return o;
    });
    // Sensitivity on the longer term option
    const long = opts[1];
    const base = d.inflation || 0;
    let rates = uniq([0, 3, base].map((r) => +(+r).toFixed(2))).sort((a, b) => a - b);
    if (rates.length < 3) rates = uniq(rates.concat([5])).sort((a, b) => a - b);
    const sens = rates.map((r) => {
      let sum = 0;
      if (r === +(+base).toFixed(2)) sum = series.reduce((a, b) => a + b, 0);
      else for (let y = 1; y <= N; y++) sum += d.firstYearSavings * Math.pow(1 + r / 100, y - 1);
      return { rate: r, base: r === +(+base).toFixed(2), value: sum - long.total };
    });
    return { N, series, fromReport, opts, sens };
  }

  // ------------------------------------------------------------------
  // 5. Formatting
  // ------------------------------------------------------------------

  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = (v, dp = 0) => Number(v).toLocaleString('en-GB', { minimumFractionDigits: dp, maximumFractionDigits: dp });
  const gbp = (v, dp = 0) => (v < 0 ? '(£' + nf(Math.abs(v), dp) + ')' : '£' + nf(v, dp));
  const gbpSigned = (v) => (Math.round(v) < 0 ? '(£' + nf(Math.abs(v)) + ')' : '+£' + nf(v));
  const cls = (v) => (Math.round(v) < 0 ? 'neg' : 'pos');
  const short = (v) => {
    const a = Math.abs(v);
    let s;
    if (a >= 1e6) s = '£' + (a / 1e6).toFixed(2) + 'm';
    else if (a >= 1e5) s = '£' + (a / 1e6).toFixed(2) + 'm';
    else s = '£' + nf(a);
    return v < 0 ? '(' + s + ')' : s;
  };
  const words = { 5: 'five', 7: 'seven' };
  const pct = (v, dp = 1) => nf(v, dp) + '%';
  const trimNum = (v, dp) => { const s = nf(v, dp); return s.replace(/\.0+$/, ''); };

  function displayClient(d) { return d.client || 'Client'; }

  // ------------------------------------------------------------------
  // 6. Page templates (Garland UK Brand Guidelines v2.0)
  // ------------------------------------------------------------------

  const TOTAL_PAGES = 5;

  function head(d, section) {
    return `<div class="hd"><span><b>Solar PV asset finance</b> | ${esc(displayClient(d))}</span><span>${esc(section)}</span></div>`;
  }
  function foot(d, n) {
    const ref = d.offerNo ? `Illustration based on PV*SOL report ${esc(d.offerNo)}. ` : '';
    return `<div class="ft"><img src="${LOGO_H}" alt="Garland UK"><span>${ref}Not a finance quotation | Prepared ${esc(d.prepared)}</span><span class="pg">${n}</span></div>`;
  }
  function intro(num, title, lead, extra = '', caption = '') {
    return `<div class="intro">
      <div class="secnum">${num}</div>
      <h1 class="h1">${esc(title)}</h1>
      ${lead ? `<p class="lead">${lead}</p>` : ''}
      ${extra}
      ${caption ? `<p class="caption">${caption}</p>` : ''}
    </div>`;
  }

  function headlineCopy(d, m) {
    const [o5, o7] = m.opts;
    const S1 = gbp(d.firstYearSavings);
    const cond = 'Subject to validated demand, tariffs, scope and lender terms.';
    const best = o7.y1net >= o5.y1net ? o7 : o5;
    const label = `Year 1 net cash with ${best.years} year finance`;
    if (o5.y1net >= 0 && o7.y1net >= 0) {
      return {
        big: '+' + gbp(best.y1net), label,
        title: 'Both finance options start cash flow positive',
        body: `Forecast Year 1 electricity savings of ${S1} are more than the annual repayments under either option. ${cond}`,
      };
    }
    if (best.y1net >= 0) {
      return {
        big: '+' + gbp(best.y1net), label,
        title: `The ${best.years} year option starts cash flow positive`,
        body: `Forecast Year 1 electricity savings of ${S1} are more than the annual repayments of about ${gbp(best.annual)}. ${cond}`,
      };
    }
    const cover = Math.round((d.firstYearSavings / best.annual) * 100);
    const title = best.firstPosAnnual && best.firstPosAnnual <= best.years
      ? `Savings overtake repayments from Year ${best.firstPosAnnual}`
      : 'Savings build strongly once finance ends';
    const tail = best.firstPosCum
      ? `With the modelled energy prices, the ${best.years} year option is cash positive overall from Year ${best.firstPosCum}.`
      : `With the modelled energy prices, the ${best.years} year option does not become cash positive overall within ${m.N} years.`;
    return {
      big: gbp(best.y1net), label, title,
      body: `Forecast Year 1 electricity savings of ${S1} cover ${cover}% of the ${best.years} year option's annual repayments of ${gbp(best.annual)}. ${tail} ${cond}`,
    };
  }

  const direction = (o) => (o || '').replace(/\s*\d+(\.\d+)?\s*°?$/, '').trim();

  function pageCover(d, m) {
    const client = displayClient(d);
    const fs = client.length > 40 ? 26 : client.length > 26 ? 30 : 34;
    const meta = [];
    if (d.address) meta.push(['Site', esc(d.address)]);
    if (d.offerNo) meta.push(['PV*SOL offer', esc(d.offerNo) + (d.designer ? `, designed by ${esc(d.designer)}` : '')]);
    if (d.reportDate) meta.push(['Report dated', esc(d.reportDate)]);
    meta.push(['Prepared', esc(d.prepared)]);
    return `<section class="page cover">
      <img class="photo" src="${HERO}" alt="">
      <svg class="grad" viewBox="0 0 297 134" preserveAspectRatio="none" aria-hidden="true">
        <defs><radialGradient id="pvgrad" cx="1.02" cy="1.18" r="1.15">
          <stop offset="0" stop-color="#40C1AC"/><stop offset=".24" stop-color="#1C7A6E"/>
          <stop offset=".48" stop-color="#0E4A43"/><stop offset=".74" stop-color="#062E29"/><stop offset="1" stop-color="#062E29"/>
        </radialGradient></defs>
        <rect width="297" height="134" fill="url(#pvgrad)"/>
      </svg>
      <div class="titleblock">
        <div class="kicker">Solar PV asset finance illustration</div>
        <h1 class="ch1" style="font-size:${fs}pt">Financing solar PV at ${esc(client)}</h1>
        <p class="clead">How 5 and 7 year asset finance compare with the modelled savings from a ${trimNum(d.kwp, 2)} kWp solar PV system.</p>
      </div>
      <img class="logo" src="${LOGO_REV}" alt="Garland UK">
      <div class="foot-g">
        <dl class="meta">${meta.map(([k, v]) => `<dt>${k}</dt><dd>${v}</dd>`).join('')}</dl>
        <div class="signoff">Partners for the life of your building</div>
      </div>
    </section>`;
  }

  function pageGlance(d, m) {
    const h = headlineCopy(d, m);
    const areas = (d.areas || []).filter((a) => a.name);
    const lead = `PV*SOL models a ${trimNum(d.kwp, 2)} kWp system${areas.length > 1 ? ` across ${areas.length} roof areas` : ''}, generating ${nf(d.generation)} kWh of electricity a year. In Year 1 that is worth ${gbp(d.firstYearSavings)} in reduced electricity bills.`;
    const src = [`Figures from PV*SOL report ${esc(d.offerNo || '')}`];
    if (d.reportDate) src.push(`dated ${esc(d.reportDate)}`);
    if (d.location) src.push(`using ${esc(d.location)} climate data`);
    const rows = [];
    if (d.modules) rows.push(['PV modules', `${nf(d.modules)}${d.moduleType ? ' × ' + esc(d.moduleType) : ''}`]);
    if (d.inverters) rows.push(['Inverters', `${nf(d.inverters)}${d.inverterModel ? ' × ' + esc(d.inverterModel) : ''}`]);
    if (areas.length) rows.push(['Roof areas', areas.map((a) => `${esc(a.name)}${a.count ? `, ${a.count} modules` : ''}${a.orient ? `, facing ${esc(direction(a.orient).toLowerCase())}` : ''}`).join('<br>')]);
    else if (d.areasText) rows.push(['Roof areas', esc(d.areasText)]);
    if (d.ownUsePct != null) rows.push(['Used on site', `${trimNum(d.ownUsePct, 1)}% of the electricity generated`]);
    if (d.selfSufficiency != null) rows.push(['Site demand met by PV', `${trimNum(d.selfSufficiency, 1)}%`]);
    if (d.co2 != null) rows.push(['CO<sub>2</sub> emissions avoided', `${trimNum(d.co2 / 1000, 1)} tonnes a year`]);
    return `<section class="page">
      ${head(d, '01 | At a glance')}
      <div class="body">
        ${intro('01', 'The system and what it saves', lead, '', src.join(', ') + '.')}
        <div class="content">
          <div class="stats">
            <div class="stat"><div class="v">${trimNum(d.kwp, 2)} kWp</div><div class="l">System size</div></div>
            <div class="stat"><div class="v">${gbp(d.investment)}</div><div class="l">Estimated budget, excluding VAT</div></div>
            <div class="stat"><div class="v">${nf(d.generation)} kWh</div><div class="l">Electricity generated a year</div></div>
            <div class="stat"><div class="v">${gbp(d.firstYearSavings)}</div><div class="l">Modelled saving in Year 1</div></div>
          </div>
          <div class="bluepanel" style="margin-top:8mm">
            <div><div class="big">${h.big}</div><div class="bl">${esc(h.label)}</div></div>
            <div><div class="t">${esc(h.title)}</div><div class="d">${esc(h.body)}</div></div>
          </div>
          ${rows.length ? `<table class="bt compact" style="margin-top:7mm">
            <thead><tr><th style="width:52mm">System</th><th>From the PV*SOL design</th></tr></thead>
            <tbody>${rows.map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('')}</tbody>
          </table>` : ''}
        </div>
      </div>
      ${foot(d, 2)}
    </section>`;
  }

  function pageFinance(d, m) {
    const [o5, o7] = m.opts;
    const row = (label, a, b, em) => `<tr${em ? ' class="em"' : ''}><td>${label}</td><td class="r">${a}</td><td class="r">${b}</td></tr>`;
    const yr = (y) => (y ? 'Year ' + y : `After Year ${m.N}`);
    return `<section class="page">
      ${head(d, '02 | Finance options')}
      <div class="body">
        ${intro('02', 'Two ways to finance one installation',
          `Both options fund the same ${gbp(d.investment)} budget and are measured against the same ${gbp(d.firstYearSavings)} of modelled Year 1 savings.`,
          '', 'Totals use 60 or 84 fixed monthly payments and are rounded to the nearest pound, so a subtraction may differ by £1.')}
        <div class="content">
          <table class="bt">
            <thead><tr><th>Measure</th><th class="r" style="width:42mm">5 year finance</th><th class="r" style="width:42mm">7 year finance</th></tr></thead>
            <tbody>
              ${row('Finance term', '60 months', '84 months')}
              ${row('Monthly repayment', gbp(o5.monthly, 2), gbp(o7.monthly, 2))}
              ${row('Annual repayments', gbp(o5.annual), gbp(o7.annual))}
              ${row('Electricity saving in Year 1', gbp(d.firstYearSavings), gbp(d.firstYearSavings))}
              ${row('Net cash in Year 1', gbpSigned(o5.y1net), gbpSigned(o7.y1net), true)}
              ${row('First year cash positive overall', yr(o5.firstPosCum), yr(o7.firstPosCum))}
              ${row('Total repaid over the term', gbp(o5.total), gbp(o7.total))}
              ${row('Cost of finance above the budget', gbp(o5.financeCost), gbp(o7.financeCost))}
              ${row(`Net cash by Year ${m.N}`, gbp(o5.endNet), gbp(o7.endNet), true)}
            </tbody>
          </table>
          <h2 class="h3" style="margin-top:10mm">Why this matters for decision makers</h2>
          <div class="cols3" style="margin-top:4mm">
            <div class="ruled"><h3 class="h3">Capital preservation</h3><p>Financing may avoid a large capital purchase at the outset, leaving funds available for other priorities.</p></div>
            <div class="ruled"><h3 class="h3">Predictable repayments</h3><p>Agreed monthly instalments can be compared directly with the forecast fall in electricity bought from the grid.</p></div>
            <div class="ruled"><h3 class="h3">Long-term value</h3><p>Once finance ends, the modelled electricity savings continue for the rest of the appraisal period, subject to system performance.</p></div>
          </div>
        </div>
      </div>
      ${foot(d, 3)}
    </section>`;
  }

  function niceStep(span) {
    const raw = span / 5;
    const p = Math.pow(10, Math.floor(Math.log10(raw)));
    const n = raw / p;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p;
  }

  function chartSvg(m) {
    const Wd = 700, Ht = 262, L = 56, R = 8, T = 22, B = 34;
    const [o5, o7] = m.opts;
    const pts = (o) => [{ y: 0, cum: 0 }].concat(o.rows.map((r) => ({ y: r.y, cum: r.cum })));
    const all = o5.rows.concat(o7.rows).map((r) => r.cum).concat([0]);
    let lo = Math.min(...all), hi = Math.max(...all);
    const step = niceStep(Math.max(hi - lo, 1));
    hi = Math.ceil(hi / step) * step;
    // Keep the area below zero only as deep as the data needs
    lo = lo < 0 ? Math.floor(lo / (step / 4)) * (step / 4) : 0;
    const x = (y) => L + (y / m.N) * (Wd - L - R);
    const yv = (v) => T + (1 - (v - lo) / (hi - lo)) * (Ht - T - B);
    const fmtAxis = (v) => {
      const a = Math.abs(v);
      const s = a >= 1e6 ? '£' + (a / 1e6).toFixed(2).replace(/\.?0+$/, '') + 'm' : a >= 1e3 ? '£' + Math.round(a / 1e3) + 'k' : '£' + a;
      return v < 0 ? '-' + s : s;
    };
    const txt = 'fill="#101820" fill-opacity=".7" font-size="10"';
    let g = '';
    const ticks = [];
    for (let v = 0; v <= hi + 1e-6; v += step) ticks.push(v);
    for (let v = -step; v >= lo - 1e-6; v -= step) ticks.push(v);
    for (const v of ticks) {
      const zero = Math.abs(v) < 1e-6;
      g += `<line x1="${L}" x2="${Wd - R}" y1="${yv(v)}" y2="${yv(v)}" stroke="${zero ? '#9BB8D3' : '#E7EAED'}" stroke-width="${zero ? 1.4 : 1}"/>`;
      g += `<text x="${L - 8}" y="${yv(v) + 3.5}" text-anchor="end" ${txt}>${fmtAxis(v)}</text>`;
    }
    if (lo < 0) g += `<line x1="${L}" x2="${Wd - R}" y1="${yv(lo)}" y2="${yv(lo)}" stroke="#E7EAED" stroke-width="1"/>`;
    uniq([1, 5, 7, 10, 15, 20, m.N].filter((t) => t <= m.N)).forEach((t) => {
      g += `<text x="${x(t)}" y="${Ht - B + 17}" text-anchor="middle" ${txt}>${t}</text>`;
    });
    g += `<text x="${L}" y="${Ht - 2}" ${txt}>Years from installation</text>`;
    [o5, o7].forEach((o, i) => {
      const xx = x(o.years);
      g += `<line x1="${xx}" x2="${xx}" y1="${T - 4}" y2="${Ht - B}" stroke="#101820" stroke-opacity=".35" stroke-width="1" stroke-dasharray="2 3"/>`;
      g += `<text x="${xx + (i ? 5 : -5)}" y="${T - 8}" text-anchor="${i ? 'start' : 'end'}" ${txt}>${o.years} year finance ends</text>`;
    });
    const path = (o) => pts(o).map((p, i) => (i ? 'L' : 'M') + x(p.y).toFixed(1) + ' ' + yv(p.cum).toFixed(1)).join(' ');
    g += `<path d="${path(o5)}" fill="none" stroke="#003C71" stroke-width="2.6" stroke-linejoin="round"/>`;
    g += `<path d="${path(o7)}" fill="none" stroke="#FF6720" stroke-width="2.6" stroke-linejoin="round"/>`;
    return `<svg viewBox="0 0 ${Wd} ${Ht}" role="img" aria-label="Cumulative net cash after finance over ${m.N} years">${g}</svg>`;
  }

  function pageValue(d, m) {
    const [o5, o7] = m.opts;
    const stack = `<div class="stack">
      <div class="stat"><div class="v">${short(o7.cumAtTermEnd)}</div><div class="l">Cash position when 7 year finance ends</div></div>
      <div class="stat"><div class="v">${short(o5.endNet)}</div><div class="l">Net cash by Year ${m.N} with 5 year finance</div></div>
      <div class="stat"><div class="v">${short(o7.endNet)}</div><div class="l">Net cash by Year ${m.N} with 7 year finance</div></div>
    </div>`;
    const sens = m.sens.map((s) => `<div class="ruled s">
        <div class="k">${trimNum(s.rate, 2)}% a year${s.base ? ', PV*SOL assumption' : ''}</div>
        <div class="v">${short(s.value)}</div>
        <div class="l">Net cash by Year ${m.N}</div>
      </div>`).join('');
    return `<section class="page">
      ${head(d, '03 | Long-term value')}
      <div class="body">
        ${intro('03', `A ${m.N} year view of the cash benefit`, 'Cumulative cash after repayments, using the yearly electricity savings from PV*SOL.', stack)}
        <div class="content">
          <div class="legend"><span><i style="background:#003C71"></i>5 year finance</span><span><i style="background:#FF6720"></i>7 year finance</span></div>
          <div class="chart">${chartSvg(m)}</div>
          <h2 class="h3" style="margin-top:5mm">If energy prices rise more slowly</h2>
          <div class="cols3 sens" style="margin-top:3mm">${sens}</div>
          <p class="caption" style="margin-top:2.5mm">Shown for 7 year finance. Only the yearly growth in savings changes; the Year 1 saving and the repayments stay the same. Nominal figures, not discounted, and before maintenance and other costs.</p>
          <div class="lightpanel" style="margin-top:5mm">
            <h3 class="h3">Next step</h3>
            <p>Confirm annual and ideally half-hourly electricity demand, recent tariffs, roof and electrical details and VAT status. Garland can then prepare a feasibility study for the site and seek indicative finance terms.</p>
          </div>
        </div>
      </div>
      ${foot(d, 4)}
    </section>`;
  }

  function pageSchedule(d, m) {
    const [o5, o7] = m.opts;
    const rowh = Math.min(4.6, 128 / m.N);
    const rowf = Math.min(8, rowh * 1.75);
    const pay = (p) => (p ? `(${gbp(p).replace(/[()]/g, '')})` : '–');
    const rows = m.series.map((s, i) => {
      const a = o5.rows[i], b = o7.rows[i];
      const c = [i + 1 === 5 || i + 1 === 7 ? 'end' : '', i + 1 === m.N ? 'last' : ''].filter(Boolean).join(' ');
      return `<tr${c ? ` class="${c}"` : ''}><td class="yr">${i + 1}</td><td>${gbp(s)}</td>
        <td class="${a.pay ? '' : 'nil'}">${pay(a.pay)}</td><td>${gbp(a.net)}</td><td>${gbp(a.cum)}</td>
        <td class="${b.pay ? '' : 'nil'}">${pay(b.pay)}</td><td>${gbp(b.net)}</td><td>${gbp(b.cum)}</td></tr>`;
    }).join('');
    const own = d.ownUsePct != null ? ` PV*SOL models ${trimNum(d.ownUsePct, 1)}% of the electricity being used on site.` : '';
    const price = d.energyPrice != null ? `an energy price of £${nf(d.energyPrice, 4).replace(/0+$/, '')} per kWh, rising by ` : 'energy prices rising by ';
    const exportNote = d.exportIncome ? ` PV*SOL also models ${gbp(d.exportIncome)} of export income in Year 1, which is not included here.` : ' Export income is not included.';
    const seriesNote = m.fromReport ? '' : ' Yearly savings were calculated from the PV*SOL first year saving and inflation rate because the cash flow table could not be read.';
    const notes = `<div class="leftnotes">
      <h3 class="h3">How to read the table</h3>
      <p>Net is the electricity saving less finance repayments for that year. Cumulative is the running total. Repayments stop after Year 5 or Year 7; later savings assume the system keeps generating at modelled energy prices.</p>
      <h3 class="h3">Basis and exclusions</h3>
      <p>The ${gbp(d.investment)} budget, excluding VAT, is assumed to be financed in full with 60 or 84 fixed monthly payments and no separate deposit, balloon payment or fees. The PV*SOL model assumes ${price}${trimNum(d.inflation, 2)}% a year.${own}${exportNote} Finance availability, VAT, any capital contribution and actual interest terms need confirming. Access and scaffolding, electrical upgrades, maintenance, insurance, inverter replacement and other lifecycle costs are outside this illustration. Figures are indicative and not guaranteed.</p>
    </div>`;
    return `<section class="page" style="--rowh:${rowh}mm;--rowf:${rowf}pt">
      ${head(d, '04 | Year by year')}
      <div class="body">
        ${intro('04', 'Year by year cash flow', 'Both finance options side by side, in nominal pounds.', notes)}
        <div class="content">
          <table class="st">
            <thead>
              <tr><th class="yr" rowspan="2">Year</th><th rowspan="2">Electricity saving</th><th class="grp" colspan="3"><span>5 year finance</span></th><th class="grp" colspan="3"><span>7 year finance</span></th></tr>
              <tr><th>Repayment</th><th>Net</th><th>Cumulative</th><th>Repayment</th><th>Net</th><th>Cumulative</th></tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
          ${seriesNote ? `<p class="caption" style="margin-top:3mm">${seriesNote.trim()}</p>` : ''}
        </div>
      </div>
      ${foot(d, 5)}
    </section>`;
  }

  // ------------------------------------------------------------------
  // 7. App state and UI
  // ------------------------------------------------------------------

  const LOGO_H = document.getElementById('asset-logo-h').textContent.trim();
  const LOGO_REV = document.getElementById('asset-logo-rev').textContent.trim();
  const HERO = document.getElementById('asset-hero').textContent.trim();
  const state = { data: null, fileName: '', pay: { 5: null, 7: null } };

  function todayLong() {
    return new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  }

  function deriveClient(d) {
    if (d.company) return d.company;
    if (d.projectName) return d.projectName;
    if (d.address) return d.address.split(',')[0].trim();
    return '';
  }

  function missingRequired(d) {
    const miss = [];
    FIELDS.forEach(([, list]) => list.forEach(([key, label, , req]) => {
      if (req && (d[key] == null || d[key] === '' || (typeof d[key] === 'number' && !Number.isFinite(d[key])))) miss.push(label);
    }));
    return miss;
  }

  function renderReview() {
    const d = state.data;
    const box = $('#review');
    if (!d) { box.hidden = true; return; }
    box.hidden = false;
    let found = 0, total = 0;
    const html = FIELDS.map(([grp, list]) => `<h3>${grp}</h3>` + list.map(([key, label, type]) => {
      total++;
      const v = d[key];
      const has = !(v == null || v === '');
      if (has) found++;
      const shown = has ? (type === 'num' ? String(v) : v) : '';
      return `<div class="rf${has ? '' : ' missing'}"><label for="f-${key}">${esc(label)}</label><input type="text" id="f-${key}" data-key="${key}" data-type="${type}" value="${esc(shown)}"></div>`;
    }).join('')).join('');
    box.querySelector('.grp').innerHTML = html;
    box.querySelector('summary span').textContent = `Check what was read from the report (${found} of ${total})`;
    if (missingRequired(d).length) box.open = true;
  }

  function onReviewInput(e) {
    const el = e.target;
    if (!el.dataset || !el.dataset.key) return;
    const key = el.dataset.key;
    const raw = el.value.trim();
    if (el.dataset.type === 'num') state.data[key] = raw === '' ? null : toNum(raw);
    else state.data[key] = raw;
    if (key === 'areasText') state.data.areas = [];
    el.closest('.rf').classList.toggle('missing', raw === '');
    render();
  }

  function setStatus(kind, html) { const s = $('#status'); s.className = 'status ' + kind; s.innerHTML = html; }

  function parsePay(v) { const n = toNum(v); return n != null && n > 0 ? n : null; }

  function render() {
    const d = state.data;
    const out = $('#sheets');
    const empty = $('#empty');
    const btn = $('#print');
    if (!d) { out.innerHTML = ''; empty.hidden = false; btn.disabled = true; return; }
    const miss = missingRequired(d);
    const pay = state.pay;
    if (miss.length || !pay[5] || !pay[7]) {
      const need = [];
      if (!pay[5] || !pay[7]) need.push('both monthly repayments');
      if (miss.length) need.push(miss.join(', '));
      empty.hidden = false;
      empty.querySelector('.big').textContent = 'Almost there';
      empty.querySelector('.msg').textContent = 'Enter ' + need.join(' and ') + ' to see the explainer.';
      out.innerHTML = '';
      btn.disabled = true;
      return;
    }
    empty.hidden = true;
    const m = model(d, pay);
    out.innerHTML = pageCover(d, m) + pageGlance(d, m) + pageFinance(d, m) + pageValue(d, m) + pageSchedule(d, m);
    btn.disabled = false;
    fitZoom();
  }

  function fitZoom() {
    const pv = $('.preview');
    const avail = pv.clientWidth - 56;
    const pagePx = 297 * 96 / 25.4;
    const z = Math.max(0.35, Math.min(1, avail / pagePx));
    $('#sheets').style.setProperty('--z', z.toFixed(3));
  }

  async function handleFile(file) {
    if (!file) return;
    const name = file.name || 'file';
    const ext = (name.split('.').pop() || '').toLowerCase();
    if (!['docx', 'pdf'].includes(ext)) { setStatus('err', 'Please use a PV*SOL report saved as <b>.docx</b> or <b>.pdf</b>.'); return; }
    setStatus('', 'Reading ' + esc(name) + '…');
    try {
      const buf = await file.arrayBuffer();
      const lines = ext === 'pdf' ? await extractPdf(buf) : await extractDocx(buf);
      const d = extractFields(lines);
      if (d.kwp == null && d.investment == null && d.firstYearSavings == null) {
        throw new Error('No PV*SOL figures were found in this file. Check it is the PV*SOL project report.');
      }
      d.client = deriveClient(d);
      d.prepared = todayLong();
      d.areasText = areasToText(d.areas);
      state.data = d;
      state.fileName = name;
      renderReview();
      const miss = missingRequired(d);
      const seriesOk = buildSeries(d).fromReport;
      let msg = `Read <b>${esc(name)}</b>.`;
      if (miss.length) setStatus('warn', msg + ' Some figures were not found: ' + esc(miss.join(', ')) + '. Please enter them below.');
      else setStatus('ok', msg + (seriesOk ? ` All key figures and the ${d.savingsSeries.length >= 25 ? '25' : d.savingsSeries.length} year savings table were found.` : ' Key figures found. The yearly savings were calculated from the first year saving and inflation rate.'));
      render();
      const firstMissing = document.querySelector('.rf.missing input');
      (firstMissing || $('#pay5')).focus();
    } catch (err) {
      console.error(err);
      setStatus('err', esc(err && err.message ? err.message : 'The file could not be read.'));
    }
  }

  function init() {
    const drop = $('#drop');
    const input = $('#file');
    drop.addEventListener('click', () => input.click());
    drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => { handleFile(input.files[0]); input.value = ''; });
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement) drop.classList.remove('over'); }));
    document.addEventListener('drop', (e) => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) handleFile(f); });

    ['5', '7'].forEach((k) => $('#pay' + k).addEventListener('input', (e) => { state.pay[k] = parsePay(e.target.value); render(); }));
    $('#review').addEventListener('input', onReviewInput);
    $('#print').addEventListener('click', () => {
      const d = state.data;
      const prev = document.title;
      document.title = [displayClient(d), d.offerNo, 'Solar PV asset finance'].filter(Boolean).join(' ');
      window.print();
      setTimeout(() => { document.title = prev; }, 500);
    });
    $('#reset').addEventListener('click', () => {
      state.data = null; state.pay = { 5: null, 7: null };
      $('#pay5').value = ''; $('#pay7').value = '';
      setStatus('', ''); $('#status').className = 'status';
      renderReview(); render();
      $('#empty .big').textContent = 'Drop a PV*SOL report to start';
      $('#empty .msg').textContent = 'Word (.docx) or PDF. The file is read in your browser and never leaves your computer.';
    });
    window.addEventListener('resize', fitZoom);
    // Print from a clean copy of the pages directly under <body>; ids are renamed so the
    // copy's SVG gradient doesn't point at the hidden original.
    const printRoot = document.getElementById('print-root');
    window.addEventListener('beforeprint', () => { printRoot.innerHTML = $('#sheets').innerHTML.replace(/pvgrad/g, 'pvgrad-print'); });
    window.addEventListener('afterprint', () => { printRoot.innerHTML = ''; });
    render();
  }

  // Exposed for automated checks only
  window.__pvExplainer = { extractDocx, extractFields, model, state };
  document.addEventListener('DOMContentLoaded', init);
})();
