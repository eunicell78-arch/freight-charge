/* ═══════════════════════════════════════════════════════════════
   freight.js — 운임 계산 공통 엔진 (운임계산기 · 모의견적 · 견적서 공용)
   ───────────────────────────────────────────────────────────────
   · 단가는 freight_rates.xlsx 한 곳, 계산 방식은 이 파일 한 곳에서만 관리
     → freight-charge 저장소의 이 2개 파일만 고치면 3개 앱에 동시 반영
   · 항공: 청구중량 = max(G.W, CBM×167, 최소청구중량)
           CBM을 모르면 G.W ÷ 추정밀도(엑셀)로 추정
           항공운임 = max(구간단가×청구중량, 최소운임) + (FSC+SSC)×청구중량
           Collect Fee = max((운임+Origin)×2%, 최소)
   · 해상: CBM을 모르면 G.W ÷ 추정밀도(180), W/M = max(CBM, G.W÷1000)
   · DHL : ⑥DHL 시트 — G.W가 최대kg 이하인 첫 구간 요금
   · 엑셀을 못 불러오면 마지막으로 불러온 단가(이 기기 저장본)로 계산
   ENGINE_VERSION 은 계산 방식이 바뀔 때만 올림 (변경이력은 운임계산기 가이드에 기록)
   ═══════════════════════════════════════════════════════════════ */
(function (global) {
  'use strict';
  var ENGINE_VERSION = '1.1';   // 1.1 (2026-11-09) 간편표시 — routes({simple:true})
  var BASE = (function () {
    try { var s = document.currentScript && document.currentScript.src; if (s) return s.replace(/[^\/]*(\?.*)?$/, ''); } catch (e) {}
    return 'https://eunicell78-arch.github.io/freight-charge/';
  })();
  var XLSX_URL = BASE + 'freight_rates.xlsx';
  var XLSX_LIB = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  var CACHE_KEY = 'fe_rates_v1', FX_KEY = 'freightFX';
  var FX_FALLBACK = { USD: 1380, CNY: 190, HKD: 178, TWD: 43 };
  var VOL = 167;                         // 항공 부피중량 계수 (1 CBM = 167 kg)
  var BRACKETS = [
    { label: 'Minimum', rule: '최소 운임' },
    { label: '-45', rule: '45kg 미만' },
    { label: '+45', rule: '45kg 이상 ~ 100kg 미만' },
    { label: '+100', rule: '100kg 이상 ~ 300kg 미만' },
    { label: '+300', rule: '300kg 이상 ~ 500kg 미만' },
    { label: '+500', rule: '500kg 이상 ~ 1000kg 미만' },
    { label: '+1,000', rule: '1000kg 이상' }
  ];
  var st = { data: null, source: '', time: '', fx: null, fxTime: '', fxState: '' };

  /* ---------- helpers ---------- */
  function n(v) { if (v == null || v === '') return 0; var x = parseFloat(String(v).replace(/[^0-9.\-]/g, '')); return isNaN(x) ? 0 : x; }
  function s(v) { return v == null ? '' : String(v).trim(); }
  function r1(v) { return Math.round(v * 10) / 10; }
  function r2(v) { return Math.round(v * 100) / 100; }
  function nowTxt() { return new Date().toLocaleString('ko-KR', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }); }
  function findSheet(wb, kw) { return wb.SheetNames.find(function (x) { return x.indexOf(kw) >= 0; }); }
  function rows(wb, name) { if (!name) return []; var ws = wb.Sheets[name]; if (!ws) return []; return global.XLSX.utils.sheet_to_json(ws, { defval: null }); }
  function validCode(c) { return /^[A-Za-z0-9_]+$/.test(c); }
  function calcItem(it, qty) {
    if (it.type === 'fixed' || it.type === 'fixed_optional') return it.rate;
    if (it.type === 'perkg') return it.rate * qty;
    if (it.type === 'perkg_min' || it.type === 'percbm_min' || it.type === 'perwm_min') return Math.max(it.rate * qty, it.min || 0);
    return 0;
  }
  function bracketIndex(w) { if (w < 45) return 1; if (w < 100) return 2; if (w < 300) return 3; if (w < 500) return 4; if (w < 1000) return 5; return 6; }

  /* ---------- 엑셀 → 단가 데이터 ---------- */
  function parse(wb) {
    var d = { air: {}, airOrder: [], sea: {}, seaOrder: [], destAir: [], destSea: [], dhl: [] };
    var sAir = findSheet(wb, '항공운임'), sAirO = findSheet(wb, '항공Origin'), sSea = findSheet(wb, '해상운임'),
        sSeaO = findSheet(wb, '해상Origin'), sDest = findSheet(wb, 'Destination'), sDhl = findSheet(wb, 'DHL');
    rows(wb, sAir).forEach(function (r) {
      var code = s(r['노선코드']); if (!code || !validCode(code)) return;
      d.air[code] = {
        code: code, label: s(r['노선명']) || code, currency: s(r['통화']) || 'TWD', exchange: n(r['기본환율']) || 1,
        minimum: n(r['최소운임']), minChg: n(r['최소청구중량']), density: n(r['추정밀도']),
        air: [0, n(r['-45']), n(r['+45']), n(r['+100']), n(r['+300']), n(r['+500']), n(r['+1000'])],
        fsc: n(r['FSC']), ssc: n(r['SSC']), note: s(r['비고']), short: s(r['간편표시']), origin: []
      };
      d.airOrder.push(code);
    });
    rows(wb, sAirO).forEach(function (r) {
      var code = s(r['노선코드']), t = s(r['계산방식']); if (!code || !validCode(code) || t === 'none' || !d.air[code]) return;
      d.air[code].origin.push({ name: s(r['항목명']), type: t, rate: n(r['단가']), min: n(r['최소금액']) });
    });
    var cur = null;
    rows(wb, sSea).forEach(function (r) {
      var code = s(r['노선코드']); if (!code || !validCode(code)) return;
      if (s(r['노선명'])) {
        d.sea[code] = { code: code, label: s(r['노선명']) || code, originCurrency: s(r['현지통화']) || 'CNY', originExchange: n(r['현지환율']) || 1,
          freightCurrency: 'USD', freightExchange: n(r['USD환율']) || 1380, density: n(r['추정밀도']), note: s(r['비고']), short: s(r['간편표시']), origin: [], sea: [] };
        d.seaOrder.push(code); cur = code;
      }
      var key = cur || code;
      if (d.sea[key]) d.sea[key].sea.push({ name: s(r['항목명']), type: s(r['계산방식']) || 'percbm_min', rate: n(r['USD단가']), min: n(r['최소USD']) });
    });
    rows(wb, sSeaO).forEach(function (r) {
      var code = s(r['노선코드']), t = s(r['계산방식']); if (!code || !validCode(code) || t === 'none' || !d.sea[code]) return;
      d.sea[code].origin.push({ name: s(r['항목명']), type: t, rate: n(r['단가']), min: n(r['최소금액']) });
    });
    rows(wb, sDest).forEach(function (r) {
      var div = s(r['구분']); if (!div) return;
      var it = { name: s(r['항목명']), type: s(r['계산방식']), rate: n(r['금액/비율']), min: n(r['최소금액']), note: s(r['비고']) };
      if (div === '항공') d.destAir.push(it); else if (div === '해상') d.destSea.push(it);
    });
    rows(wb, sDhl).forEach(function (r) {
      var mx = n(r['최대kg']), rate = n(r['요금KRW']); if (!(mx > 0) || !(rate > 0)) return;
      d.dhl.push({ label: s(r['구간']) || ('~' + mx + 'kg'), maxKg: mx, rate: rate });
    });
    d.dhl.sort(function (a, b) { return a.maxKg - b.maxKg; });
    return d;
  }
  function isValid(d) { return d && (Object.keys(d.air || {}).length || Object.keys(d.sea || {}).length); }
  function setData(d, source) {
    st.data = d; st.source = source || ''; st.time = nowTxt();
    if (source !== 'cache') { try { localStorage.setItem(CACHE_KEY, JSON.stringify({ time: st.time, data: d })); } catch (e) {} }
    return d;
  }
  function setWorkbook(wb, source) { var d = parse(wb); if (!isValid(d)) throw new Error('인식 가능한 시트 없음'); return setData(d, source || 'file'); }

  function ensureXLSX() {
    if (global.XLSX && global.XLSX.read) return Promise.resolve();
    return new Promise(function (res, rej) {
      var sc = document.createElement('script'); sc.src = XLSX_LIB;
      sc.onload = function () { res(); }; sc.onerror = function () { rej(new Error('엑셀 라이브러리 로드 실패')); };
      document.head.appendChild(sc);
    });
  }
  function loadCache() {
    try { var c = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null'); if (c && isValid(c.data)) { st.data = c.data; st.source = 'cache'; st.time = c.time || ''; return true; } } catch (e) {}
    return false;
  }
  /* load() → Promise<{data, source:'live'|'cache', time}> */
  function load(opts) {
    opts = opts || {};
    var url = opts.url || XLSX_URL;
    return ensureXLSX()
      .then(function () { return fetch(url + (url.indexOf('?') < 0 ? '?t=' + Date.now() : ''), { cache: 'no-store' }); })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
      .then(function (buf) { var wb = global.XLSX.read(new Uint8Array(buf), { type: 'array' }); setWorkbook(wb, 'live'); return { data: st.data, source: 'live', time: st.time }; })
      .catch(function (e) {
        if (loadCache()) return { data: st.data, source: 'cache', time: st.time, error: e.message };
        throw e;
      });
  }

  /* ---------- 실시간 환율 (KRW per 1 unit) — 3중 API 폴백 + 캐시 ---------- */
  function fxFromUsdBase(rates) {
    if (!rates) return null;
    var get = function (k) { return rates[k] != null ? rates[k] : rates[k.toLowerCase()]; };
    var krw = get('KRW'); if (!krw) return null;
    var cny = get('CNY'), hkd = get('HKD'), twd = get('TWD'); if (!cny || !hkd || !twd) return null;
    return { USD: krw, CNY: krw / cny, HKD: krw / hkd, TWD: krw / twd };
  }
  function getJSON(u) { return fetch(u).then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); }); }
  var FX_TRIES = [
    function () { return getJSON('https://open.er-api.com/v6/latest/USD').then(function (j) { return fxFromUsdBase(j.rates); }); },
    function () { return getJSON('https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/usd.json').then(function (j) { var u = j.usd || {}; return fxFromUsdBase({ KRW: u.krw, CNY: u.cny, HKD: u.hkd, TWD: u.twd }); }); },
    function () { return getJSON('https://api.frankfurter.app/latest?from=USD&to=KRW,CNY,HKD,TWD').then(function (j) { return fxFromUsdBase(j.rates); }); }
  ];
  /* loadFX(force, onUpdate(state)) — state: '캐시' | '실시간' | '실패' */
  function loadFX(force, onUpdate) {
    var cb = typeof onUpdate === 'function' ? onUpdate : function () {};
    if (!force) { try { var c = JSON.parse(localStorage.getItem(FX_KEY) || 'null'); if (c && c.rates) { st.fx = c.rates; st.fxTime = c.time; st.fxState = '캐시'; cb('캐시'); } } catch (e) {} }
    var i = 0;
    function next() {
      if (i >= FX_TRIES.length) {
        if (!st.fx) { st.fx = Object.assign({}, FX_FALLBACK); st.fxState = '실패'; cb('실패'); }
        return Promise.resolve(st.fx);
      }
      var run = FX_TRIES[i++];
      return run().then(function (r) {
        if (!r) return next();
        st.fx = r; st.fxTime = nowTxt(); st.fxState = '실시간';
        try { localStorage.setItem(FX_KEY, JSON.stringify({ rates: r, time: st.fxTime })); } catch (e) {}
        cb('실시간'); return r;
      }).catch(function () { return next(); });
    }
    return next();
  }
  function fxRate(ccy) { if (ccy === 'KRW') return 1; return (st.fx && st.fx[ccy]) ? st.fx[ccy] : null; }

  /* ---------- 계산 ---------- */
  /* calcAir(preset, destList, {gw, cbm, ex}) — cbm 비우면 G.W÷추정밀도로 추정 */
  function calcAir(p, dest, inp) {
    inp = inp || {}; dest = dest || [];
    var gw = n(inp.gw), cbmIn = n(inp.cbm), ex = n(inp.ex) || p.exchange || 1;
    var est = !(cbmIn > 0);
    var cbm = est ? (p.density > 0 && gw > 0 ? gw / p.density : 0) : cbmIn;
    var volW = r1(cbm * VOL), minChg = p.minChg || 0;
    var chg = gw > 0 ? Math.max(gw, volW, minChg) : 0;
    var basis = !chg ? '' : (chg === gw ? 'G.W' : (chg === volW ? '부피중량' : '최소청구중량'));
    var bi = bracketIndex(chg), rate = (p.air && p.air[bi]) || 0;
    var afRaw = rate * chg, min = p.minimum || 0, af = Math.max(afRaw, min);
    var fsc = (p.fsc || 0) * chg, ssc = (p.ssc || 0) * chg;
    var freight = chg ? af + fsc + ssc : 0, freightKrw = freight * ex;
    var origin = (p.origin || []).map(function (it) { var a = chg ? calcItem(it, chg) : 0; return { name: it.name, type: it.type, rate: it.rate, min: it.min, amt: a, krw: a * ex }; });
    var originFx = origin.reduce(function (t, o) { return t + o.amt; }, 0), originKrw = originFx * ex;
    var destOut = dest.map(function (d) {
      var a = !chg ? 0 : (d.type === 'pct_freight_min' ? Math.max((freightKrw + originKrw) * d.rate, d.min || 0) : (d.type === 'fixed' || d.type === 'fixed_optional' ? d.rate : 0));
      return { name: d.name, type: d.type, rate: d.rate, min: d.min, amt: a };
    });
    var destKrw = destOut.reduce(function (t, o) { return t + o.amt; }, 0);
    var total = Math.round(originKrw + freightKrw + destKrw);
    return { mode: 'air', code: p.code, label: p.label, currency: p.currency, ex: ex, gw: gw, cbm: r2(cbm), cbmEstimated: est && cbm > 0, density: p.density,
      volW: volW, chg: chg, basis: basis, bracket: bi, bracketLabel: BRACKETS[bi].label, rate: rate, af: af, minApplied: min > 0 && afRaw < min,
      fsc: fsc, ssc: ssc, freight: freight, freightKrw: freightKrw, origin: origin, originFx: originFx, originKrw: originKrw,
      dest: destOut, destKrw: destKrw, total: total, unit: chg ? Math.round(total / chg) : 0, unitLabel: 'kg' };
  }
  /* calcSea(preset, destList, {gw, cbm, wm, ex, usd, duty}) — cbm 비우면 G.W÷추정밀도, wm 비우면 max(CBM, G.W÷1000) */
  function calcSea(p, dest, inp) {
    inp = inp || {}; dest = dest || [];
    var gw = n(inp.gw), cbmIn = n(inp.cbm), wmIn = n(inp.wm);
    var ex = n(inp.ex) || p.originExchange || 1, usd = n(inp.usd) || p.freightExchange || 1380, duty = n(inp.duty);
    var est = !(cbmIn > 0);
    var cbm = est ? (p.density > 0 && gw > 0 ? gw / p.density : 0) : cbmIn;
    cbm = r2(cbm);
    var wm = wmIn > 0 ? wmIn : r2(Math.max(cbm, gw / 1000));
    var on = cbm > 0 || wm > 0;
    var origin = (p.origin || []).map(function (it) { var a = !on ? 0 : calcItem(it, it.type === 'perwm_min' ? wm : 1); return { name: it.name, type: it.type, rate: it.rate, min: it.min, amt: a, krw: a * ex }; });
    var originFx = origin.reduce(function (t, o) { return t + o.amt; }, 0), originKrw = originFx * ex;
    var sea = (p.sea || []).map(function (it) { var a = !on ? 0 : calcItem(it, cbm); return { name: it.name, type: it.type, rate: it.rate, min: it.min, amt: a, krw: a * usd }; });
    var seaUsd = sea.reduce(function (t, o) { return t + o.amt; }, 0), freightKrw = seaUsd * usd;
    var destOut = dest.map(function (d) {
      var a = !on ? 0 : (d.type === 'perwm_min' ? Math.max(d.rate * wm, d.min || 0) : (d.type === 'pct_duty' ? duty * d.rate : (d.type === 'pct_freight_min' ? Math.max(freightKrw * d.rate, d.min || 0) : d.rate)));
      return { name: d.name, type: d.type, rate: d.rate, min: d.min, amt: a };
    });
    var destKrw = destOut.reduce(function (t, o) { return t + o.amt; }, 0);
    var total = Math.round(originKrw + freightKrw + destKrw);
    return { mode: 'sea', code: p.code, label: p.label, currency: p.originCurrency, ex: ex, usd: usd, gw: gw, cbm: cbm, cbmEstimated: est && cbm > 0, density: p.density,
      wm: wm, wmAuto: !(wmIn > 0), origin: origin, originFx: originFx, originKrw: originKrw, sea: sea, seaUsd: seaUsd, freightKrw: freightKrw,
      dest: destOut, destKrw: destKrw, total: total, unit: cbm ? Math.round(total / cbm) : 0, unitLabel: 'CBM' };
  }
  /* calcDhl(gw, list) — G.W가 최대kg 이하인 첫 구간 */
  function calcDhl(gw, list) {
    list = list || (st.data && st.data.dhl) || []; gw = n(gw);
    var idx = -1;
    for (var i = 0; i < list.length; i++) { if (gw > 0 && gw <= list[i].maxKg) { idx = i; break; } }
    var over = gw > 0 && idx < 0 && list.length > 0;
    var b = idx >= 0 ? list[idx] : null;
    return { mode: 'dhl', gw: gw, index: idx, bracketLabel: b ? b.label : '', total: b ? b.rate : 0, over: over,
      maxKg: list.length ? list[list.length - 1].maxKg : 0, unit: b && gw ? Math.round(b.rate / gw) : 0, unitLabel: 'kg' };
  }

  /* ---------- 앱용 간편 함수 ---------- */
  /* routes(opts) → [{key:'dhl'|'air:HKG'|'sea:CNSZX', mode, code, label}]
     opts.simple=true → 엑셀 '간편표시'가 있는 노선만, 그 이름으로 (모의견적용). 간편표시가 하나도 없으면 전체 */
  function routes(opts) {
    var d = st.data, out = []; if (!d) return out;
    var simple = !!(opts && opts.simple) && (d.airOrder.some(function (c) { return d.air[c].short; }) || d.seaOrder.some(function (c) { return d.sea[c].short; }));
    if (d.dhl && d.dhl.length) out.push({ key: 'dhl', mode: 'dhl', code: '', label: 'DHL 특송' });
    d.airOrder.forEach(function (c) { var p = d.air[c]; if (simple && !p.short) return; out.push({ key: 'air:' + c, mode: 'air', code: c, label: simple ? p.short : '항공 · ' + p.label }); });
    d.seaOrder.forEach(function (c) { var p = d.sea[c]; if (simple && !p.short) return; out.push({ key: 'sea:' + c, mode: 'sea', code: c, label: simple ? p.short : '해상 · ' + p.label }); });
    return out;
  }
  /* quote(key, {gw, cbm, ex?, usd?, duty?}) — 환율 비우면 실시간 → 엑셀 기본환율 */
  function quote(key, inp) {
    var d = st.data; if (!d || !key) return null; inp = inp || {};
    var parts = String(key).split(':'), mode = parts[0], code = parts[1];
    if (mode === 'dhl') return calcDhl(inp.gw, d.dhl);
    if (mode === 'air' && d.air[code]) { var p = d.air[code]; return calcAir(p, d.destAir, { gw: inp.gw, cbm: inp.cbm, ex: n(inp.ex) || fxRate(p.currency) || p.exchange }); }
    if (mode === 'sea' && d.sea[code]) { var q = d.sea[code]; return calcSea(q, d.destSea, { gw: inp.gw, cbm: inp.cbm, wm: inp.wm, duty: inp.duty, ex: n(inp.ex) || fxRate(q.originCurrency) || q.originExchange, usd: n(inp.usd) || fxRate('USD') || q.freightExchange }); }
    return null;
  }
  function fmt(v, dg) { return Number(v || 0).toLocaleString('ko-KR', { maximumFractionDigits: dg == null ? 0 : dg }); }
  /* describe(result) → 한 줄 설명 (화면 표시용) */
  function describe(r) {
    if (!r) return '';
    if (r.mode === 'dhl') {
      if (!r.gw) return 'G.W(kg)를 입력하세요';
      if (r.over) return 'DHL 요율표는 ' + r.maxKg + 'kg까지 — 초과는 포워더 항공으로 계산하세요';
      return 'DHL ' + r.bracketLabel + ' 구간 적용';
    }
    if (r.mode === 'air') {
      if (!r.chg) return 'G.W(kg)를 입력하세요';
      var t = '청구중량 ' + fmt(r.chg, 1) + 'kg (' + r.basis + ' 적용';
      if (r.cbm > 0) t += ' · ' + (r.cbmEstimated ? '추정 ' : '') + 'CBM ' + fmt(r.cbm, 2) + (r.cbmEstimated ? ' = G.W÷' + r.density : '');
      return t + ') · ' + r.bracketLabel + ' 구간';
    }
    if (r.mode === 'sea') {
      if (!r.cbm && !r.wm) return 'G.W(kg) 또는 CBM을 입력하세요';
      return (r.cbmEstimated ? '추정 ' : '') + 'CBM ' + fmt(r.cbm, 2) + (r.cbmEstimated ? ' (G.W÷' + r.density + ')' : '') + ' · W/M ' + fmt(r.wm, 2) + (r.wmAuto ? ' (자동)' : '');
    }
    return '';
  }

  global.FreightEngine = {
    VERSION: ENGINE_VERSION, XLSX_URL: XLSX_URL, VOL_FACTOR: VOL, BRACKETS: BRACKETS,
    load: load, parse: parse, setWorkbook: setWorkbook, loadCache: loadCache,
    data: function () { return st.data; }, source: function () { return st.source; }, loadedAt: function () { return st.time; },
    loadFX: loadFX, fxRate: fxRate, getFX: function () { return st.fx; }, fxTime: function () { return st.fxTime; }, fxState: function () { return st.fxState; },
    calcAir: calcAir, calcSea: calcSea, calcDhl: calcDhl, calcItem: calcItem, bracketIndex: bracketIndex,
    routes: routes, quote: quote, describe: describe
  };
})(window);
