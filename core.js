/* core.js — Motor comum (CSV + Meteo + Cálculo + Estado) */
/* global window, fetch, AbortController */

(() => {
  "use strict";

  const STORAGE_KEY = "mdp_bovinos_state_v2";
  const METEO_CACHE_KEY = "mdp_bovinos_meteo_cache_v2";

  const DEFAULTS = {
    SITE: { lat: 38.17355612872988, lon: -7.986520046258665 },
    CONSERVATIVE_FALLBACK_GMD: 1.10,
    METEO_TIMEOUT_MS: 9000,
    METEO_CONCURRENCY: 6
  };

  /* ===================== HELPERS ===================== */
  function clean(s){ return String(s ?? "").trim(); }
  function escapeHtml(str){
    return String(str ?? "").replace(/[&<>"']/g, m => ({
      "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"
    }[m]));
  }
  function parseNumber(x){
    x = clean(x);
    if(!x) return NaN;
    x = x.replace(",", ".");
    const n = parseFloat(x);
    return Number.isFinite(n) ? n : NaN;
  }
  function parseDatePT(s){
    s = clean(s);
    if(!s) return null;

    // YYYY-MM-DD
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if(m){
      const y=+m[1], mo=+m[2], d=+m[3];
      const dt = new Date(Date.UTC(y, mo-1, d));
      return Number.isFinite(dt.getTime()) ? dt : null;
    }

    // DD-MM-YYYY or DD/MM/YYYY
    m = s.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})$/);
    if(m){
      const d=+m[1], mo=+m[2], y=+m[3];
      const dt = new Date(Date.UTC(y, mo-1, d));
      return Number.isFinite(dt.getTime()) ? dt : null;
    }

    return null;
  }
  function fmtDate(d){
    if(!d) return "—";
    const dd = String(d.getUTCDate()).padStart(2,"0");
    const mm = String(d.getUTCMonth()+1).padStart(2,"0");
    const yy = d.getUTCFullYear();
    return `${dd}-${mm}-${yy}`;
  }
  function isoDateUTC(d){
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth()+1).padStart(2,"0");
    const da = String(d.getUTCDate()).padStart(2,"0");
    return `${y}-${m}-${da}`;
  }
  function addDaysUTC(dateUTC, days){
    const ms = dateUTC.getTime() + (days * 24*60*60*1000);
    return new Date(ms);
  }
  function daysBetweenUTC(a,b){
    if(!a || !b) return NaN;
    const ms = b.getTime() - a.getTime();
    return Math.floor(ms / (1000*60*60*24));
  }
  function detectDelimiter(text){
  const first = (text.split(/\r?\n/)[0] || "");

  const counts = {
    ";": (first.match(/;/g) || []).length,
    ",": (first.match(/,/g) || []).length,
    "\t": (first.match(/\t/g) || []).length
  };

  const sorted = Object.entries(counts)
    .sort((a,b) => b[1] - a[1]);

  return sorted[0][1] > 0 ? sorted[0][0] : ";";
}
function safeFloat(x, digits=2){
  return Number.isFinite(x) ? x.toFixed(digits) : "—";
}

function safeInt(x){
  return Number.isFinite(x) ? String(Math.max(0, Math.ceil(x))) : "—";
}
  /* ===================== REGRAS (iguais à base estável) ===================== */
  function performanceStatus(gmdInd, gmdMediaGrupo){
    if(!Number.isFinite(gmdInd) || !Number.isFinite(gmdMediaGrupo) || gmdMediaGrupo<=0) return ["— (sem histórico)","muted", 9, "none"];
    const r = gmdInd / gmdMediaGrupo;
    if(r >= 0.95) return ["🟢 Normal","ok", 3, "g"];
    if(r >= 0.80) return ["🟡 A vigiar","warn", 2, "o"];
    return ["🔴 Atrasado","bad", 1, "r"];
  }
  function confidenceByDays(days){
    if(!Number.isFinite(days)) return ["—","muted"];
    if(days < 14) return ["Alta","ok"];
    if(days < 35) return ["Média","warn"];
    return ["Baixa","bad"];
  }
  function factorSexo(sexo){ return (clean(sexo).toUpperCase()==="F") ? 0.92 : 1.00; }
  function factorMaturidade(peso, sexo){
    const s = clean(sexo).toUpperCase();
    const limiar = (s==="F") ? 460 : 520;
    if(peso < limiar) return 1.00;
    if(peso < limiar + 60) return 0.90;
    return 0.80;
  }

  /* ===================== METEO (Open-Meteo histórico) ===================== */
  function factorFromTempMean(t){
    if(!Number.isFinite(t)) return 0.95;
    if(t <= 20) return 1.00;
    if(t <= 25) return 0.95;
    if(t <= 30) return 0.85;
    return 0.70;
  }

  function fetchWithTimeout(url, timeoutMs){
    const controller = new AbortController();
    const t = setTimeout(()=>controller.abort(), timeoutMs);
    return fetch(url, { signal: controller.signal })
      .finally(()=>clearTimeout(t));
  }

  function loadMeteoCache(){
    try{
      const raw = localStorage.getItem(METEO_CACHE_KEY);
      if(!raw) return new Map();
      const obj = JSON.parse(raw);
      const m = new Map();
      for(const [k,v] of Object.entries(obj || {})){
        if(v && typeof v === "object") m.set(k, v);
      }
      return m;
    }catch{
      return new Map();
    }
  }

  function saveMeteoCache(map){
    try{
      const obj = {};
      for(const [k,v] of map.entries()){
        obj[k] = v;
      }
      localStorage.setItem(METEO_CACHE_KEY, JSON.stringify(obj));
    }catch{
      /* ignore */
    }
  }

  async function getTempMeanForPeriod(site, meteoCacheMap, startUTC, endUTC, timeoutMs){
    const start = isoDateUTC(startUTC);
    const end = isoDateUTC(endUTC);
    const key = `${start}|${end}`;
    if(meteoCacheMap.has(key)) return meteoCacheMap.get(key);

    const url =
      `https://archive-api.open-meteo.com/v1/archive` +
      `?latitude=${site.lat}&longitude=${site.lon}` +
      `&start_date=${start}&end_date=${end}` +
      `&daily=temperature_2m_mean` +
      `&timezone=auto`;

    try{
      const r = await fetchWithTimeout(url, timeoutMs);
      if(!r.ok) throw new Error(`HTTP ${r.status}`);
      const j = await r.json();
      const temps = j?.daily?.temperature_2m_mean || [];
      if(!temps.length){
        const pack = { tmean: NaN, factor: 0.95 };
        meteoCacheMap.set(key, pack);
        return pack;
      }
      const tmean = temps.reduce((a,b)=>a+b,0)/temps.length;
      const pack = { tmean, factor: factorFromTempMean(tmean) };
      meteoCacheMap.set(key, pack);
      return pack;
    }catch{
      const pack = { tmean: NaN, factor: 0.95 };
      meteoCacheMap.set(key, pack);
      return pack;
    }
  }

  async function preloadMeteoPeriods(cfg, site, meteoCacheMap, periodKeys, onProgress){
    const keys = Array.from(periodKeys);
    let done = 0;
    const total = keys.length;

    const worker = async () => {
      while(keys.length){
        const key = keys.shift();
        const [s,e] = key.split("|");
        const sD = new Date(Date.UTC(+s.slice(0,4), +s.slice(5,7)-1, +s.slice(8,10)));
        const eD = new Date(Date.UTC(+e.slice(0,4), +e.slice(5,7)-1, +e.slice(8,10)));
        await getTempMeanForPeriod(site, meteoCacheMap, sD, eD, cfg.METEO_TIMEOUT_MS);
        done++;
        if(onProgress) onProgress(done, total);
      }
    };

    const workers = [];
    const n = Math.max(1, Math.min(cfg.METEO_CONCURRENCY, total || 1));
    for(let i=0;i<n;i++) workers.push(worker());
    await Promise.all(workers);
  }

  /* ===================== FORECAST HELPERS ===================== */
  function pickGmdUsed(realGmd, fallback){
    return Number.isFinite(realGmd) && realGmd>0 ? realGmd : fallback;
  }
  function classifyReady(pesoEst, alvo){
    if(!Number.isFinite(pesoEst) || !Number.isFinite(alvo)) return ["—","muted"];
    if(pesoEst >= alvo) return ["Pronto ✅","ok"];
    if(pesoEst >= alvo - 20) return ["Quase lá 🟡","warn"];
    return ["Em engorda","muted"];
  }
  function calcDaysToTarget(pesoEst, alvo, gmd){
    if(!Number.isFinite(pesoEst) || !Number.isFinite(alvo) || !Number.isFinite(gmd) || gmd<=0) return NaN;
    if(pesoEst >= alvo) return 0;
    return (alvo - pesoEst) / gmd;
  }

  /* ===================== STATE (localStorage) ===================== */
  function getEmptyState(){
    return {
      version: 2,
      generated_at: null,
      config: {
        SITE: { ...DEFAULTS.SITE },
        CONSERVATIVE_FALLBACK_GMD: DEFAULTS.CONSERVATIVE_FALLBACK_GMD
      },
      meta: {
        delimiter: ";",
        lines: 0,
        processed_ok: 0,
        processed_fail: 0
      },
      todayUTC: null,

      // outputs
      animalsOut: [],
      groupsOut: [],
      groupAgg: {},
      gmdEstimativaGrupo: {},
      gmdMediaGrupo: {},

      // remember last targets for convenience (UI)
      targets: { targetM: 620, targetF: 520 }
    };
  }

  function loadState(){
    try{
      const raw = localStorage.getItem(STORAGE_KEY);
      if(!raw) return null;
      const st = JSON.parse(raw);
      if(!st || typeof st !== "object") return null;
      return st;
    }catch{
      return null;
    }
  }

  function saveState(state){
    try{
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      return true;
    }catch{
      return false;
    }
  }

  function clearState(){
    localStorage.removeItem(STORAGE_KEY);
  }

  function ensureState(){
    return loadState() || getEmptyState();
  }

  /* ===================== CSV PROCESS (igual à base estável, mas “headless”) ===================== */
  async function processCSVText(csvText, opts = {}){
    const cfg = {
      SITE: opts.SITE || DEFAULTS.SITE,
      CONSERVATIVE_FALLBACK_GMD: Number.isFinite(opts.CONSERVATIVE_FALLBACK_GMD) ? opts.CONSERVATIVE_FALLBACK_GMD : DEFAULTS.CONSERVATIVE_FALLBACK_GMD,
      METEO_TIMEOUT_MS: Number.isFinite(opts.METEO_TIMEOUT_MS) ? opts.METEO_TIMEOUT_MS : DEFAULTS.METEO_TIMEOUT_MS,
      METEO_CONCURRENCY: Number.isFinite(opts.METEO_CONCURRENCY) ? opts.METEO_CONCURRENCY : DEFAULTS.METEO_CONCURRENCY
    };

    const onProgress = (typeof opts.onProgress === "function") ? opts.onProgress : null;

    const st = getEmptyState();
    st.config.SITE = { ...cfg.SITE };
    st.config.CONSERVATIVE_FALLBACK_GMD = cfg.CONSERVATIVE_FALLBACK_GMD;

    const delim = detectDelimiter(csvText);
    st.meta.delimiter = delim;

    const lines = csvText.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
    if(lines.length < 2){
      throw new Error("CSV vazio ou inválido.");
    }
    st.meta.lines = lines.length - 1;

    const now = new Date();
    const todayUTC = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    st.todayUTC = todayUTC.toISOString();

    const header = lines[0].split(delim).map(clean);
    const idx = (name) => header.indexOf(name);

    const i_animal = idx("animal_id") >= 0 ? idx("animal_id") : 0;
    const i_sexo   = idx("sexo") >= 0 ? idx("sexo") : 2;
    const i_grupo  = idx("grupo") >= 0 ? idx("grupo") : 3;
    const i_dant   = idx("data_peso_anterior") >= 0 ? idx("data_peso_anterior") : 4;
    const i_pant   = idx("peso_anterior") >= 0 ? idx("peso_anterior") : 5;
    const i_datual = idx("data_peso_atual") >= 0 ? idx("data_peso_atual") : 6;
    const i_patual = idx("peso_atual") >= 0 ? idx("peso_atual") : 7;
    const i_nasc = idx("data_nasc") >= 0 ? idx("data_nasc") : -1;

    if(onProgress) onProgress({ phase:"parse", message:`A ler dados…`, done:0, total:st.meta.lines });

    const rows = [];
    const gmdIndSamples = {}; // grupo -> [gmdInd]
    const periodKeys = new Set();

    for(let li=1; li<lines.length; li++){
  const cols = lines[li].split(delim);

  const animal = clean(cols[i_animal]) || "—";
  const sexo   = clean(cols[i_sexo]).toUpperCase() || "—";
  const grupo  = clean(cols[i_grupo]) || "—";

  // 👇 AQUI
  const dNasc = i_nasc >= 0 ? parseDatePT(cols[i_nasc]) : null;

  const dAnt   = parseDatePT(cols[i_dant]);
  const pAnt   = parseNumber(cols[i_pant]);
  const dAtual = parseDatePT(cols[i_datual]);
  const pAtual = parseNumber(cols[i_patual]);

      let gmdInd = NaN;
      if(dAnt && dAtual && Number.isFinite(pAnt) && Number.isFinite(pAtual)){
        const d = daysBetweenUTC(dAnt, dAtual);
        if(Number.isFinite(d) && d > 0){
          gmdInd = (pAtual - pAnt) / d;
          if(Number.isFinite(gmdInd)){
            (gmdIndSamples[grupo] ||= []).push(gmdInd);
          }
        }
      }

      if(dAtual){
        periodKeys.add(`${isoDateUTC(dAtual)}|${isoDateUTC(todayUTC)}`);
      }

      rows.push({animal, sexo, grupo, dNasc, dAnt, pAnt, dAtual, pAtual, gmdInd});

      if(onProgress && li % 200 === 0){
        onProgress({ phase:"parse", message:`A ler dados…`, done:li-1, total:st.meta.lines });
      }
    }

    // meteo cache (persistente entre runs)
    const meteoCacheMap = loadMeteoCache();

    if(onProgress) onProgress({ phase:"meteo", message:`A pedir meteorologia…`, done:0, total:periodKeys.size });

    await preloadMeteoPeriods(cfg, cfg.SITE, meteoCacheMap, periodKeys, (done,total)=>{
      if(onProgress) onProgress({ phase:"meteo", message:`A pedir meteorologia…`, done, total });
    });

    saveMeteoCache(meteoCacheMap);

    // médias reais por grupo (misto) para estado
    const gmdMediaGrupo = {};
    for(const r of rows){
      const g = r.grupo || "—";
      if(gmdMediaGrupo[g] == null){
        const arr = gmdIndSamples[g] || [];
        gmdMediaGrupo[g] = arr.length ? (arr.reduce((a,b)=>a+b,0) / arr.length) : NaN;
      }
    }
    st.gmdMediaGrupo = gmdMediaGrupo;

    // gmd estimativa por grupo (misto)
    const gmdEstimativaGrupo = {};
    for(const r of rows){
      const g = r.grupo || "—";
      if(gmdEstimativaGrupo[g] == null){
        const arr = gmdIndSamples[g] || [];
        gmdEstimativaGrupo[g] = arr.length ? (arr.reduce((a,b)=>a+b,0) / arr.length) : cfg.CONSERVATIVE_FALLBACK_GMD;
      }
    }
    st.gmdEstimativaGrupo = gmdEstimativaGrupo;

    const animalsOut = [];
    const groupAgg = {};
    let okRows=0, badRows=0;

    for(const r of rows){
      const hasAtual = r.dAtual && Number.isFinite(r.pAtual);
      if(!hasAtual){
        badRows++;
        animalsOut.push({
  sortKey: 99,
  animal:r.animal, grupo:r.grupo, sexo:r.sexo,
  pesoAtual:"—", dataAtual:"—",
  temp:"—", fatorClima: NaN,
  estimado:"—",
  dmi: "—",
  fcr: "—",
  eficiencia: "—",
  conf:"—", confClass:"muted",
  estado:"—", estadoClass:"muted",
  bucket:"none",
  estKg: NaN
});
        continue;
      }

      const today = todayUTC;
      const daysSince = daysBetweenUTC(r.dAtual, today);
      const [conf, confClass] = confidenceByDays(daysSince);

      const meteo = await getTempMeanForPeriod(cfg.SITE, meteoCacheMap, r.dAtual, today, cfg.METEO_TIMEOUT_MS);
      const fc = meteo.factor;

      const gmdBase = Number.isFinite(r.gmdInd) ? r.gmdInd : (gmdEstimativaGrupo[r.grupo] ?? cfg.CONSERVATIVE_FALLBACK_GMD);
      const gmdFinal = gmdBase * factorSexo(r.sexo) * factorMaturidade(r.pAtual, r.sexo) * fc;
      const estKg = r.pAtual + (gmdFinal * daysSince);
// ===== ALIMENTAÇÃO INTELIGENTE =====

// % base por peso
let percBase = 0.025;
if (r.pAtual < 300) percBase = 0.028;
else if (r.pAtual > 500) percBase = 0.022;

// sexo
const sexoNorm = clean(r.sexo).toUpperCase();
const factorSexoAlim = (sexoNorm === "F") ? 1.02 : 0.98;

// idade
let factorIdade = 1.0;
if (r.dNasc) {
  const idadeDias = daysBetweenUTC(r.dNasc, today);
  if (idadeDias < 300) factorIdade = 1.05;
  else if (idadeDias > 600) factorIdade = 0.95;
}

// clima (já tens)
const factorClima = fc;

// DMI estimado
const dmi = r.pAtual * percBase * factorSexoAlim * factorIdade * factorClima;

// FCR estimado
const fcr = (gmdFinal > 0) ? dmi / gmdFinal : NaN;
      let eficiencia = "—";
if (Number.isFinite(fcr)) {
  if (fcr < 6) eficiencia = "🟢 Excelente";
  else if (fcr < 7.5) eficiencia = "🟡 Normal";
  else eficiencia = "🔴 Ineficiente";
}
      const [estado, estadoClass, estadoSort, bucket] = performanceStatus(r.gmdInd, gmdMediaGrupo[r.grupo]);

      okRows++;

      animalsOut.push({
        sortKey: estadoSort,
        animal:r.animal, grupo:r.grupo, sexo:r.sexo,
        pesoAtual:`${r.pAtual.toFixed(1)} kg`,
        dataAtual:fmtDate(r.dAtual),
        temp: Number.isFinite(meteo.tmean) ? `${meteo.tmean.toFixed(1)} °C` : "—",
        fatorClima: fc,
        estimado: Number.isFinite(estKg) ? `${estKg.toFixed(1)} kg` : "—",
        dmi: Number.isFinite(dmi) ? `${dmi.toFixed(1)} kg/dia` : "—",
fcr: Number.isFinite(fcr) ? fcr.toFixed(2) : "—",
eficiencia,
        conf, confClass,
        estado, estadoClass,
        bucket,
        estKg
      });

      const g = r.grupo || "—";
if(!groupAgg[g]){
  groupAgg[g] = {
    name:g,n:0,m:0,f:0,
    sumPesoM:0,sumPesoF:0,
    sumEstM:0,sumEstF:0,
    sumGmdM:0,sumGmdF:0,
    nGmdM:0,nGmdF:0,
    sumTemp:0,nTemp:0,

    // 🔥 NOVO MODELO CORRETO
    sumDmi:0,
    sumGanho:0,

    ok:0,warn:0,bad:0
  };
}

      const ga = groupAgg[g];
ga.n++;


if (Number.isFinite(fcr)) {
  ga.sumFcr += fcr;
  ga.nFcr++;
}

const sx = clean(r.sexo).toUpperCase();
      if(sx==="M"){
        ga.m++; ga.sumPesoM += r.pAtual;
        if(Number.isFinite(estKg)) ga.sumEstM += estKg;
      }else if(sx==="F"){
        ga.f++; ga.sumPesoF += r.pAtual;
        if(Number.isFinite(estKg)) ga.sumEstF += estKg;
      }

      if(Number.isFinite(meteo.tmean)){ ga.sumTemp += meteo.tmean; ga.nTemp++; }

      if(Number.isFinite(r.gmdInd)){
        if(sx==="M"){ ga.sumGmdM += r.gmdInd; ga.nGmdM++; }
        else if(sx==="F"){ ga.sumGmdF += r.gmdInd; ga.nGmdF++; }

        if(bucket==="g") ga.ok++;
        else if(bucket==="o") ga.warn++;
        else if(bucket==="r") ga.bad++;
      }
    }

    // ordenar animais (🔴, 🟡, 🟢, sem histórico)
    animalsOut.sort((a,b)=>a.sortKey-b.sortKey || a.grupo.localeCompare(b.grupo) || a.animal.localeCompare(b.animal));

    // gruposOut
    const groupsOut = Object.values(groupAgg).map(g=>{
      g.avgPesoM = g.m ? g.sumPesoM/g.m : NaN;
      g.avgPesoF = g.f ? g.sumPesoF/g.f : NaN;
      g.avgEstM  = g.m ? g.sumEstM/g.m : NaN;
      g.avgEstF  = g.f ? g.sumEstF/g.f : NaN;

      g.avgGmdM  = g.nGmdM ? g.sumGmdM/g.nGmdM : NaN;
      g.avgGmdF  = g.nGmdF ? g.sumGmdF/g.nGmdF : NaN;
      g.avgTemp  = g.nTemp ? g.sumTemp/g.nTemp : NaN;
      // 🔥 DMI médio (mantém)
g.avgDmi = g.n ? g.sumDmi / g.n : NaN;

// 🔥 FCR CORRETO (não é média!)
g.avgFcr = (g.sumGanho > 0) ? (g.sumDmi / g.sumGanho) : NaN;

// eficiência
let eficienciaGrupo = "—";
if (Number.isFinite(g.avgFcr)) {
  if (g.avgFcr < 6) eficienciaGrupo = "🟢 Excelente";
  else if (g.avgFcr < 7.5) eficienciaGrupo = "🟡 Normal";
  else eficienciaGrupo = "🔴 Ineficiente";
}
g.eficiencia = eficienciaGrupo;
g.avgDmi = g.nDmi ? g.sumDmi / g.nDmi : NaN;
g.avgFcr = g.nFcr ? g.sumFcr / g.nFcr : NaN;

let eficienciaGrupo = "—";
if (Number.isFinite(g.avgFcr)) {
  if (g.avgFcr < 6) eficienciaGrupo = "🟢 Excelente";
  else if (g.avgFcr < 7.5) eficienciaGrupo = "🟡 Normal";
  else eficienciaGrupo = "🔴 Ineficiente";
}
g.eficiencia = eficienciaGrupo;
      const totalStatus = g.ok + g.warn + g.bad;
      g.risk = totalStatus ? ((g.warn + g.bad) / totalStatus) : 0;

      g.sortKey = (g.bad*1000) + (g.warn*100) - (g.ok*10);
      return g;
    }).sort((a,b)=>b.sortKey-a.sortKey || b.risk-a.risk || a.name.localeCompare(b.name));

    st.animalsOut = animalsOut;
    st.groupsOut = groupsOut;
    st.groupAgg = groupAgg;

    st.meta.processed_ok = okRows;
    st.meta.processed_fail = badRows;
    st.meta.delimiter = delim;
    st.generated_at = new Date().toISOString();

    return st;
  }

  /* ===================== ALERTAS (para dashboard) ===================== */
  function buildAlerts(groupsOut){
    const rows = [];
    for(const g of groupsOut || []){
      const total = (g.ok||0) + (g.warn||0) + (g.bad||0);
      if(total === 0) continue;

      const pRed = (g.bad||0) / total;
      const pRisk = ((g.bad||0) + (g.warn||0)) / total;

      if(pRed >= 0.15){
        rows.push({level:"bad", text:`🔥 ${g.name} — ALERTA VERMELHO`, meta:`🔴 ${(pRed*100).toFixed(0)}% | risco ${(pRisk*100).toFixed(0)}% | hist ${total}`});
      }else if(pRisk >= 0.30){
        rows.push({level:"warn", text:`⚠️ ${g.name} — ALERTA AMARELO`, meta:`risco ${(pRisk*100).toFixed(0)}% | hist ${total}`});
      }
    }

    rows.sort((a,b)=>{
      const la = a.level==="bad" ? 2 : 1;
      const lb = b.level==="bad" ? 2 : 1;
      if(lb!==la) return lb-la;
      return a.text.localeCompare(b.text);
    });

    return rows;
  }

  /* ===================== FORECAST (para planeamento) ===================== */
  function computeForecast(state, targets){
    const st = state;
    const todayUTC = st?.todayUTC ? new Date(st.todayUTC) : new Date();

    const targetM = Number.isFinite(targets?.targetM) ? targets.targetM : 620;
    const targetF = Number.isFinite(targets?.targetF) ? targets.targetF : 520;
    const fallback = Number.isFinite(st?.config?.CONSERVATIVE_FALLBACK_GMD) ? st.config.CONSERVATIVE_FALLBACK_GMD : DEFAULTS.CONSERVATIVE_FALLBACK_GMD;

    const rows = (st?.groupsOut || []).map(g=>{
      const pM = g.avgEstM;
      const pF = g.avgEstF;

      const realM = g.avgGmdM;
      const realF = g.avgGmdF;

      const gmdUsedM = pickGmdUsed(realM, fallback);
      const gmdUsedF = pickGmdUsed(realF, fallback);

      const daysM = calcDaysToTarget(pM, targetM, gmdUsedM);
      const daysF = calcDaysToTarget(pF, targetF, gmdUsedF);

      const dateM = Number.isFinite(daysM) ? fmtDate(addDaysUTC(todayUTC, Math.ceil(daysM))) : "—";
      const dateF = Number.isFinite(daysF) ? fmtDate(addDaysUTC(todayUTC, Math.ceil(daysF))) : "—";

      const [stM, stMClass] = classifyReady(pM, targetM);
      const [stF, stFClass] = classifyReady(pF, targetF);

      const minDays = (() => {
        const a = Number.isFinite(daysM) ? daysM : Infinity;
        const b = Number.isFinite(daysF) ? daysF : Infinity;
        const m = Math.min(a,b);
        return (m === Infinity) ? NaN : m;
      })();

      let estadoTexto = "";
      let estadoClass = "muted";
      if((g.m||0)>0 && (g.f||0)>0){
        if(stMClass==="ok" && stFClass==="ok"){ estadoTexto = "M e F prontos ✅"; estadoClass="ok"; }
        else if(stMClass==="ok" || stFClass==="ok"){ estadoTexto = "Parcialmente pronto (misto)"; estadoClass="warn"; }
        else if(stMClass==="warn" || stFClass==="warn"){ estadoTexto = "Quase lá (misto)"; estadoClass="warn"; }
        else { estadoTexto = "Em engorda (misto)"; estadoClass="muted"; }
      }else if((g.m||0)>0){
        estadoTexto = stM; estadoClass = stMClass;
      }else if((g.f||0)>0){
        estadoTexto = stF; estadoClass = stFClass;
      }else{
        estadoTexto = "—"; estadoClass="muted";
      }

      return {
        name: g.name,
        m: g.m, f: g.f,
        pM, pF,
        realM, realF,
        gmdUsedM, gmdUsedF,
        daysM, daysF,
        dateM, dateF,
        minDays,
        estadoTexto,
        estadoClass
      };
    });

    rows.sort((a,b)=>{
      const aReady = (Number.isFinite(a.minDays) && a.minDays===0) ? 0 : 1;
      const bReady = (Number.isFinite(b.minDays) && b.minDays===0) ? 0 : 1;
      if(aReady !== bReady) return aReady - bReady;

      const am = Number.isFinite(a.minDays) ? a.minDays : Infinity;
      const bm = Number.isFinite(b.minDays) ? b.minDays : Infinity;
      if(am !== bm) return am - bm;

      return a.name.localeCompare(b.name);
    });

    return rows;
  }
/* ===================== SMART PLANNING (IA lógica) ===================== */

function buildSmartPlanning(state, targets){

  const forecast = computeForecast(state, targets);

  const suggestions = [];

  for(const g of forecast){

    if(!Number.isFinite(g.minDays)) continue;

    let prioridade = "baixa";
    let acao = "";
    let nota = "";

    if(g.minDays === 0){
      prioridade = "alta";
      acao = "Agendar saída imediata";
      nota = "Grupo já atingiu peso alvo.";
    }
    else if(g.minDays <= 14){
      prioridade = "alta";
      acao = "Preparar logística de venda";
      nota = `Previsão atingir peso alvo em ${Math.ceil(g.minDays)} dias.`;
    }
    else if(g.minDays <= 45){
      prioridade = "media";
      acao = "Monitorizar evolução";
      nota = "Grupo a caminho do peso alvo.";
    }
    else{
      prioridade = "baixa";
      acao = "Manter plano alimentar";
      nota = "Ainda longe do peso objetivo.";
    }

    suggestions.push({
      grupo: g.name,
      prioridade,
      acao,
      nota,
      diasPrevistos: g.minDays,
      estado: g.estadoTexto
    });
  }

  const order = { alta:0, media:1, baixa:2 };

  suggestions.sort((a,b)=>{
    return order[a.prioridade] - order[b.prioridade]
      || a.diasPrevistos - b.diasPrevistos;
  });

  return suggestions;
}
  /* ===================== API GLOBAL ===================== */
const Core = {
  // config (read)
  DEFAULTS,

  // state
  loadState,
  saveState,
  clearState,

  // meteo cache
  clearMeteoCache(){ localStorage.removeItem(METEO_CACHE_KEY); },

  // main import
  async importCSVText(csvText, options = {}){
    const state = await processCSVText(csvText, options);
    saveState(state);
    return state;
  },

  // helpers for UIs
  escapeHtml,
  safeFloat,
  safeInt,
  fmtDate,

  detectDelimiter,
  parseNumber,

  // dashboard helpers
  buildAlerts,

  // planning helpers
  computeForecast,
  buildSmartPlanning,

  // targets convenience
  getTargets(){
    const st = ensureState();
    const t = st.targets || { targetM:620, targetF:520 };
    return {
      targetM: Number.isFinite(t.targetM) ? t.targetM : 620,
      targetF: Number.isFinite(t.targetF) ? t.targetF : 520
    };
  },
  setTargets(targetM, targetF){
    const st = ensureState();
    st.targets = {
      targetM: Number.isFinite(targetM) ? targetM : 620,
      targetF: Number.isFinite(targetF) ? targetF : 520
    };
    saveState(st);
    return st.targets;
  }
};

window.Core = Core;
})();
