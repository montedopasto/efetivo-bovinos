// =============================
// CONFIG SHAREPOINT
// =============================

const SITE_ID = "montedopastopt.sharepoint.com,37408cfe-6b54-4cad-a7b7-c735c2a1adec,8fb9c4d3-c9b7-4d95-b563-9a81f2dd4f76";
const LIST_ANIMAIS_ID = "b15b7096-1e78-47ee-ad00-45afa575736a";
const LIST_PESAGENS_ID = "b67fc146-880a-4eb1-b578-5d24de132853";
let isSyncRunning = false;
// =============================
// SYNC PRINCIPAL
// =============================

async function syncToSharePoint(rows){

  isSyncRunning = true;

  const token = await Auth.getAccessToken();

  const existentesAnimais = await spGetAllAnimais(token);
  const pesagensParaEnviar = [];

  for(const r of rows){

    const animalId = String(r.animal).trim();
    if(!animalId || animalId === "—") continue;

    // ANIMAIS
    if(!existentesAnimais.has(animalId)){

      spCreateAnimal({
  Title: animalId,
  Sexo: r.sexo || "",
  GrupoAtual: r.grupo || ""
}, token).catch(e => console.error("Erro animal:", e));

      existentesAnimais.add(animalId);
    }

    // PESO ATUAL
    if(r.pesoAtualNum != null && r.dataAtual){
      pesagensParaEnviar.push({
        Title: animalId,
        DataPesagem: r.dataAtual,
        Peso: normalizePeso(r.pesoAtualNum),
        Origem: "Atual"
      });
    }

    // PESO ANTERIOR
    if(r.pesoAnteriorNum != null && r.dataAnterior){
      pesagensParaEnviar.push({
        Title: animalId,
        DataPesagem: r.dataAnterior,
        Peso: normalizePeso(r.pesoAnteriorNum),
        Origem: "Anterior"
      });
    }
  }

  // 🔥 AQUI É QUE ACONTECE A MAGIA
  console.log("TOTAL PESAGENS:", pesagensParaEnviar.length);

  await spBatchCreatePesagens(pesagensParaEnviar, token);

    console.log("✅ Sync concluído (batch)");

  isSyncRunning = false;

  if(typeof carregarMeteo === "function"){
    carregarMeteo();
  }
}
async function spGetAnimal(animalId, token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ANIMAIS_ID}/items?$filter=fields/Title eq '${animalId}'`;

  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const j = await r.json();

  if(!j.value){
  console.error("Erro Graph (Animal):", j);
  return false;
}
return j.value.length > 0;
}
async function spCreateAnimal(data, token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ANIMAIS_ID}/items`;

  await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ fields: data })
  });
}
async function spCreatePesagem(data, token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items`;

const dataNorm = normalizeDate(data.DataPesagem);

if(!dataNorm){
  console.error("❌ Data inválida:", data.DataPesagem);
  return;
}

const dataFinal = `${dataNorm}T00:00:00Z`;

  const chave = `${data.Title}|${Math.round(Number(data.Peso))}`;

const body = {
  fields: {
    Title: String(data.Title).trim(),
    DataPesagem: dataFinal,
    Peso: Number(data.Peso),
    Origem: String(data.Origem),
    Chave: chave
  }
};

  console.log("🚀 A enviar pesagem:", body);

  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  const txt = await res.text();

  if(!res.ok){

  if(txt.toLowerCase().includes("duplicate") || txt.toLowerCase().includes("unique")){
    console.warn("⚠️ Duplicado ignorado:", data.Title, data.Peso);
    return;
  }

  console.error("❌ ERRO SHAREPOINT PESAGEM:", txt);
  throw new Error("Erro ao criar pesagem");
}

  console.log("✅ Pesagem criada:", txt);
}
function formatDateToISO(ptDate){

  if(!ptDate) return null;

  ptDate = String(ptDate).trim();

  // 🔥 CASO 1: já vem com T (ISO completo)
  if(ptDate.includes("T")){
    const d = new Date(ptDate);
    if(isNaN(d.getTime())){
      console.error("❌ Data inválida:", ptDate);
      return null;
    }
    return d.toISOString().split("T")[0];
  }

  // 🔥 CASO 2: ISO simples (YYYY-MM-DD)
  if(/^\d{4}-\d{2}-\d{2}$/.test(ptDate)){
    return ptDate;
  }

  // 🔥 CASO 3: formato PT (DD-MM-YYYY)
  if(/^\d{2}-\d{2}-\d{4}$/.test(ptDate)){
    const [d,m,y] = ptDate.split("-");
    return `${y}-${m}-${d}`;
  }

  console.error("❌ Formato inválido:", ptDate);
  return null;
}
async function spGetAllPesagens(token){

  let url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items?$expand=fields&$top=999`;

  const set = new Set();

  while(url){

    const r = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    const j = await r.json();

    (j.value || []).forEach(item => {

      const animal = item.fields?.Title?.trim();
      const peso = item.fields?.Peso;

      if(animal && peso != null){

        const pesoNorm = Math.round(Number(peso));
        const key = `${animal}|${pesoNorm}`;

        set.add(key);
      }

    });

    url = j['@odata.nextLink'] || null;
  }

  console.log("📊 TOTAL PESAGENS SHAREPOINT:", set.size);

  return set;
}
function normalizeDate(dateStr){

  if(!dateStr) return null;

  dateStr = String(dateStr).trim();

  // já vem ISO com hora
  if(dateStr.includes("T")){
    return dateStr.split("T")[0];
  }

  // formato YYYY-MM-DD
  if(/^\d{4}-\d{2}-\d{2}$/.test(dateStr)){
    return dateStr;
  }

  // formato PT DD-MM-YYYY
  if(/^\d{2}-\d{2}-\d{4}$/.test(dateStr)){
    const [d,m,y] = dateStr.split("-");
    return `${y}-${m}-${d}`;
  }

  console.error("❌ Data inválida:", dateStr);
  return null;
}
async function spGetAllAnimais(token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ANIMAIS_ID}/items?$expand=fields&$top=5000`;

  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const j = await r.json();

  const set = new Set();

  (j.value || []).forEach(i => {

    const title = i.fields?.Title?.trim();

    if(title){
      set.add(title);
    }

  });

  return set;
}
function normalizePeso(p){
  return Math.round(Number(p));
}
async function spBatchCreatePesagens(pesagens, token){

  const BATCH_SIZE = 20;

  for(let i = 0; i < pesagens.length; i += BATCH_SIZE){
    console.log(`🚀 Batch ${i / BATCH_SIZE + 1} de ${Math.ceil(pesagens.length / BATCH_SIZE)}`);

    const chunk = pesagens.slice(i, i + BATCH_SIZE);

    const requests = chunk.map((p, index) => {

      const dataNorm = normalizeDate(p.DataPesagem);
      if(!dataNorm) return null;

      const dataFinal = `${dataNorm}T00:00:00Z`;
      const chave = `${p.Title}|${Math.round(Number(p.Peso))}`;

      return {
        id: String(index),
        method: "POST",
        url: `/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items`,
        headers: {
          "Content-Type": "application/json"
        },
        body: {
          fields: {
            Title: String(p.Title).trim(),
            DataPesagem: dataFinal,
            Peso: Number(p.Peso),
            Origem: String(p.Origem),
            Chave: chave
          }
        }
      };
    }).filter(Boolean);

    if(requests.length === 0) continue;

    const res = await fetch("https://graph.microsoft.com/v1.0/$batch", {
  method: "POST",
  headers: {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json"
  },
  body: JSON.stringify({ requests })
});

if(!res.ok){
  console.error("❌ Erro no batch request:", res.status);
}

let json = {};
try {
  json = await res.json();
} catch(e) {
  console.warn("⚠️ Resposta vazia no batch");
}

// 🔥 validar respostas
(json.responses || []).forEach(r => {

  if(r.status >= 400){

    const msg = JSON.stringify(r.body || "").toLowerCase();

    if(msg.includes("duplicate") || msg.includes("unique")){
      console.warn("⚠️ Duplicado ignorado");
    } else {
      console.error("❌ Erro no batch:", r);
    }

  }

});

// 🔥 PAUSA ENTRE BATCHES (MUITO IMPORTANTE)
await new Promise(r => setTimeout(r, 50));

  }

  console.log("🚀 Batch concluído");
}
