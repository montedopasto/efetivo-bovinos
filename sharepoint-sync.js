// =============================
// CONFIG SHAREPOINT
// =============================

const SITE_ID = "montedopastopt.sharepoint.com,37408cfe-6b54-4cad-a7b7-c735c2a1adec,8fb9c4d3-c9b7-4d95-b563-9a81f2dd4f76";
const LIST_ANIMAIS_ID = "b15b7096-1e78-47ee-ad00-45afa575736a";
const LIST_PESAGENS_ID = "b67fc146-880a-4eb1-b578-5d24de132853";

// =============================
// SYNC PRINCIPAL
// =============================

async function syncToSharePoint(rows){

  const token = await Auth.getAccessToken();

  // 🔥 carregar existentes (IMPORTANTÍSSIMO)
  const existentesAnimais = await spGetAllAnimais(token);
  const existentesPesagens = await spGetAllPesagens(token);
console.log("📚 EXISTENTES (10):", [...existentesPesagens].slice(0,10));
  for(const r of rows){

    const animalId = String(r.animal).trim();

    if(!animalId || animalId === "—") continue;

    // =============================
    // ANIMAIS (sem duplicados)
    // =============================
    if(!existentesAnimais.has(animalId)){

      await spCreateAnimal({
        Title: animalId,
        Sexo: r.sexo || "",
        GrupoAtual: r.grupo || ""
      }, token);

      existentesAnimais.add(animalId);
    }

    // =============================
    // PESAGENS
    // =============================

    // 👉 PESO ATUAL
    if(r.dataAtual && r.pesoAtualNum){

      const dataNorm = normalizeDate(r.dataAtual);

      if(dataNorm){

        const key = `${animalId}|${dataNorm}`;
console.log("🆕 KEY NOVA:", key);
console.log("📦 EXISTE?", existentesPesagens.has(key));
        if(!existentesPesagens.has(key)){

          await spCreatePesagem({
            Title: animalId,
            DataPesagem: dataNorm,
            Peso: r.pesoAtualNum,
            Origem: "Atual"
          }, token);

          existentesPesagens.add(key);
        }
      }
    }

    // 👉 PESO ANTERIOR
    if(r.dataAnterior && r.pesoAnteriorNum){

      const dataNorm = normalizeDate(r.dataAnterior);

      if(dataNorm){

        const key = `${animalId}|${dataNorm}`;
console.log("🆕 KEY NOVA (ANTERIOR):", key);
console.log("📦 EXISTE? (ANTERIOR)", existentesPesagens.has(key));
        if(!existentesPesagens.has(key)){

          await spCreatePesagem({
            Title: animalId,
            DataPesagem: dataNorm,
            Peso: r.pesoAnteriorNum,
            Origem: "Anterior"
          }, token);

          existentesPesagens.add(key);
        }
      }
    }

  }

  console.log("✅ Sync concluído (sem duplicados)");
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

  const dataISO = formatDateToISO(data.DataPesagem);

  if(!dataISO){
    console.error("❌ Data inválida:", data.DataPesagem);
    return;
  }

  // 🔥 CONVERSÃO CERTA PARA SHAREPOINT
  const dateObj = new Date(dataISO + "T00:00:00");

if(isNaN(dateObj.getTime())){
  console.error("❌ Data inválida após conversão:", dataISO);
  return;
}

const dataFinal = dateObj.toISOString();

  const body = {
    fields: {
      Title: String(data.Title).trim(),
      DataPesagem: dataFinal,
      Peso: Number(data.Peso),
      Origem: String(data.Origem)
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

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items?$expand=fields&$top=5000`;

  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const j = await r.json();

  const set = new Set();

  (j.value || []).forEach(item => {

    const animal = item.fields?.Title?.trim();
    const data = item.fields?.DataPesagem;

    if(animal && data){

      const dataNorm = normalizeDate(data);

      if(dataNorm){
        set.add(`${animal}|${dataNorm}`);
      }

    } // 🔥 ESTA CHAVE FALTAVA

  });

  return set;
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
function normalizeDate(dateStr){

  if(!dateStr) return null;

  dateStr = String(dateStr).trim();

  // ISO completo → devolver igual normalizado
  if(dateStr.includes("T")){
    const d = new Date(dateStr);
    if(isNaN(d)) return null;
    return d.toISOString();
  }

  // YYYY-MM-DD → converter para ISO completo
  if(/^\d{4}-\d{2}-\d{2}$/.test(dateStr)){
    return dateStr + "T00:00:00.000Z";
  }

  // DD-MM-YYYY → converter para ISO completo
  if(/^\d{2}-\d{2}-\d{4}$/.test(dateStr)){
    const [d,m,y] = dateStr.split("-");
    return `${y}-${m}-${d}T00:00:00.000Z`;
  }

  console.error("❌ Data inválida:", dateStr);
  return null;
}
