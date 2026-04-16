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
if(r.pesoAtualNum){

 const key = `${animalId}|${Number(r.pesoAtualNum)}|${normalizeDate(r.dataAtual)}`;

  console.log("🆕 KEY NOVA:", key);
  console.log("📦 EXISTE?", existentesPesagens.has(key));

  if(!existentesPesagens.has(key)){

    await spCreatePesagem({
      Title: animalId,
      DataPesagem: r.dataAtual,
      Peso: r.pesoAtualNum,
      Origem: "Atual"
    }, token);

    existentesPesagens.add(key);
  }
}

    // 👉 PESO ANTERIOR
if(r.pesoAnteriorNum){

  const key = `${animalId}|${Number(r.pesoAnteriorNum)}`;

  console.log("🆕 KEY NOVA (ANTERIOR):", key);
  console.log("📦 EXISTE? (ANTERIOR)", existentesPesagens.has(key));

  if(!existentesPesagens.has(key)){

    await spCreatePesagem({
      Title: animalId,
      DataPesagem: r.dataAnterior,
      Peso: r.pesoAnteriorNum,
      Origem: "Anterior"
    }, token);

    existentesPesagens.add(key);
  }
}

  } // fecha o for

  console.log("✅ Sync concluído (sem duplicados)");
} // 🔥 fecha a função syncToSharePoint
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
    const peso = item.fields?.Peso;

    if(animal && peso != null){

      const key = `${animal}|${Number(peso)}`;
      set.add(key);

    }

  });

  return set;
}
