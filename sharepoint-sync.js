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

  for(const r of rows){

  const animalId = r.animal;

  // 1. Criar animal se não existir
  const existsAnimal = await spGetAnimal(animalId, token);

  if(!existsAnimal){
    await spCreateAnimal({
      Title: animalId,
      Sexo: r.sexo || "",
      GrupoAtual: r.grupo || ""
    }, token);
  }

  // 2. PESO ATUAL
  if(r.dataAtual && r.pesoAtualNum){
    await spCreatePesagem({
      Title: animalId,
      DataPesagem: r.dataAtual,
      Peso: r.pesoAtualNum,
      Origem: "Atual"
    }, token);
  }

  // 3. PESO ANTERIOR
  if(r.dataAnterior && r.pesoAnteriorNum){
    await spCreatePesagem({
      Title: animalId,
      DataPesagem: r.dataAnterior,
      Peso: r.pesoAnteriorNum,
      Origem: "Anterior"
    }, token);
  }

}

  console.log("✅ Sync concluído");
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

  const body = {
    fields: {
      Title: String(data.Title),
      DataPesagem: formatDateToISO(data.DataPesagem) + "T00:00:00Z",
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

  const [d,m,y] = ptDate.split("-");
  return `${y}-${m}-${d}`;
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
      const dataNorm = new Date(data).toISOString();
      set.add(`${animal}|${dataNorm}`);
    }
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
