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

    const animalId = r.Titulo;

    // 1. Verifica animal
    const existsAnimal = await spGetAnimal(animalId, token);

    if(!existsAnimal){
      await spCreateAnimal({
        Title: animalId,
        Sexo: "",
        GrupoAtual: ""
      }, token);
    }

    // 2. Verifica pesagem
    const existsPeso = await spGetPesagem(animalId, r.DataPesagem, token);

    if(!existsPeso){
      await spCreatePesagem({
        Title: animalId,
        DataPesagem: r.DataPesagem,
        Peso: r.Peso,
        Origem: r.Origem
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
async function spGetPesagem(animalId, data, token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items?$filter=fields/Title eq '${animalId}' and fields/DataPesagem eq '${data}'`;

  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const j = await r.json();

  if(!j.value){
    console.error("Erro Graph (Pesagem):", j);
    return false;
  }

  return j.value.length > 0;
}
async function spCreatePesagem(data, token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items`;

  await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ fields: data })
  });
}
function formatDateToISO(ptDate){
  if(!ptDate) return null;

  const [d,m,y] = ptDate.split("-");
  return `${y}-${m}-${d}`;
}
