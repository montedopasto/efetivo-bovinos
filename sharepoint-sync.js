// =============================
// CONFIG SHAREPOINT
// =============================

const SITE_ID = "montedopastopt.sharepoint.com,30b32348-8df0-4dbe-9...5cbe7,3a90922f-7a65-44d9-ae1e-ef11c749a820";
const LIST_ANIMAIS_ID = "COLOCA_AQUI";
const LIST_PESAGENS_ID = "COLOCA_AQUI";

// =============================
// SYNC PRINCIPAL
// =============================

async function syncToSharePoint(rows){

  const token = await Auth.getAccessToken();

  for(const r of rows){

    const animalId = r.animal;
const pesagens = [
  {
    data: r.dataAnterior ? formatDateToISO(r.dataAnterior) : null,
    peso: r.pesoAnteriorNum
  },
  {
    data: formatDateToISO(r.dataAtual),
    peso: r.pesoAtualNum
  }
];

    // 1. Verifica animal
    const existsAnimal = await spGetAnimal(animalId, token);

    if(!existsAnimal){
      await spCreateAnimal({
        animal_id: animalId,
        nif: r.nif || "",
        raca: r.raca || "",
        data_entrada: r.data_entrada || null
      }, token);
    }

    // 2. Verifica pesagem
    for(const p of pesagens){

  if(!p.data || !Number.isFinite(p.peso)) continue;

  const existsPeso = await spGetPesagem(animalId, p.data, token);

  if(!existsPeso){
    await spCreatePesagem({
      animal_id: animalId,
      data: p.data,
      peso: p.peso
    }, token);
  }
}

  }

  console.log("✅ Sync concluído");
}
async function spGetAnimal(animalId, token){

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ANIMAIS_ID}/items?$filter=fields/animal_id eq '${animalId}'`;

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

  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_PESAGENS_ID}/items?$filter=fields/animal_id eq '${animalId}' and fields/data eq '${data}'`;

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
