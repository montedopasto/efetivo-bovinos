// =============================
// CONFIG SHAREPOINT
// =============================

const SITE_ID = "COLOCA_AQUI";
const LIST_ANIMAIS_ID = "COLOCA_AQUI";
const LIST_PESAGENS_ID = "COLOCA_AQUI";

// =============================
// SYNC PRINCIPAL
// =============================

async function syncToSharePoint(rows){

  const token = await Auth.getAccessToken();

  for(const r of rows){

    const animalId = r.animal;
const data = r.dataAtual; // depois ajustamos formato se quiseres
const peso = parseFloat((r.pesoAtual || "").replace(" kg",""));

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
    const existsPeso = await spGetPesagem(animalId, data, token);

    if(!existsPeso){
      await spCreatePesagem({
        animal_id: animalId,
        data: data,
        peso: peso
      }, token);
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
const data = formatDateToISO(r.dataAtual);
