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

    const animalId = r.animal_id;
    const data = r.data;
    const peso = parseFloat(r.peso);

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
