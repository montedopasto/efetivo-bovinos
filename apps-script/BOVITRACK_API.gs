const SHEET_ANIMAIS = "animais";
const SHEET_ALERTAS = "alertas_operacionais";
const SHEET_ANIMAIS_BACKUP = "_animais_backup";
const SYNC_TOKEN_PROPERTY = "BOVITRACK_SYNC_TOKEN";
const MAX_IMPORT_ROWS = 20000;

function doGet(e){
  const params = e && e.parameter ? e.parameter : {};
  const type = params.type || "";
  const action = params.action || "";

  if(type === "summary") return json(getPublicSummary());
  if(type === "alertas") return getAlertas();
  if(action === "update_alert") return json(updateAlert(params));
  if(action === "create_alert") return json(createAlert(params));
  return getAnimais();
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    let result;

    if(body.action === "create_alert") result = createAlert(body);
    else if(body.action === "update_alert") result = updateAlert(body);
    else if(body.action === "replace_animals") result = replaceAnimals(body);
    else result = {ok:false,error:"Ação inválida"};

    return json(result);
  }catch(err){
    return json({ok:false,error:publicError(err)});
  }
}

function replaceAnimals(body){
  assertSyncToken(body.token);
  const csv = String(body.csv || "");
  if(!csv.trim()) throw new Error("O ficheiro está vazio");
  if(csv.length > 12000000) throw new Error("O ficheiro é demasiado grande");

  const firstLine = csv.split(/\r?\n/,1)[0] || "";
  const delimiter = (firstLine.match(/;/g)||[]).length >= (firstLine.match(/,/g)||[]).length ? ";" : ",";
  const parsed = Utilities.parseCsv(csv, delimiter);
  if(parsed.length < 2) throw new Error("O ficheiro não contém animais");
  if(parsed.length - 1 > MAX_IMPORT_ROWS) throw new Error("O ficheiro excede o limite de animais");

  const headers = parsed.shift().map(function(value,index){
    return cleanText(index === 0 ? String(value).replace(/^\uFEFF/,"") : value);
  });
  const required = [
    "animal_id","grupo","sexo","data_nasc","peso_anterior","data_peso_anterior",
    "peso_atual","data_peso_atual","raca","nif","data_entrada","Alocacao"
  ];
  required.forEach(function(header){
    if(headers.indexOf(header) < 0) throw new Error("Falta a coluna obrigatória: " + header);
  });

  const idIndex = headers.indexOf("animal_id");
  const sexIndex = headers.indexOf("sexo");
  const seen = {};
  const rows = parsed.filter(function(row){
    return row.some(function(value){return cleanText(value) !== "";});
  }).map(function(row,rowIndex){
    const normalized = headers.map(function(_,columnIndex){
      return safeCell(row[columnIndex]);
    });
    const animalId = cleanText(normalized[idIndex]);
    if(!animalId) throw new Error("Animal sem identificação na linha " + (rowIndex + 2));
    if(seen[animalId]) throw new Error("Animal repetido no ficheiro: " + animalId);
    seen[animalId] = true;
    const sex = cleanText(normalized[sexIndex]).toUpperCase();
    if(sex !== "M" && sex !== "F") throw new Error("Sexo inválido na linha " + (rowIndex + 2));
    normalized[sexIndex] = sex;
    return normalized;
  });
  if(!rows.length) throw new Error("O ficheiro não contém animais válidos");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const book = SpreadsheetApp.getActive();
    const target = book.getSheetByName(SHEET_ANIMAIS);
    if(!target) throw new Error("A folha central de animais não existe");
    let backup = book.getSheetByName(SHEET_ANIMAIS_BACKUP);
    if(!backup) backup = book.insertSheet(SHEET_ANIMAIS_BACKUP);

    const oldData = target.getDataRange().getValues();
    backup.clearContents();
    if(oldData.length && oldData[0].length){
      ensureSheetSize(backup,oldData.length,oldData[0].length);
      backup.getRange(1,1,oldData.length,oldData[0].length).setValues(oldData);
      backup.hideSheet();
    }

    try{
      target.clearContents();
      ensureSheetSize(target,rows.length+1,headers.length);
      target.getRange(1,1,1,headers.length).setValues([headers]);
      target.getRange(2,1,rows.length,headers.length).setValues(rows);
      target.setFrozenRows(1);
    }catch(writeError){
      target.clearContents();
      if(oldData.length && oldData[0].length){
        ensureSheetSize(target,oldData.length,oldData[0].length);
        target.getRange(1,1,oldData.length,oldData[0].length).setValues(oldData);
      }
      throw writeError;
    }
    SpreadsheetApp.flush();
  }finally{
    lock.releaseLock();
  }

  const summary = getPublicSummary();
  return {ok:true,totalAnimals:summary.totalAnimals};
}

function assertSyncToken(candidate){
  const expected = PropertiesService.getScriptProperties().getProperty(SYNC_TOKEN_PROPERTY);
  if(!expected) throw new Error("Sincronização ainda não configurada");
  if(cleanText(candidate) !== expected) throw new Error("Código de sincronização inválido");
}

function getPublicSummary(){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ANIMAIS);
  if(!sheet) throw new Error("Folha de animais indisponível");
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) throw new Error("Sem animais disponíveis");
  const headers = data.shift().map(cleanText);
  const index = {};
  headers.forEach(function(header,column){index[header]=column;});
  const aggregate = {M:emptySex(),F:emptySex()};
  const today = startUtcDay(new Date());

  data.forEach(function(row){
    const sex = cleanText(row[index.sexo]).toUpperCase();
    if(!aggregate[sex]) return;
    const bucket = aggregate[sex];
    bucket.count += 1;

    const currentWeight = validNumber(row[index.peso_atual]);
    if(currentWeight !== null && currentWeight > 0){
      bucket.weightTotal += currentWeight;
      bucket.weightCount += 1;
    }

    const birthDate = validDate(row[index.data_nasc]);
    if(birthDate && birthDate <= today){
      bucket.ageTotal += daysBetween(birthDate,today) / 30.4375;
      bucket.ageCount += 1;
    }

    const previousWeight = validNumber(row[index.peso_anterior]);
    const previousDate = validDate(row[index.data_peso_anterior]);
    const currentDate = validDate(row[index.data_peso_atual]);
    if(previousWeight !== null && currentWeight !== null && previousDate && currentDate){
      const days = daysBetween(previousDate,currentDate);
      const gmd = days > 0 ? (currentWeight-previousWeight)/days : null;
      if(gmd !== null && gmd >= 0 && gmd <= 3){
        bucket.gmdTotal += gmd;
        bucket.gmdCount += 1;
      }
    }
  });

  return {
    schemaVersion:2,
    generatedAt:new Date().toISOString(),
    totalAnimals:aggregate.M.count+aggregate.F.count,
    male:publicSex(aggregate.M),
    female:publicSex(aggregate.F)
  };
}

function getAnimais(){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ANIMAIS);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return json(data.map(function(row){
    const object = {};
    headers.forEach(function(header,index){object[header]=row[index];});
    return object;
  }));
}

function getAlertas(){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ALERTAS);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return json(data.map(function(row){
    const object = {};
    headers.forEach(function(header,index){object[header]=row[index];});
    return object;
  }));
}

function createAlert(body){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ALERTAS);
  sheet.appendRow([
    Utilities.getUuid(),body.grupo,"aberto",body.prioridade,new Date(),body.prazo,"","",new Date()
  ]);
  return {ok:true};
}

function updateAlert(body){
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ALERTAS);
  const data = sheet.getDataRange().getValues();
  for(let row=1;row<data.length;row++){
    if(data[row][0] == body.id){
      sheet.getRange(row+1,3).setValue(body.estado);
      sheet.getRange(row+1,7).setValue(body.responsavel);
      sheet.getRange(row+1,8).setValue(body.acao);
      sheet.getRange(row+1,9).setValue(new Date());
      break;
    }
  }
  return {ok:true};
}

function emptySex(){
  return {count:0,weightTotal:0,weightCount:0,gmdTotal:0,gmdCount:0,ageTotal:0,ageCount:0};
}

function publicSex(bucket){
  return {
    averageWeightKg:round(bucket.weightCount?bucket.weightTotal/bucket.weightCount:null,1),
    averageGmdKgDay:round(bucket.gmdCount?bucket.gmdTotal/bucket.gmdCount:null,2),
    averageAgeMonths:round(bucket.ageCount?bucket.ageTotal/bucket.ageCount:null,1)
  };
}

function validNumber(value){
  const number = Number.parseFloat(cleanText(value).replace(",","."));
  return Number.isFinite(number)?number:null;
}

function validDate(value){
  if(Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) return startUtcDay(value);
  const input = cleanText(value);
  let match = input.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(match) return new Date(Date.UTC(+match[1],+match[2]-1,+match[3]));
  match = input.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})$/);
  return match?new Date(Date.UTC(+match[3],+match[2]-1,+match[1])):null;
}

function startUtcDay(date){
  return new Date(Date.UTC(date.getFullYear(),date.getMonth(),date.getDate()));
}

function daysBetween(start,end){
  return Math.floor((end.getTime()-start.getTime())/86400000);
}

function round(value,digits){
  if(!Number.isFinite(value)) return null;
  const factor = Math.pow(10,digits);
  return Math.round(value*factor)/factor;
}

function cleanText(value){
  return String(value === null || value === undefined ? "" : value).trim();
}

function safeCell(value){
  const text = cleanText(value);
  return /^[=+@]/.test(text)?"'"+text:text;
}

function ensureSheetSize(sheet,rows,columns){
  if(sheet.getMaxRows()<rows) sheet.insertRowsAfter(sheet.getMaxRows(),rows-sheet.getMaxRows());
  if(sheet.getMaxColumns()<columns) sheet.insertColumnsAfter(sheet.getMaxColumns(),columns-sheet.getMaxColumns());
}

function publicError(error){
  const message = error && error.message ? error.message : String(error);
  return message.replace(/^Exception:\s*/,"");
}

function json(object){
  return ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON);
}
