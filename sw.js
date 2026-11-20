function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Registros") || ss.getActiveSheet();
    
    // Cria cabeçalho automaticamente se a planilha estiver vazia
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["ID_Leitura", "Data_Hora_Bip", "Codigo_QR", "Posto_Maquina", "Operador", "Data_Sincronizacao"]);
    }

    const payload = JSON.parse(e.postData.contents);
    const leituras = payload.leituras || [];

    if (leituras.length === 0) {
      return respostaJSON({ status: "ok", gravados: 0 });
    }

    // Carrega IDs existentes para impedir duplicidade em quedas de rede
    const ultimaLinha = sheet.getLastRow();
    const idsExistentes = new Set();
    if (ultimaLinha > 1) {
      const colA = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      for (let i = 0; i < colA.length; i++) {
        idsExistentes.add(String(colA[i][0]));
      }
    }

    const agoraSync = new Date().toLocaleString("pt-BR");
    const novasLinhas = [];

    leituras.forEach(item => {
      if (!idsExistentes.has(String(item.id))) {
        novasLinhas.push([
          item.id,
          item.dataHora,
          item.codigoQR,
          item.posto,
          item.operador,
          agoraSync
        ]);
      }
    });

    if (novasLinhas.length > 0) {
      sheet.getRange(ultimaLinha + 1, 1, novasLinhas.length, 6).setValues(novasLinhas);
    }

    return respostaJSON({ status: "ok", gravados: novasLinhas.length });
  } catch (erro) {
    return respostaJSON({ status: "erro", mensagem: erro.toString() });
  } finally {
    lock.releaseLock();
  }
}

function respostaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}