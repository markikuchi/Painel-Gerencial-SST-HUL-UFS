/**
 * API Central — Painel Gerencial SOST/SEG
 *
 * COMO INSTALAR:
 * 1. Abra a planilha Google Sheets do Painel
 * 2. Extensões > Apps Script
 * 3. Apague o conteúdo do Code.gs e cole este código
 * 4. Clique em "Implantar" > "Nova implantação"
 * 5. Tipo: "App da Web"
 * 6. Executar como: "Eu" (sua conta)
 * 7. Quem tem acesso: "Qualquer pessoa"
 * 8. Clique em "Implantar"
 * 9. Copie a URL gerada e cole no exportador HTML (campo SCRIPT_URL)
 *
 * FORMATO DO POST:
 * {
 *   "modulo": "adicionais",          // identifica quem está enviando
 *   "sheets": [
 *     { "name": "Resumo",    "rows": [["col1","col2",...], ["val1","val2",...]] },
 *     { "name": "Empregados", "rows": [["col1","col2",...], [...], ...] }
 *   ]
 * }
 *
 * Cada módulo futuro envia seus dados no mesmo formato,
 * informando as abas que são dele.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return _jsonResponse({ status: 'error', message: 'Servidor ocupado. Tente novamente.' });
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultado = [];

    if (!data.sheets || !Array.isArray(data.sheets)) {
      return _jsonResponse({ status: 'error', message: 'Formato inválido: "sheets" deve ser um array.' });
    }

    data.sheets.forEach(function(sheetData) {
      var sheet = ss.getSheetByName(sheetData.name);
      if (!sheet) {
        sheet = ss.insertSheet(sheetData.name);
      }

      // Limpa a aba inteira antes de escrever
      sheet.clear();

      if (sheetData.rows && sheetData.rows.length > 0) {
        // Garante que todas as linhas tenham o mesmo número de colunas
        var maxCols = 0;
        sheetData.rows.forEach(function(row) {
          if (row.length > maxCols) maxCols = row.length;
        });
        var normalizado = sheetData.rows.map(function(row) {
          while (row.length < maxCols) row.push('');
          return row;
        });

        sheet.getRange(1, 1, normalizado.length, maxCols).setValues(normalizado);
      }

      resultado.push({ aba: sheetData.name, linhas: (sheetData.rows || []).length });
    });

    // Log de atualização
    var logSheet = ss.getSheetByName('_log');
    if (!logSheet) {
      logSheet = ss.insertSheet('_log');
      logSheet.getRange(1, 1, 1, 4).setValues([['timestamp', 'modulo', 'abas', 'total_linhas']]);
    }
    var totalLinhas = resultado.reduce(function(s, r) { return s + r.linhas; }, 0);
    var abas = resultado.map(function(r) { return r.aba; }).join(', ');
    logSheet.appendRow([new Date(), data.modulo || 'desconhecido', abas, totalLinhas]);

    return _jsonResponse({
      status: 'ok',
      modulo: data.modulo || 'desconhecido',
      resultado: resultado,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    return _jsonResponse({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return _jsonResponse({
    status: 'ok',
    message: 'API Painel Gerencial SOST/SEG ativa.',
    versao: '1.0',
    timestamp: new Date().toISOString()
  });
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
