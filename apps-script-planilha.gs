// Super Liga Águias — recebe o placar do sistema e escreve nas abas Placar, Vendas e Medalhas.
// Como usar: na planilha > Extensões > Apps Script > cole este código > Implantar > Nova implantação
// > Tipo: App da Web > Executar como: Eu > Quem tem acesso: Qualquer pessoa > Implantar. Copie a URL /exec
// e cole em ADM > Config > "URL da planilha".

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  escreve_(ss, 'Placar', d.placar);
  escreve_(ss, 'Vendas', d.vendas);
  escreve_(ss, 'Medalhas', d.medalhas);
  return ContentService.createTextOutput('ok');
}

function doGet() {
  return ContentService.createTextOutput('Liga Aguias: webhook ativo');
}

function escreve_(ss, nome, linhas) {
  var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
  sh.clear();
  if (!linhas || !linhas.length) return;
  var w = linhas[0].length;
  sh.getRange(1, 1, linhas.length, w).setValues(linhas);
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#17130d').setFontColor('#ffc800');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, w);
}
