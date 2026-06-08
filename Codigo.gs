/**
 * ============================================================================
 * ARQUIVO: Codigo.gs (ou Codigo.js)
 * DESCRIÇÃO: Backend do aplicativo TarefasPRI (Gerenciador Premium)
 * VERSÃO: 1.2 - Melhorias de performance e sugestão de segurança (ActiveUser)
 * ============================================================================
 */

/**
 * Função padrão que serve a página HTML quando a URL do Web App é acessada.
 * Configura o título, o modo de frame e a responsividade.
 * * DICA DE SEGURANÇA FUTURA: Se você for usar este app dentro de uma conta Google Workspace corporativa, 
 * você pode restringir o acesso apenas para você alterando a implantação ou validando o email usando:
 * const usuarioLogado = Session.getActiveUser().getEmail();
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TarefasPRI - Produtividade Premium')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * SALVA dados na nuvem (Google Sheets) em lote
 * Recebe a string JSON do HTML e sobrescreve a planilha com os dados atualizados.
 */
function salvarBackupNoSheets(dadosTarefas) {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let aba = planilha.getSheetByName('TarefasPRI') || planilha.insertSheet('TarefasPRI');
    
    // 1. Limpa todos os dados existentes para receber o backup fresco
    aba.clear();
    
    // 2. Transforma toda a planilha em "Texto Puro" (@).
    // Isso impede que o Google Sheets engula os zeros à esquerda das datas (Ex: 08/06/2026 virar 8/6/2026)
    aba.getRange(1, 1, aba.getMaxRows(), aba.getMaxColumns()).setNumberFormat('@');
    
    const tarefas = JSON.parse(dadosTarefas);
    
    // 3. Montamos a primeira linha (Cabeçalho) em uma matriz de dados
    const matrizDados = [
      ['ID', 'Nome', 'Hora', 'Status', 'Prioridade', 'Categoria', 'Data Criação', 'Justificativa', 'Recorrente']
    ];
    
    // 4. Preenchemos a matriz com as tarefas
    tarefas.forEach(t => {
      matrizDados.push([
        t.id, 
        t.nome, 
        t.hora, 
        t.status, 
        t.prioridade, 
        t.categoria || "Outros", 
        t.data, 
        t.justificativa || "",
        t.recorrente ? "Sim" : "Não"
      ]);
    });
    
    // 5. Escreve todos os dados na planilha de uma só vez (muito mais rápido que appendRow)
    // Pega o intervalo exato: da linha 1, coluna 1, até o número total de linhas e colunas da matriz
    aba.getRange(1, 1, matrizDados.length, matrizDados[0].length).setValues(matrizDados);
    
    return "Backup TarefasPRI salvo com sucesso!";
  } catch (erro) {
    return "Erro no backup: " + erro.toString();
  }
}

/**
 * LÊ dados da nuvem e envia para o front-end
 * Retorna as tarefas armazenadas no Google Sheets no formato JSON.
 */
function carregarDadosDoSheets() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName('TarefasPRI');
    
    // Se a aba não existe, retorna um array vazio
    if (!aba) return JSON.stringify([]); 
    
    // getDisplayValues garante que pegaremos o texto visualizado, já protegido pelo formato "Texto Puro"
    const dados = aba.getDataRange().getDisplayValues();
    if (dados.length <= 1) return JSON.stringify([]); 
    
    const tarefasDaNuvem = [];
    
    // Pula a linha 0 (cabeçalho) e monta os objetos para o HTML
    for (let i = 1; i < dados.length; i++) {
      tarefasDaNuvem.push({
        id: dados[i][0],
        nome: dados[i][1],
        hora: dados[i][2],
        status: dados[i][3],
        prioridade: dados[i][4],
        categoria: dados[i][5] || "Outros",
        data: dados[i][6], // Agora a data chega intacta e com os zeros, ex: "08/06/2026"
        justificativa: dados[i][7] || "",
        recorrente: dados[i][8] === "Sim" // Transforma de volta para o booleano (true/false) esperado no HTML
      });
    }
    
    return JSON.stringify(tarefasDaNuvem);
  } catch (erro) {
    return JSON.stringify({ erro: erro.toString() });
  }
}
