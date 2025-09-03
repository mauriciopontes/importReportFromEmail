/**
 * @fileoverview Script para automatizar a importação e tratamento de relatórios .xlsx
 * enviados por e-mail para uma Planilha Google. A execução é manual via menu,
 * o valor de corte é dinâmico e a lógica de tratamento (corte e ordenação)
 * se adapta ao tipo de extração selecionado ("Caixa" ou "Sobra").
 *
 * @version 5.0.0
 * @license Apache-2.0
 */

//================================================================
// CONFIGURAÇÕES GLOBAIS
// As constantes são nomeadas em maiúsculas para indicar que são valores fixos e de fácil acesso.
//================================================================

/** @const {string} E-mail do remetente que envia o relatório. */
const REMETENTE_ESPERADO = "remetente@exemplo.com"; 

/** @const {string} Parte do assunto do e-mail para identificação. */
const ASSUNTO_DO_EMAIL = "Seu Relatório Diário"; 

/** @const {string} Nome da aba onde os dados tratados serão inseridos. */
const NOME_DA_ABA_PRINCIPAL = "exempl: planilha1"; 

/** @const {string} Nome da aba para onde as linhas com caixa negativo serão movidas. */
const NOME_DA_ABA_NEGATIVOS = "Exemplo: planilha2";

/** @const {Array<string>} Cabeçalho para a aba de caixas negativos. */
const CABECALHO_NEGATIVOS = ["Escolha", "os Cabeçalhos", "Exemplo: loja", "Caixa"];


//================================================================
// FUNÇÃO DE MENU (onOpen)
//================================================================

/**
 * Cria um menu personalizado na interface da Planilha Google sempre que o
 * arquivo é aberto. Este é um gatilho simples do Apps Script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Automação Completa')
      .addItem('Extração por Caixa', 'iniciarExtracaoPorCaixa')
      .addItem('Extração por Sobra', 'iniciarExtracaoPorSobra')
      .addToUi();
}

//================================================================
// FUNÇÕES DE INICIALIZAÇÃO (CHAMADAS PELO MENU)
//================================================================

/**
 * Inicia o fluxo de extração usando o critério "caixa".
 * Ponto de entrada para o botão de menu "Extração por Caixa".
 */
function iniciarExtracaoPorCaixa() {
  executarFluxoDeExtracao("caixa");
}

/**
 * Inicia o fluxo de extração usando o critério "sobra".
 * Ponto de entrada para o botão de menu "Extração por Sobra".
 */
function iniciarExtracaoPorSobra() {
  executarFluxoDeExtracao("sobra");
}

//================================================================
// ORQUESTRADOR PRINCIPAL DO FLUXO
//================================================================

/**
 * Orquestra todo o processo de extração: pede o valor de corte, busca e-mail,
 * converte anexo, processa os dados e escreve os resultados na planilha.
 * @param {string} criterioDeExtracao - O critério para as regras ("caixa" ou "sobra").
 */
function executarFluxoDeExtracao(criterioDeExtracao) {
  const ui = SpreadsheetApp.getUi();
  let idPlanilhaTemporaria = null;

  try {
    // Etapa 1: Obter o valor de corte do usuário através da caixa de diálogo.
    const valorDeCorte = obterValorDeCorteDoUsuario(criterioDeExtracao);
    if (valorDeCorte === null) return; // Interrompe se o usuário cancelar.

    // Etapa 2: Encontrar o anexo mais recente no Gmail.
    const anexo = encontrarAnexoMaisRecente();
    if (!anexo) throw new Error("Nenhum anexo .xlsx válido foi encontrado no e-mail mais recente.");

    // Etapa 3: Converter o anexo para dados de planilha.
    const resultadoConversao = converterAnexoParaDados(anexo);
    idPlanilhaTemporaria = resultadoConversao.idArquivoTemp;
    const dadosImportados = resultadoConversao.dados;
    
    if (!dadosImportados || dadosImportados.length <= 1) {
      throw new Error("O arquivo Excel importado está vazio ou contém apenas o cabeçalho.");
    }

    // Etapa 4: Processar os dados com base nas regras de negócio.
    const resultadoDoTratamento = processarDados(dadosImportados, criterioDeExtracao, valorDeCorte);
    
    // Etapa 5: Escrever os resultados nas abas de destino.
    escreverResultadosNaPlanilha(resultadoDoTratamento);
    
    Logger.log('Extração por "' + criterioDeExtracao + '" concluída com sucesso!');

  } catch (e) {
    Logger.log("Ocorreu um erro no processo: " + e.toString());
    ui.alert('Ocorreu um erro: ' + e.message);
  } finally {
    // Etapa 6: Limpeza - garantir que o arquivo temporário seja sempre excluído.
    if (idPlanilhaTemporaria) {
      DriveApp.getFileById(idPlanilhaTemporaria).setTrashed(true);
      Logger.log("Arquivo temporário " + idPlanilhaTemporaria + " movido para a lixeira.");
    }
  }
}

//================================================================
// FUNÇÕES AUXILIARES (RESPONSABILIDADES ÚNICAS)
//================================================================

/**
 * Exibe uma caixa de diálogo dinâmica para o usuário, captura e valida o valor de corte.
 * @param {string} criterio - O critério de extração ("caixa" ou "sobra") para personalizar o texto.
 * @returns {number|null} O valor de corte numérico ou nulo se a operação for cancelada.
 */
function obterValorDeCorteDoUsuario(criterio) {
  const ui = SpreadsheetApp.getUi();
  
  const titulo = 'Valor de Corte para Filtro';
  const texto = (criterio === 'sobra') 
    ? 'Filtrar SOBRAS a partir de (use "." para decimais):' 
    : 'Filtrar CAIXAS a partir de (use "." para decimais):';

  const respostaPrompt = ui.prompt(titulo, texto, ui.ButtonSet.OK_CANCEL);

  if (respostaPrompt.getSelectedButton() !== ui.Button.OK) {
    Logger.log('Operação cancelada pelo usuário.');
    return null; 
  }

  const valorDeCorteTexto = respostaPrompt.getResponseText();
  const valorDeCorteNumerico = parseFloat(valorDeCorteTexto);

  if (isNaN(valorDeCorteNumerico)) {
    throw new Error('Valor de corte inválido. Por favor, insira apenas números.');
  }
  
  return valorDeCorteNumerico;
}

/**
 * Procura no Gmail pela conversa mais recente que corresponde aos critérios
 * e retorna o anexo .xlsx mais recente dentro dessa conversa.
 * @returns {GoogleAppsScript.Gmail.GmailAttachment | null} O objeto do anexo ou nulo se não for encontrado.
 */
function encontrarAnexoMaisRecente() {
  const query = `from:${REMETENTE_ESPERADO} subject:"${ASSUNTO_DO_EMAIL}" has:attachment newer_than:1d`;
  const threads = GmailApp.search(query, 0, 1);
  if (threads.length === 0) return null;
  
  const messages = threads[0].getMessages();
  for (let i = messages.length - 1; i >= 0; i--) {
    const attachments = messages[i].getAttachments();
    for (let j = 0; j < attachments.length; j++) {
      if (attachments[j].getName().endsWith('.xlsx')) {
        return attachments[j];
      }
    }
  }
  return null;
}

/**
 * Converte um anexo .xlsx em um array 2D de dados usando o Google Drive.
 * @param {GoogleAppsScript.Gmail.GmailAttachment} anexoBlob - O anexo a ser convertido.
 * @returns {{dados: Object[][], idArquivoTemp: string}} Um objeto contendo os dados e o ID do arquivo temporário.
 */
function converterAnexoParaDados(anexoBlob) {
  const metaDadosDoArquivo = {
    title: 'temp_conversao_' + new Date().getTime(),
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  const arquivoPlanilha = Drive.Files.create(metaDadosDoArquivo, anexoBlob);
  const idArquivoTemp = arquivoPlanilha.id;
  const planilhaTemporaria = SpreadsheetApp.openById(idArquivoTemp);
  const dados = planilhaTemporaria.getSheets()[0].getDataRange().getValues();
  return { dados: dados, idArquivoTemp: idArquivoTemp };
}

/**
 * Limpa e escreve os dados processados nas abas de destino.
 * @param {{dadosFinais: Object[][], linhasNegativas: Object[][]}} resultado - O objeto com os dados a serem escritos.
 */
function escreverResultadosNaPlanilha(resultado) {
  const planilhaFinal = SpreadsheetApp.getActiveSpreadsheet();
  const abaPrincipal = planilhaFinal.getSheetByName(NOME_DA_ABA_PRINCIPAL);
  const abaNegativos = planilhaFinal.getSheetByName(NOME_DA_ABA_NEGATIVOS);

  abaPrincipal.clear(); 
  if (resultado.dadosFinais.length > 0) {
    abaPrincipal.getRange(1, 1, resultado.dadosFinais.length, resultado.dadosFinais[0].length).setValues(resultado.dadosFinais);
  }
  
  abaNegativos.clear();
  if (resultado.linhasNegativas.length > 0) {
    abaNegativos.getRange(1, 1, 1, CABECALHO_NEGATIVOS.length).setValues([CABECALHO_NEGATIVOS]);
    abaNegativos.getRange(2, 1, resultado.linhasNegativas.length, CABECALHO_NEGATIVOS.length).setValues(resultado.linhasNegativas);
  }
}

/**
 * Converte um valor de texto (ex: "R$ 1.234,56") em um número de ponto flutuante.
 * @param {string | number} valor - O valor da célula a ser convertido.
 * @returns {number} O valor convertido para número. Retorna 0 se a conversão falhar.
 */
function parsearValorNumerico(valor) {
  if (typeof valor === 'number') return valor;
  if (typeof valor !== 'string') return 0;
  const valorLimpo = valor.replace(/[^0-9.,-]/g, "").replace(/\./g, "").replace(",", ".");
  return parseFloat(valorLimpo) || 0;
}

/**
 * Aplica as regras de negócio aos dados importados e ordena o resultado.
 * @param {Object[][]} dados - O array 2D de dados vindo da planilha, com cabeçalho.
 * @param {string} criterio - O critério de extração ("caixa" ou "sobra").
 * @param {number} valorDeCorte - O valor mínimo para o valor ser mantido (Regra 3).
 * @returns {{dadosFinais: Object[][], linhasNegativas: Object[][]}} Objeto com os resultados.
 */
function processarDados(dados, criterio, valorDeCorte) {
  const cabecalho = dados.shift(); 
  const idx = {
    advisor: cabecalho.indexOf("Advisor"), sinacor: cabecalho.indexOf("Sinacor"), corretora: cabecalho.indexOf("Corretora"),
    caixa: cabecalho.indexOf("Caixa"), caixaDN: cabecalho.indexOf("Caixa D+N"),
    sobraCaixa: cabecalho.indexOf("Sobra de Caixa"), sobraCaixaDN: cabecalho.indexOf("Sobra de Caixa D+N")
  };

  // Validação das colunas necessárias para cada tipo de extração.
  if (idx.caixa === -1 || idx.caixaDN === -1) {
      throw new Error('As colunas "Caixa" e/ou "Caixa D+N" não foram encontradas no arquivo.');
  }
  if (criterio === 'sobra' && (idx.sobraCaixa === -1 || idx.sobraCaixaDN === -1)) {
      throw new Error('Para "Extração por Sobra", as colunas "Sobra de Caixa" e/ou "Sobra de Caixa D+N" são obrigatórias.');
  }

  const dadosProcessados = [];
  const linhasNegativas = [];

  dados.forEach(linha => {
    const caixaOriginal = parsearValorNumerico(linha[idx.caixa]);
    
    // REGRA 1: Se o caixa for negativo, a linha vai para a lista de negativos e o processo para aqui.
    if (caixaOriginal < 0) {
      linhasNegativas.push([ linha[idx.advisor], linha[idx.sinacor], linha[idx.corretora], caixaOriginal ]);
      return; 
    }
    
    let linhaFinal = [...linha];
    let valorParaCorte;

    // REGRA 2: A lógica de substituição muda com base no critério.
    if (criterio === "caixa") {
      const caixaDN = parsearValorNumerico(linha[idx.caixaDN]);
      
      // Condição: Se Caixa > Caixa D+N, o valor de Caixa é substituído pelo de Caixa D+N.
      if (caixaOriginal > caixaDN) {
        linhaFinal[idx.caixa] = caixaDN;
      }
      // O valor para o filtro de corte é o valor da coluna "Caixa" após a possível substituição.
      valorParaCorte = parsearValorNumerico(linhaFinal[idx.caixa]);

    } else { // criterio === "sobra"
      const sobraCaixa = parsearValorNumerico(linha[idx.sobraCaixa]);
      const sobraCaixaDN = parsearValorNumerico(linha[idx.sobraCaixaDN]);

      // Condição: Se Sobra de Caixa > Sobra de Caixa D+N, o valor de Sobra de Caixa é substituído.
      if (sobraCaixa > sobraCaixaDN) {
        linhaFinal[idx.sobraCaixa] = sobraCaixaDN;
      }
      // O valor para o filtro de corte é o valor da coluna "Sobra de Caixa" após a possível substituição.
      valorParaCorte = parsearValorNumerico(linhaFinal[idx.sobraCaixa]);
    }

    // REGRA 3: Manter a linha apenas se o valor de corte definido for atingido.
    if (valorParaCorte >= valorDeCorte) {
      dadosProcessados.push(linhaFinal);
    }
  });

  // ORDENAÇÃO: Define qual coluna será usada para a ordenação final com base no critério.
  const chaveDeOrdenacao = (criterio === 'sobra') ? idx.sobraCaixa : idx.caixa;
  dadosProcessados.sort((a, b) => {
    // Garante que os valores a serem comparados são numéricos.
    const valorA = parsearValorNumerico(a[chaveDeOrdenacao]);
    const valorB = parsearValorNumerico(b[chaveDeOrdenacao]);
    return valorB - valorA; // b - a para ordem decrescente.
  });

  // Adiciona o cabeçalho de volta ao topo dos dados processados antes de retornar.
  return {
    dadosFinais: [cabecalho, ...dadosProcessados],
    linhasNegativas: linhasNegativas
  };
}
