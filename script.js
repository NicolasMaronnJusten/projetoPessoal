const NOME_ABA_PRINCIPAL = "Jornadas";
const NOME_ARQUIVO_FINAL = "Relatorio_Jornadas_Sem_Segundos.xlsx";
const LIMITE_PREVIEW = 20;
const TERMOS_COLUNAS_TEMPO = [
  "total em servico",
  "total em direcao",
  "tempo em descanso",
  "tempo em refeicao",
  "he diurna",
  "he diurna 100",
  "hora noturna",
  "he noturna",
  "he noturna 100",
  "total em espera",
  "interjornada",
  "jornada",
  "horas",
  "hora",
  "tempo",
  "servico",
  "direcao",
  "descanso",
  "refeicao",
];

const estado = {
  arquivo: null,
  tipoLeitura: null,
  workbook: null,
  abaProcessada: null,
  nomeAbaProcessada: "",
  alertas: [],
  celulasCorrigidas: 0,
  linhas: 0,
  colunas: 0,
  preview: [],
  processado: false,
};

const elementos = {};

document.addEventListener("DOMContentLoaded", inicializarEventos);

function inicializarEventos() {
  elementos.fileInput = document.getElementById("fileInput");
  elementos.dropArea = document.getElementById("dropArea");
  elementos.fileName = document.getElementById("fileName");
  elementos.processButton = document.getElementById("processButton");
  elementos.downloadButton = document.getElementById("downloadButton");
  elementos.statusMessage = document.getElementById("statusMessage");
  elementos.alertsBox = document.getElementById("alertsBox");
  elementos.previewContainer = document.getElementById("previewContainer");
  elementos.previewHint = document.getElementById("previewHint");
  elementos.statSheet = document.getElementById("statSheet");
  elementos.statRows = document.getElementById("statRows");
  elementos.statColumns = document.getElementById("statColumns");
  elementos.statChanges = document.getElementById("statChanges");

  elementos.fileInput.addEventListener("change", (evento) => {
    const arquivo = evento.target.files[0];
    selecionarArquivo(arquivo);
  });

  elementos.dropArea.addEventListener("dragover", (evento) => {
    evento.preventDefault();
    elementos.dropArea.classList.add("drag-over");
  });

  elementos.dropArea.addEventListener("dragleave", () => {
    elementos.dropArea.classList.remove("drag-over");
  });

  elementos.dropArea.addEventListener("drop", (evento) => {
    evento.preventDefault();
    elementos.dropArea.classList.remove("drag-over");
    const arquivo = evento.dataTransfer.files[0];
    selecionarArquivo(arquivo);
  });

  elementos.dropArea.addEventListener("keydown", (evento) => {
    if (evento.key === "Enter" || evento.key === " ") {
      evento.preventDefault();
      elementos.fileInput.click();
    }
  });

  elementos.processButton.addEventListener("click", processarPlanilha);
  elementos.downloadButton.addEventListener("click", baixarPlanilhaFinal);
}

function selecionarArquivo(arquivo) {
  limparResultado();

  if (!arquivo) {
    estado.arquivo = null;
    elementos.fileName.textContent = "Nenhum arquivo selecionado";
    elementos.processButton.disabled = true;
    return;
  }

  if (!arquivoEhExcel(arquivo)) {
    estado.arquivo = null;
    elementos.fileName.textContent = arquivo.name;
    elementos.processButton.disabled = true;
    mostrarMensagem(
      "error",
      "Arquivo inválido. Selecione uma planilha .xlsx ou .xls.",
    );
    mostrarAlertas([
      {
        tipo: "error",
        texto:
          "O arquivo selecionado não parece ser uma planilha Excel válida.",
      },
    ]);
    return;
  }

  estado.arquivo = arquivo;
  elementos.fileName.textContent = arquivo.name;
  elementos.processButton.disabled = false;
  mostrarMensagem("info", "Arquivo pronto para processamento.");
}

async function processarPlanilha() {
  if (!estado.arquivo) {
    mostrarMensagem("error", "Selecione uma planilha antes de processar.");
    return;
  }

  bloquearInterface(true);
  limparResultado(false);
  mostrarMensagem("info", "Lendo e processando a planilha...");

  try {
    const leitura = await lerArquivoExcel(estado.arquivo);
    estado.tipoLeitura = leitura.tipo;
    estado.workbook = leitura.workbook;

    const aba = encontrarAbaJornadas(leitura.workbook, leitura.tipo);
    estado.abaProcessada = aba.planilha;
    estado.nomeAbaProcessada = aba.nome;

    const resultado = processarAbaJornadas(aba.planilha, leitura.tipo);
    estado.celulasCorrigidas = resultado.celulasCorrigidas;
    estado.linhas = resultado.linhas;
    estado.colunas = resultado.colunas;
    estado.preview = gerarPreview(
      aba.planilha,
      leitura.tipo,
      resultado.linhas,
      resultado.colunas,
    );
    estado.processado = true;

    if (resultado.celulasCorrigidas === 0) {
      registrarAlerta(
        "warning",
        "Nenhum horário ou duração com segundos foi encontrado para corrigir.",
      );
    }

    if (resultado.invalidos.length > 0) {
      const exemplos = resultado.invalidos.slice(0, 8).join(", ");
      registrarAlerta(
        "warning",
        `Algumas células parecem horários, mas estão em formato inválido: ${exemplos}. Elas foram mantidas sem alteração.`,
      );
    }

    atualizarResumo();
    renderizarPreview();
    mostrarAlertas();
    elementos.downloadButton.disabled = false;
    mostrarMensagem(
      "success",
      "Planilha processada. Você já pode baixar o arquivo final.",
    );
  } catch (erro) {
    console.error(erro);
    estado.processado = false;
    elementos.downloadButton.disabled = true;
    registrarAlerta(
      "error",
      "O arquivo não pôde ser lido ou processado. Verifique se ele não está corrompido ou protegido por senha.",
    );
    mostrarAlertas();
    mostrarMensagem("error", "Não foi possível processar a planilha.");
  } finally {
    bloquearInterface(false);
  }
}

async function lerArquivoExcel(arquivo) {
  const extensao = obterExtensao(arquivo.name);
  const arrayBuffer = await arquivo.arrayBuffer();

  // ExcelJS preserva melhor a estrutura de arquivos .xlsx: abas, larguras,
  // mesclagens e estilos básicos continuam no mesmo workbook original.
  if (extensao === "xlsx") {
    if (!window.ExcelJS) {
      throw new Error("ExcelJS não foi carregado.");
    }

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      return { tipo: "exceljs", workbook };
    } catch (erro) {
      if (!window.XLSX) {
        throw erro;
      }

      registrarAlerta(
        "warning",
        "A leitura principal do .xlsx falhou. O sistema usou uma leitura alternativa, que pode preservar menos estilos, mas mantém os dados e as abas.",
      );
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
        cellNF: true,
        cellStyles: true,
        sheetStubs: true,
      });
      return { tipo: "sheetjs", workbook };
    }
  }

  // Arquivos .xls antigos são lidos com SheetJS. A saída continua sendo .xlsx.
  if (extensao === "xls") {
    if (!window.XLSX) {
      throw new Error("SheetJS não foi carregado.");
    }

    registrarAlerta(
      "info",
      "Arquivo .xls detectado. Os dados e as abas serão mantidos, mas alguns estilos do formato antigo podem não ser preservados integralmente.",
    );
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
      cellNF: true,
      cellStyles: true,
      sheetStubs: true,
    });
    return { tipo: "sheetjs", workbook };
  }

  throw new Error("Formato de arquivo não suportado.");
}

function encontrarAbaJornadas(workbook, tipoLeitura) {
  // Primeiro tenta localizar exatamente "Jornadas". Depois aceita variações de
  // maiúsculas/minúsculas e espaços para evitar falhas por digitação simples.
  if (tipoLeitura === "exceljs") {
    const exata = workbook.getWorksheet(NOME_ABA_PRINCIPAL);
    if (exata) {
      return { planilha: exata, nome: exata.name, encontrou: true };
    }

    const normalizada = normalizarTexto(NOME_ABA_PRINCIPAL);
    const parecida = workbook.worksheets.find(
      (sheet) => normalizarTexto(sheet.name) === normalizada,
    );
    if (parecida) {
      return { planilha: parecida, nome: parecida.name, encontrou: true };
    }

    const primeira = workbook.worksheets[0];
    if (!primeira) {
      throw new Error("Nenhuma aba encontrada.");
    }

    registrarAlerta(
      "warning",
      `A aba "${NOME_ABA_PRINCIPAL}" não foi encontrada. A primeira aba, "${primeira.name}", foi usada no processamento.`,
    );
    return { planilha: primeira, nome: primeira.name, encontrou: false };
  }

  const nomes = workbook.SheetNames || [];
  const nomeExato = nomes.find((nome) => nome === NOME_ABA_PRINCIPAL);
  if (nomeExato) {
    return {
      planilha: workbook.Sheets[nomeExato],
      nome: nomeExato,
      encontrou: true,
    };
  }

  const normalizada = normalizarTexto(NOME_ABA_PRINCIPAL);
  const nomeParecido = nomes.find(
    (nome) => normalizarTexto(nome) === normalizada,
  );
  if (nomeParecido) {
    return {
      planilha: workbook.Sheets[nomeParecido],
      nome: nomeParecido,
      encontrou: true,
    };
  }

  const primeiroNome = nomes[0];
  if (!primeiroNome) {
    throw new Error("Nenhuma aba encontrada.");
  }

  registrarAlerta(
    "warning",
    `A aba "${NOME_ABA_PRINCIPAL}" não foi encontrada. A primeira aba, "${primeiroNome}", foi usada no processamento.`,
  );
  return {
    planilha: workbook.Sheets[primeiroNome],
    nome: primeiroNome,
    encontrou: false,
  };
}

function processarAbaJornadas(planilha, tipoLeitura) {
  if (tipoLeitura === "exceljs") {
    return processarAbaExcelJS(planilha);
  }

  return processarAbaSheetJS(planilha);
}

function processarAbaExcelJS(planilha) {
  const dimensoes = obterDimensoesExcelJS(planilha);
  const colunasProvaveis = mapearColunasProvaveisExcelJS(
    planilha,
    dimensoes.colunas,
  );
  const invalidos = [];
  let celulasCorrigidas = 0;

  // Percorre a aba inteira sem remover linhas, colunas ou células. Quando uma
  // célula é duração/hora com segundos, somente o valor dela é ajustado.
  for (let linha = 1; linha <= dimensoes.linhas; linha += 1) {
    const row = planilha.getRow(linha);

    for (let coluna = 1; coluna <= dimensoes.colunas; coluna += 1) {
      const cell = row.getCell(coluna);
      const valor = extrairValorSimples(cell.value);

      if (
        cell.formula ||
        valor instanceof Date ||
        valor === null ||
        valor === undefined ||
        valor === ""
      ) {
        continue;
      }

      if (pareceDuracaoInvalida(valor)) {
        adicionarInvalido(
          invalidos,
          `${enderecoExcel(linha, coluna)} (${String(valor)})`,
        );
        continue;
      }

      const contexto = {
        converterNumero: deveConverterNumeroExcel(
          valor,
          cell.numFmt,
          colunasProvaveis[coluna],
          cell.text,
        ),
        manterDuasCasasHora: deveManterDuasCasasHora(cell.numFmt, cell.text),
        textoFormatado: cell.text,
        numFmt: cell.numFmt,
      };

      const valorFormatado = formatarDuracaoSemSegundos(valor, contexto);
      if (valorFormatado !== valor) {
        cell.value = valorFormatado;
        celulasCorrigidas += 1;
      }
    }
  }

  return {
    celulasCorrigidas,
    linhas: dimensoes.linhas,
    colunas: dimensoes.colunas,
    invalidos,
  };
}

function processarAbaSheetJS(planilha) {
  const range = obterRangeSheetJS(planilha);
  const colunasProvaveis = mapearColunasProvaveisSheetJS(planilha, range);
  const invalidos = [];
  let celulasCorrigidas = 0;

  if (!range) {
    return { celulasCorrigidas: 0, linhas: 0, colunas: 0, invalidos };
  }

  // SheetJS mantém a matriz da aba. Alteramos apenas as células que são
  // horários/durações; as demais abas ficam intocadas no workbook.
  for (let linha = range.s.r; linha <= range.e.r; linha += 1) {
    for (let coluna = range.s.c; coluna <= range.e.c; coluna += 1) {
      const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
      const cell = planilha[endereco];

      if (!cell || cell.f) {
        continue;
      }

      const valor = cell.v;
      if (
        valor instanceof Date ||
        valor === null ||
        valor === undefined ||
        valor === ""
      ) {
        continue;
      }

      const valorParaValidar = typeof valor === "string" ? valor : cell.w;
      if (pareceDuracaoInvalida(valorParaValidar)) {
        adicionarInvalido(
          invalidos,
          `${endereco} (${String(valorParaValidar)})`,
        );
        continue;
      }

      const colunaAbsoluta = coluna + 1;
      const contexto = {
        converterNumero: deveConverterNumeroExcel(
          valor,
          cell.z,
          colunasProvaveis[colunaAbsoluta],
          cell.w,
        ),
        manterDuasCasasHora: deveManterDuasCasasHora(cell.z, cell.w),
        textoFormatado: cell.w,
        numFmt: cell.z,
      };

      const valorFormatado = formatarDuracaoSemSegundos(valor, contexto);
      if (valorFormatado !== valor) {
        cell.v = valorFormatado;
        cell.t = "s";
        delete cell.w;
        delete cell.z;
        celulasCorrigidas += 1;
      }
    }
  }

  return {
    celulasCorrigidas,
    linhas: range.e.r - range.s.r + 1,
    colunas: range.e.c - range.s.c + 1,
    invalidos,
  };
}

function formatarDuracaoSemSegundos(valor, contexto = {}) {
  if (typeof valor === "string") {
    const partes = obterPartesDuracao(valor);
    if (!partes) {
      return valor;
    }

    if (partes.segundos === null) {
      return valor;
    }

    return `${partes.prefixo}${partes.horas}:${partes.minutos}${partes.sufixo}`;
  }

  // Números do Excel são dias: 0.5 equivale a 12 horas. Só convertemos quando
  // a célula ou a coluna indica que aquele número representa tempo/duração.
  if (typeof valor === "number" && contexto.converterNumero) {
    if (typeof contexto.textoFormatado === "string") {
      const textoFormatado = formatarDuracaoSemSegundos(
        contexto.textoFormatado,
      );
      if (
        textoFormatado !== contexto.textoFormatado ||
        ehDuracaoSemSegundos(textoFormatado)
      ) {
        return textoFormatado;
      }
    }

    return converterNumeroExcelParaDuracao(valor, contexto);
  }

  return valor;
}

function ehDuracaoComSegundos(valor) {
  const partes = obterPartesDuracao(valor);
  return Boolean(partes && partes.segundos !== null);
}

function ehDuracaoSemSegundos(valor) {
  const partes = obterPartesDuracao(valor);
  return Boolean(partes && partes.segundos === null);
}

function obterPartesDuracao(valor) {
  if (typeof valor !== "string") {
    return null;
  }

  const match = valor.match(/^(\s*)(\d{1,4}):(\d{2})(?::(\d{2}))?(\s*)$/);
  if (!match) {
    return null;
  }

  const minutos = Number(match[3]);
  const segundos = match[4] === undefined ? null : Number(match[4]);

  if (minutos > 59 || (segundos !== null && segundos > 59)) {
    return null;
  }

  return {
    prefixo: match[1],
    horas: match[2],
    minutos: match[3],
    segundos,
    sufixo: match[5],
  };
}

function converterNumeroExcelParaDuracao(valor, contexto = {}) {
  if (typeof valor !== "number" || !Number.isFinite(valor) || valor < 0) {
    return valor;
  }

  const totalSegundos = Math.floor(valor * 24 * 60 * 60 + 0.000001);
  const totalMinutos = Math.floor(totalSegundos / 60);
  const horas = Math.floor(totalMinutos / 60);
  const minutos = totalMinutos % 60;
  const horasTexto =
    contexto.manterDuasCasasHora || horas < 10
      ? String(horas).padStart(2, "0")
      : String(horas);

  return `${horasTexto}:${String(minutos).padStart(2, "0")}`;
}

function gerarPreview(planilha, tipoLeitura, totalLinhas, totalColunas) {
  if (tipoLeitura === "exceljs") {
    const linhas = [];
    const limiteLinhas = Math.min(totalLinhas, LIMITE_PREVIEW);

    for (let linha = 1; linha <= limiteLinhas; linha += 1) {
      const row = [];
      for (let coluna = 1; coluna <= totalColunas; coluna += 1) {
        const cell = planilha.getCell(linha, coluna);
        row.push(valorParaPreview(cell.value, cell.text));
      }
      linhas.push(row);
    }

    return linhas;
  }

  const range = obterRangeSheetJS(planilha);
  if (!range) {
    return [];
  }

  const linhas = [];
  const ultimaLinha = Math.min(range.e.r, range.s.r + LIMITE_PREVIEW - 1);

  for (let linha = range.s.r; linha <= ultimaLinha; linha += 1) {
    const row = [];
    for (let coluna = range.s.c; coluna <= range.e.c; coluna += 1) {
      const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
      const cell = planilha[endereco];
      row.push(cell ? valorParaPreview(cell.v, cell.w) : "");
    }
    linhas.push(row);
  }

  return linhas;
}

async function baixarPlanilhaFinal() {
  if (!estado.processado || !estado.workbook) {
    mostrarMensagem("error", "Processe uma planilha antes de baixar.");
    return;
  }

  try {
    elementos.downloadButton.disabled = true;

    if (estado.tipoLeitura === "exceljs") {
      const buffer = await estado.workbook.xlsx.writeBuffer();
      salvarArquivo(buffer, NOME_ARQUIVO_FINAL);
      elementos.downloadButton.disabled = false;
      return;
    }

    const buffer = XLSX.write(estado.workbook, {
      bookType: "xlsx",
      type: "array",
      cellStyles: true,
    });
    salvarArquivo(buffer, NOME_ARQUIVO_FINAL);
    elementos.downloadButton.disabled = false;
  } catch (erro) {
    console.error(erro);
    elementos.downloadButton.disabled = false;
    mostrarMensagem("error", "Não foi possível gerar a planilha final.");
  }
}

function salvarArquivo(buffer, nomeArquivo) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = nomeArquivo;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  mostrarMensagem("success", "Download iniciado.");
}

function deveConverterNumeroExcel(
  valor,
  numFmt,
  colunaProvavelTempo,
  textoFormatado,
) {
  if (typeof valor !== "number" || !Number.isFinite(valor) || valor < 0) {
    return false;
  }

  if (formatoPareceHorario(numFmt)) {
    return true;
  }

  if (
    typeof textoFormatado === "string" &&
    (ehDuracaoComSegundos(textoFormatado) ||
      ehDuracaoSemSegundos(textoFormatado))
  ) {
    return true;
  }

  return Boolean(colunaProvavelTempo && valor <= 31);
}

function formatoPareceHorario(numFmt = "") {
  const formato = limparFormatoExcel(numFmt);
  if (!formato || !formato.includes(":") || formatoPareceData(formato)) {
    return false;
  }

  return /(\[h+\]|h{1,2}|s{1,2})/.test(formato);
}

function formatoPareceData(formato = "") {
  const texto = limparFormatoExcel(formato);
  return (
    /\b(d{1,4}|y{2,4}|a{2,4})\b/.test(texto) ||
    /(^|[^a-z])d{1,2}[\/.-]/.test(texto)
  );
}

function limparFormatoExcel(numFmt = "") {
  return String(numFmt)
    .toLowerCase()
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[(?!h+\])([^\]]+)\]/g, "")
    .trim();
}

function deveManterDuasCasasHora(numFmt = "", textoFormatado = "") {
  const formato = limparFormatoExcel(numFmt);
  if (typeof textoFormatado === "string" && /^\s*\d{2}:/.test(textoFormatado)) {
    return true;
  }

  return /\bhh\b/.test(formato);
}

function pareceDuracaoInvalida(valor) {
  if (typeof valor !== "string") {
    return false;
  }

  const texto = valor.trim();
  if (
    !texto.includes(":") ||
    ehDuracaoComSegundos(texto) ||
    ehDuracaoSemSegundos(texto)
  ) {
    return false;
  }

  const partes = texto.split(":");
  const pareceTempo =
    partes.length >= 2 &&
    partes.length <= 3 &&
    partes.every((parte) => /^\d+$/.test(parte));

  return pareceTempo;
}

function mapearColunasProvaveisExcelJS(planilha, totalColunas) {
  const mapa = {};
  const limiteLinhasCabecalho = Math.min(15, planilha.rowCount || 15);

  for (let coluna = 1; coluna <= totalColunas; coluna += 1) {
    let textoColuna = "";

    for (let linha = 1; linha <= limiteLinhasCabecalho; linha += 1) {
      const cell = planilha.getCell(linha, coluna);
      textoColuna += ` ${valorParaPreview(cell.value, cell.text)}`;
    }

    mapa[coluna] = colunaPareceTempo(textoColuna);
  }

  return mapa;
}

function mapearColunasProvaveisSheetJS(planilha, range) {
  const mapa = {};
  if (!range) {
    return mapa;
  }

  const ultimaLinhaCabecalho = Math.min(range.e.r, range.s.r + 14);

  for (let coluna = range.s.c; coluna <= range.e.c; coluna += 1) {
    let textoColuna = "";

    for (let linha = range.s.r; linha <= ultimaLinhaCabecalho; linha += 1) {
      const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
      const cell = planilha[endereco];
      if (cell) {
        textoColuna += ` ${valorParaPreview(cell.v, cell.w)}`;
      }
    }

    mapa[coluna + 1] = colunaPareceTempo(textoColuna);
  }

  return mapa;
}

function colunaPareceTempo(texto) {
  const normalizado = normalizarTexto(texto);
  if (!normalizado) {
    return false;
  }

  return TERMOS_COLUNAS_TEMPO.some((termo) => normalizado.includes(termo));
}

function obterDimensoesExcelJS(planilha) {
  let maxLinha = 0;
  let maxColuna = 0;

  planilha.eachRow({ includeEmpty: false }, (row, numeroLinha) => {
    maxLinha = Math.max(maxLinha, numeroLinha);
    maxColuna = Math.max(maxColuna, row.cellCount || 0);

    row.eachCell({ includeEmpty: false }, (cell, numeroColuna) => {
      if (valorTemConteudo(cell.value)) {
        maxColuna = Math.max(maxColuna, numeroColuna);
      }
    });
  });

  return {
    linhas: Math.max(
      maxLinha,
      planilha.actualRowCount || 0,
      planilha.rowCount || 0,
    ),
    colunas: Math.max(
      maxColuna,
      planilha.actualColumnCount || 0,
      planilha.columnCount || 0,
    ),
  };
}

function obterRangeSheetJS(planilha) {
  if (!planilha || !planilha["!ref"]) {
    return null;
  }

  return XLSX.utils.decode_range(planilha["!ref"]);
}

function valorTemConteudo(valor) {
  if (valor === null || valor === undefined) {
    return false;
  }

  if (typeof valor === "string") {
    return valor.length > 0;
  }

  return true;
}

function extrairValorSimples(valor) {
  if (!valor || typeof valor !== "object" || valor instanceof Date) {
    return valor;
  }

  if (Array.isArray(valor.richText)) {
    return valor.richText.map((parte) => parte.text || "").join("");
  }

  if (Object.prototype.hasOwnProperty.call(valor, "text")) {
    return valor.text;
  }

  if (Object.prototype.hasOwnProperty.call(valor, "result")) {
    return valor.result;
  }

  return valor;
}

function valorParaPreview(valor, textoFormatado = "") {
  const simples = extrairValorSimples(valor);

  if (simples === null || simples === undefined) {
    return "";
  }

  if (simples instanceof Date) {
    return simples.toLocaleDateString("pt-BR");
  }

  if (typeof simples === "object") {
    return textoFormatado || JSON.stringify(simples);
  }

  return String(simples);
}

function enderecoExcel(linha, coluna) {
  let numero = coluna;
  let letras = "";

  while (numero > 0) {
    const resto = (numero - 1) % 26;
    letras = String.fromCharCode(65 + resto) + letras;
    numero = Math.floor((numero - 1) / 26);
  }

  return `${letras}${linha}`;
}

function arquivoEhExcel(arquivo) {
  const extensao = obterExtensao(arquivo.name);
  return extensao === "xlsx" || extensao === "xls";
}

function obterExtensao(nomeArquivo) {
  return String(nomeArquivo).split(".").pop().toLowerCase();
}

function normalizarTexto(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9% ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function adicionarInvalido(lista, valor) {
  if (lista.length < 20 && !lista.includes(valor)) {
    lista.push(valor);
  }
}

function registrarAlerta(tipo, texto) {
  estado.alertas.push({ tipo, texto });
}

function mostrarAlertas(alertas = estado.alertas) {
  elementos.alertsBox.innerHTML = "";

  alertas.forEach((alerta) => {
    const item = document.createElement("div");
    item.className = `alert ${alerta.tipo}`;
    item.textContent = alerta.texto;
    elementos.alertsBox.appendChild(item);
  });
}

function atualizarResumo() {
  elementos.statSheet.textContent = estado.nomeAbaProcessada || "-";
  elementos.statRows.textContent = String(estado.linhas);
  elementos.statColumns.textContent = String(estado.colunas);
  elementos.statChanges.textContent = String(estado.celulasCorrigidas);
  elementos.previewHint.textContent = estado.nomeAbaProcessada
    ? `Aba: ${estado.nomeAbaProcessada}`
    : "Processando a aba escolhida";
}

function renderizarPreview() {
  elementos.previewContainer.innerHTML = "";

  if (!estado.preview.length) {
    const vazio = document.createElement("div");
    vazio.className = "empty-preview";
    vazio.textContent =
      "A aba processada não possui dados para pré-visualizar.";
    elementos.previewContainer.appendChild(vazio);
    return;
  }

  const tabela = document.createElement("table");
  const corpo = document.createElement("tbody");

  estado.preview.forEach((linha) => {
    const tr = document.createElement("tr");
    linha.forEach((valor) => {
      const td = document.createElement("td");
      td.textContent = valor;
      tr.appendChild(td);
    });
    corpo.appendChild(tr);
  });

  tabela.appendChild(corpo);
  elementos.previewContainer.appendChild(tabela);
}

function mostrarMensagem(tipo, texto) {
  elementos.statusMessage.className = `status-message ${tipo}`;
  elementos.statusMessage.textContent = texto;
}

function bloquearInterface(bloquear) {
  elementos.processButton.disabled = bloquear || !estado.arquivo;
  elementos.downloadButton.disabled = bloquear || !estado.processado;
}

function limparResultado(limparArquivo = true) {
  estado.tipoLeitura = null;
  estado.workbook = null;
  estado.abaProcessada = null;
  estado.nomeAbaProcessada = "";
  estado.alertas = [];
  estado.celulasCorrigidas = 0;
  estado.linhas = 0;
  estado.colunas = 0;
  estado.preview = [];
  estado.processado = false;

  if (limparArquivo) {
    estado.arquivo = null;
  }

  elementos.downloadButton.disabled = true;
  elementos.alertsBox.innerHTML = "";
  elementos.statSheet.textContent = "-";
  elementos.statRows.textContent = "0";
  elementos.statColumns.textContent = "0";
  elementos.statChanges.textContent = "0";
  elementos.previewHint.textContent = "Processando a aba escolhida";
  elementos.previewContainer.innerHTML =
    '<div class="empty-preview">A prévia aparecerá aqui depois do processamento.</div>';
  elementos.statusMessage.className = "status-message";
  elementos.statusMessage.textContent = "";
}
