import dotenv from "dotenv";
dotenv.config();
import { getAuth, infoAPI } from "./getAuth.js";
import { formatDate } from "./utils/formatDate.js";
import nodemailer from "nodemailer";
import ExcelJS from "exceljs";
import cron from "node-cron";
import pkg from "whatsapp-web.js";
const { Client, LocalAuth, MessageMedia } = pkg;
import qrcode from "qrcode-terminal";
import fs from "fs";
import path from "path";

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Formata Date para YYYY-MM-DD em horário local (sem timezone)
 */
function formatYMD(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Retorna array de datas no formato YYYY-MM-DD entre dataInicio e dataFim (inclusive)
 * Parsing seguro em local-time para evitar off-by-one por timezone
 */
function getDateRange(dataInicio, dataFim) {
  const [y1, m1, d1] = dataInicio.split("-").map(Number);
  const [y2, m2, d2] = dataFim.split("-").map(Number);
  const start = new Date(y1, m1 - 1, d1);
  const end = new Date(y2, m2 - 1, d2);

  const dates = [];
  const current = new Date(start);
  while (current <= end) {
    dates.push(formatYMD(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

/**
 * Retorna o intervalo (segunda a sexta) da semana ANTERIOR em horário local
 */
function getLastWeekRange() {
  const today = new Date();
  const localToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const dow = localToday.getDay(); // 0=domingo, 1=segunda, ... 6=sábado
  const diffToMonday = dow === 0 ? -6 : 1 - dow; // deslocamento para a segunda da semana corrente
  const currentWeekMonday = new Date(localToday);
  currentWeekMonday.setDate(localToday.getDate() + diffToMonday);

  const lastWeekMonday = new Date(currentWeekMonday);
  lastWeekMonday.setDate(currentWeekMonday.getDate() - 7);

  const lastWeekFriday = new Date(lastWeekMonday);
  lastWeekFriday.setDate(lastWeekMonday.getDate() + 4);

  return { inicio: formatYMD(lastWeekMonday), fim: formatYMD(lastWeekFriday) };
}

/**
 * Retorna a data da última segunda-feira (início da semana)
 */
function getLastMonday() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0 = domingo, 1 = segunda, ..., 6 = sábado
  const daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1; // Se domingo, volta 6 dias
  const monday = new Date(today);
  monday.setDate(today.getDate() - daysToSubtract - 7); // Semana passada
  
  const yyyy = monday.getFullYear();
  const mm = String(monday.getMonth() + 1).padStart(2, "0");
  const dd = String(monday.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Retorna a data da última sexta-feira (fim da semana)
 */
function getLastFriday() {
  const monday = new Date(getLastMonday());
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4); // Segunda + 4 dias = Sexta
  
  const yyyy = friday.getFullYear();
  const mm = String(friday.getMonth() + 1).padStart(2, "0");
  const dd = String(friday.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Retorna primeiro e último dia do mês anterior
 */
function getLastMonthRange() {
  const today = new Date();
  const firstDayLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const lastDayLastMonth = new Date(today.getFullYear(), today.getMonth(), 0);
  
  const formatDate = (date) => {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  };
  
  return {
    inicio: formatDate(firstDayLastMonth),
    fim: formatDate(lastDayLastMonth),
    mesNome: firstDayLastMonth.toLocaleString('pt-BR', { month: 'long', year: 'numeric' })
  };
}

/**
 * Formata data para formato ISO (YYYY-MM-DDTHH:mm:ssZ)
 */
function formatISODate(date) {
  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}T00:00:00Z`;
}

/**
 * Filtra produções do dia anterior baseado na data de vigência inicial
 */
function filtrarPorDiaAnterior(dataAnterior, producoes) {
  return producoes.filter((producao) => {
    const dataVigencia = producao.dataVigenciaInicial;
    if (!dataVigencia) return false;
    // Compara apenas a data (sem hora)
    const dataVigenciaStr = dataVigencia.split("T")[0]; // Remove hora se houver
    return dataVigenciaStr === dataAnterior;
  });
}

/**
 * Filtra produções do dia anterior baseado na data emitida
 */
function filtrarPorDataEmitidaDiaAnterior(dataAnterior, producoes) {
  return producoes.filter((producao) => {
    const dataEmitida = producao.dataEmitida;
    if (!dataEmitida) return false;
    // Compara apenas a data (sem hora)
    const dataEmitidaStr = dataEmitida.split("T")[0]; // Remove hora se houver
    return dataEmitidaStr === dataAnterior;
  });
}

/**
 * Filtra sinistros do dia anterior
 */
function filtrarSinistrosPorDiaAnterior(dataAnterior, sinistros) {
  return sinistros.filter((sinistro) => {
    // Assumindo que o sinistro tem um campo de data (pode ser dataAbertura, dataCadastro, dataAviso, etc)
    const dataSinistro = sinistro.dataAviso;
    if (!dataSinistro) return false;
    const dataSinistroStr = String(dataSinistro).split("T")[0];
    return dataSinistroStr === dataAnterior;
  });
}

/**
 * Busca produção completa de um período (genérico para dia, semana ou mês)
 */
async function getProducaoCompletaPeriodo(token, dataInicial, dataFinal) {
  try {
    const allResults = [];
    let page = 1;
    let nextPage = null;

    // Se data inicial e final forem iguais, adiciona 1 dia à data final para a API
    // mas mantém o filtro com a data original
    let dataFinalParaAPI = dataFinal;
    if (dataInicial === dataFinal) {
      const dataFinalDate = new Date(dataFinal + "T00:00:00");
      dataFinalDate.setDate(dataFinalDate.getDate() + 1);
      const yyyy = dataFinalDate.getFullYear();
      const mm = String(dataFinalDate.getMonth() + 1).padStart(2, "0");
      const dd = String(dataFinalDate.getDate()).padStart(2, "0");
      dataFinalParaAPI = `${yyyy}-${mm}-${dd}`;
      console.log(`Ajustando busca: API receberá ${dataInicial} até ${dataFinalParaAPI}, mas filtrando apenas ${dataInicial}`);
    }

    console.log(`Buscando produção completa de ${dataInicial} até ${dataFinalParaAPI}...`);

    do {
      const response = await fetch(
        `${infoAPI.url}/producao/pesquisar?page=${page}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            tipoData: "dataVigenciaInicial",
            dataInicial: dataInicial,
            dataFinal: dataFinalParaAPI, // Usa a data ajustada para a API
            nivel: ["1", "2"],
            tipo: ["0", "2", "4"],
            status: ["0", null, "", "1", "3", "4", "5", "6", "7"],
          }),
        }
      );

      if (!response.ok) {
        throw new Error(
          `Erro na requisição de produção completa: ${response.status}`
        );
      }

      const body = await response.json();
      nextPage = body?.links?.next ?? null;
      const producaoData = body?.data ?? [];

      allResults.push(...producaoData);
      page++;
      await delay(1000);
    } while (nextPage !== null);

    // Filtra apenas os registros do período ORIGINAL (não da data ajustada)
    const datesInRange = getDateRange(dataInicial, dataFinal); // Usa datas originais
    const filtered = allResults.filter((producao) => {
      const dataVigencia = producao.dataVigenciaInicial;
      if (!dataVigencia) return false;
      const dataVigenciaStr = dataVigencia.split("T")[0];
      return datesInRange.includes(dataVigenciaStr);
    });

    return filtered;
  } catch (err) {
    console.error("Erro ao buscar produção completa do período:", err);
    throw err;
  }
}

/**
 * Busca produção completa do dia anterior (para a planilha)
 */
async function getProducaoCompletaDiaAnterior(token, dataAnterior) {
  return await getProducaoCompletaPeriodo(token, dataAnterior, dataAnterior);
}

/**
 * Busca todas as produções (transmissões) do dia anterior
 */
async function getTransmissoesDiaAnterior(token, dataAnterior) {
  try {
    const allResults = [];
    let page = 1;
    let nextPage = null;

    // Calcula data final (dia seguinte ao anterior, ou hoje)
    const dataFinal = formatDate(0); // hoje

    do {
      const response = await fetch(
        `${infoAPI.url}/producao/pesquisar?page=${page}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            tipoData: "dataVigenciaInicial",
            dataInicial: dataAnterior,
            dataFinal: dataFinal,
            nivel: ["1", "2"],
            tipo: ["0", "2", "4"],
            status: ["0", null, "", "1", "3", "4", "5", "6", "7"],
          }),
        }
      );

      if (!response.ok) {
        throw new Error(
          `Erro na requisição de transmissões: ${response.status}`
        );
      }

      const body = await response.json();
      nextPage = body?.links?.next ?? null;
      const producaoData = body?.data ?? [];

      allResults.push(...producaoData);
      page++;
      await delay(1000);
    } while (nextPage !== null);

    // Filtra apenas os registros do dia anterior
    return filtrarPorDiaAnterior(dataAnterior, allResults);
  } catch (err) {
    console.error("Erro ao buscar transmissões:", err);
    throw err;
  }
}

/**
 * Busca apólices emitidas do dia anterior
 */
async function getApolicesEmitidasDiaAnterior(token, dataAnterior) {
  try {
    const allResults = [];
    let page = 1;
    let nextPage = null;

    // Calcula data final (dia seguinte ao anterior, ou hoje)
    const dataFinal = formatDate(0); // hoje

    do {
      const response = await fetch(
        `${infoAPI.url}/producao/pesquisar?page=${page}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            tipoData: "dataEmitida",
            dataInicial: dataAnterior,
            dataFinal: dataFinal,
            nivel: ["1", "2"],
            tipo: ["0", "2", "4"],
            status: ["0", null, "", "1", "3", "4", "5", "6", "7"],
          }),
        }
      );

      if (!response.ok) {
        throw new Error(
          `Erro na requisição de apólices emitidas: ${response.status}`
        );
      }

      const body = await response.json();
      nextPage = body?.links?.next ?? null;
      const producaoData = body?.data ?? [];

      allResults.push(...producaoData);
      page++;
      await delay(1000);
    } while (nextPage !== null);

    // Filtra apenas os registros do dia anterior baseado na dataEmitida
    return filtrarPorDataEmitidaDiaAnterior(dataAnterior, allResults);
  } catch (err) {
    console.error("Erro ao buscar apólices emitidas:", err);
    throw err;
  }
}

/**
 * Busca assistências urgentes do SULTS do dia anterior
 */
async function getAssistenciasUrgentesDiaAnterior(dataAnterior) {
  try {
    const baseUrl = "https://api.sults.com.br/api/v1/chamado/ticket";
    // Tenta ambos os nomes de variáveis de ambiente (compatibilidade)
    const accessToken = process.env.SULTS_ACCESS_TOKEN;

    if (!accessToken) {
      console.warn(
        "SULTS_ACCESS_TOKEN não configurado. Pulando busca de assistências urgentes."
      );
      return [];
    }

    // Busca uma janela maior (7 dias antes do dia anterior) para garantir que pegue tudo
    // Isso evita problemas de timezone ou delays na API
    const dataAnteriorDate = new Date(dataAnterior + "T00:00:00");
    dataAnteriorDate.setDate(dataAnteriorDate.getDate() - 7); // 7 dias antes
    const abertoStart = formatISODate(dataAnteriorDate);

    console.log(
      `Buscando assistências urgentes desde: ${abertoStart} até encontrar do dia ${dataAnterior}`
    );

    let assistencias = [];
    let totalPage;
    let page = 0;
    const limit = 100;
    let totalProcessados = 0;

    do {
      const url = `${baseUrl}?start=${page}&limit=${limit}&abertoStart=${encodeURIComponent(
        abertoStart
      )}`;

      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: accessToken,
          "Content-Type": "application/json;charset=UTF-8",
        },
      });

      if (!response.ok) {
        console.error(
          `Erro na requisição de assistências urgentes: ${response.status}`
        );
        break;
      }

      const result = await response.json();
      const chamados = result.data || [];
      totalPage = result.totalPage ?? page;
      totalProcessados += chamados.length;

      // Primeiro filtra apenas assistências urgentes
      const assistenciasUrgentes = chamados.filter(
        (item) =>
          item.titulo?.includes("Solicitação: Assistência Urgente") ||
          item.titulo?.includes("Solicitação: Assistência urgente")
      );

      // Depois filtra apenas as do dia anterior
      const filtrados = assistenciasUrgentes.filter((item) => {
        const dataAberto = item.aberto;
        if (!dataAberto) return false;

        const dataAbertoStr = dataAberto.split("T")[0]; // Remove hora
        return dataAbertoStr === dataAnterior;
      });

      // Adiciona os filtrados evitando duplicatas
      for (const item of filtrados) {
        if (!assistencias.find((a) => a.id === item.id)) {
          assistencias.push(item);
          console.log(
            `Assistência urgente encontrada: ID ${
              item.id
            } - ${item.titulo?.substring(0, 50)}...`
          );
        }
      }

      console.log(
        `Página ${page}/${totalPage} | Processados: ${totalProcessados} | Urgentes do dia ${dataAnterior}: ${assistencias.length}`
      );

      // Se já passou do dia anterior (mais recente), pode parar
      // Verifica se algum chamado tem data mais recente que o dia anterior
      const temDataMaisRecente = chamados.some((item) => {
        if (!item.aberto) return false;
        const dataAbertoStr = item.aberto.split("T")[0];
        return dataAbertoStr > dataAnterior;
      });

      if (temDataMaisRecente && assistencias.length > 0) {
        console.log("Encontrou chamados mais recentes, pode parar a busca.");
        break;
      }

      page++;
      await delay(500);
    } while (page <= totalPage);

    console.log(
      `Total de assistências urgentes do dia ${dataAnterior}: ${assistencias.length}`
    );

    return assistencias;
  } catch (err) {
    console.error("Erro ao buscar assistências urgentes:", err);
    console.error("Stack trace:", err.stack);
    // Não lançar erro para não interromper o processo principal
    return [];
  }
}

/**
 * Busca sinistros abertos do dia anterior
 */
async function getSinistrosAbertosDiaAnterior(token, dataAnterior) {
  try {
    const allResults = [];
    let page = 1;
    let nextPage = null;

    // Calcula data final (dia seguinte ao anterior, ou hoje)
    const dataFinal = formatDate(0); // hoje

    do {
      const response = await fetch(
        `${infoAPI.url}/sinistros/pesquisar?page=${page}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            tipoData: "dataAviso",
            dataInicial: dataAnterior,
            dataFinal: dataFinal,
          }),
        }
      );

      if (!response.ok) {
        throw new Error(`Erro na requisição de sinistros: ${response.status}`);
      }

      const body = await response.json();
      nextPage = body?.links?.next ?? null;
      const sinistrosData = body?.data ?? [];

      allResults.push(...sinistrosData);
      page++;
      await delay(1000);
    } while (nextPage !== null);

    // Filtra apenas os sinistros do dia anterior
    return filtrarSinistrosPorDiaAnterior(dataAnterior, allResults);
  } catch (err) {
    console.error("Erro ao buscar sinistros:", err);
    throw err;
  }
}

/**
 * Gera planilha Excel completa com abas para Produção, Sinistros e Assistências Urgentes
 */
async function gerarPlanilhaExcelCompleta(producaoData, sinistrosData, assistenciasData, periodoNome) {
  const workbook = new ExcelJS.Workbook();

  // ===== ABA 1: PRODUÇÃO =====
  const wsProd = workbook.addWorksheet("Produção");
  wsProd.columns = [
    { header: "ID Proposta", key: "propostaId", width: 15 },
    { header: "Data Vigência Inicial", key: "dataVigenciaInicial", width: 20 },
    { header: "Data Vigência Final", key: "dataVigenciaFinal", width: 20 },
    { header: "Data Emitida", key: "dataEmitida", width: 20 },
    { header: "Nível", key: "nivelLabel", width: 15 },
    { header: "Tipo", key: "tipoLabel", width: 20 },
    { header: "Status", key: "statusLabel", width: 20 },
    { header: "Comissão", key: "comissao", width: 15 },
    { header: "Prêmio Líquido", key: "premioLiquido", width: 18 },
    { header: "Prêmio Total", key: "premioTotal", width: 18 },
    { header: "Parcelas", key: "parcelas", width: 10 },
    { header: "Nome Corretor", key: "nomeCorretor", width: 30 },
    { header: "Nome Segurado", key: "nomeSegurado", width: 30 },
    { header: "Tipo Pessoa", key: "tipoPessoa", width: 15 },
    { header: "Sexo Segurado", key: "sexoSegurado", width: 15 },
    { header: "Ramo", key: "ramo", width: 20 },
    { header: "Seguradora", key: "seguradora", width: 30 },
  ];

  wsProd.getRow(1).font = { bold: true };
  wsProd.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4A04A5" },
  };
  wsProd.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

  producaoData.forEach((item) => {
    wsProd.addRow({
      propostaId: item.propostaId || "",
      dataVigenciaInicial: item.dataVigenciaInicial || "",
      dataVigenciaFinal: item.dataVigenciaFinal || "",
      dataEmitida: item.dataEmitida || "",
      nivelLabel: item.nivelLabel || "",
      tipoLabel: item.tipoLabel || "",
      statusLabel: item.statusLabel || "",
      comissao: item.comissao || "",
      premioLiquido: item.premioLiquido || "",
      premioTotal: item.premioTotal || "",
      parcelas: item.parcelas || "",
      nomeCorretor: item.corretores?.[0]?.nome || "",
      nomeSegurado: item.segurado?.nome || "",
      tipoPessoa: item.segurado?.tipoPessoaLabel || "",
      sexoSegurado: item.segurado?.sexoLabel || "",
      ramo: item.ramo?.nome || "",
      seguradora: item.companhia?.nome || "",
    });
  });

  // ===== ABA 2: SINISTROS =====
  const wsSin = workbook.addWorksheet("Sinistros");
  wsSin.columns = [
    { header: "ID Sinistro", key: "sinistroId", width: 15 },
    { header: "Valor Indenizado", key: "valorIndenizado", width: 18 },
    { header: "Data Aviso", key: "dataAviso", width: 20 },
    { header: "Data Sinistro", key: "dataSinistro", width: 20 },
    { header: "Data Vistoria", key: "dataVistoria", width: 20 },
    { header: "Data Pagamento", key: "dataPagamento", width: 20 },
    { header: "Data Autorização Reparos", key: "dataAutorizacaoReparos", width: 25 },
    { header: "Data Envio NF", key: "dataEnvioNF", width: 20 },
    { header: "Data Documentação", key: "dataDocumentacao", width: 22 },
    { header: "Corretor", key: "corretor", width: 30 },
    { header: "Seguradora", key: "seguradora", width: 30 },
    { header: "Segurado", key: "segurado", width: 35 },
    { header: "CPF/CNPJ", key: "cpfCnpj", width: 18 },
    { header: "Status", key: "status", width: 20 },
    { header: "Ramo", key: "ramo", width: 20 },
    { header: "Produtor", key: "produtor", width: 30 },
    { header: "Tipo", key: "tipo", width: 25 },
  ];

  wsSin.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  wsSin.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4A04A5" },
  };

  sinistrosData.forEach((item) => {
    wsSin.addRow({
      sinistroId: item.sinistroId || null,
      valorIndenizado: item.valorIndenizado || null,
      dataAviso: item.dataAviso || null,
      dataSinistro: item.dataSinistro || null,
      dataVistoria: item.dataVistoria || null,
      dataPagamento: item.dataPagamento || null,
      dataAutorizacaoReparos: item.dataAutorizacaoReparos || null,
      dataEnvioNF: item.dataEnvioNF || null,
      dataDocumentacao: item.dataDocumentacao || null,
      corretor: item.proposta?.corretores?.[0]?.nome || null,
      seguradora: item.companhia?.nome || null,
      segurado: item.proposta?.segurado?.nome || null,
      cpfCnpj: item.proposta?.segurado?.cpf_cnpj || null,
      status: item.statusSinistro?.nome || null,
      ramo: item.proposta?.ramo?.nome || null,
      produtor: item.proposta?.repasses?.[0]?.produtor?.nome || null,
      tipo: item.tipo?.nome || null,
    });
  });

  // ===== ABA 3: ASSISTÊNCIAS URGENTES =====
  const wsAss = workbook.addWorksheet("Assistências Urgentes");
  wsAss.columns = [
    { header: "ID", key: "id", width: 12 },
    { header: "Título", key: "titulo", width: 50 },
    { header: "Solicitante", key: "solicitante", width: 30 },
    { header: "Responsável", key: "responsavel", width: 30 },
    { header: "Unidade", key: "unidade", width: 25 },
    { header: "Departamento", key: "departamento", width: 25 },
    { header: "Data Abertura", key: "aberto", width: 20 },
    { header: "Situação", key: "situacao", width: 15 },
    { header: "Primeira Interação", key: "primeiraInteracao", width: 20 },
    { header: "Última Alteração", key: "ultimaAlteracao", width: 20 },
  ];

  wsAss.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  wsAss.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4A04A5" },
  };

  assistenciasData.forEach((item) => {
    wsAss.addRow({
      id: item.id || "",
      titulo: item.titulo || "",
      solicitante: item.solicitante?.nome || "",
      responsavel: item.responsavel?.nome || "",
      unidade: item.unidade?.nome || "",
      departamento: item.departamento?.nome || "",
      aberto: item.aberto || "",
      situacao: item.situacao || "",
      primeiraInteracao: item.primeiraInteracao || "",
      ultimaAlteracao: item.ultimaAlteracao || "",
    });
  });

  // Salvar arquivo
  const fileName = `relatorio_completo_${periodoNome.replace(/[\/\s]/g, "_")}.xlsx`;
  const filePath = path.join(process.cwd(), fileName);
  await workbook.xlsx.writeFile(filePath);

  return { filePath, fileName };
}

/**
 * Gera planilha Excel com a produção completa do dia anterior (retrocompatibilidade)
 */
async function gerarPlanilhaExcel(producaoData, dataAnterior) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Produção do Dia Anterior");

  // Cabeçalhos
  worksheet.columns = [
    { header: "ID Proposta", key: "propostaId", width: 15 },
    { header: "Data Vigência Inicial", key: "dataVigenciaInicial", width: 20 },
    { header: "Data Vigência Final", key: "dataVigenciaFinal", width: 20 },
    { header: "Data Emitida", key: "dataEmitida", width: 20 },
    { header: "Nível", key: "nivelLabel", width: 15 },
    { header: "Tipo", key: "tipoLabel", width: 20 },
    { header: "Status", key: "statusLabel", width: 20 },
    { header: "Comissão", key: "comissao", width: 15 },
    { header: "Prêmio Líquido", key: "premioLiquido", width: 18 },
    { header: "Prêmio Total", key: "premioTotal", width: 18 },
    { header: "Parcelas", key: "parcelas", width: 10 },
    { header: "Nome Corretor", key: "nomeCorretor", width: 30 },
    { header: "Nome Segurado", key: "nomeSegurado", width: 30 },
    { header: "Tipo Pessoa", key: "tipoPessoa", width: 15 },
    { header: "Sexo Segurado", key: "sexoSegurado", width: 15 },
    { header: "Ramo", key: "ramo", width: 20 },
    { header: "Seguradora", key: "seguradora", width: 30 },
    { header: "Apólice", key: "apolice", width: 30 },
  ];

  // Estilizar cabeçalho
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0E0E0" },
  };

  // Adicionar dados
  producaoData.forEach((item) => {
    worksheet.addRow({
      propostaId: item.propostaId || "",
      dataVigenciaInicial: item.dataVigenciaInicial || "",
      dataVigenciaFinal: item.dataVigenciaFinal || "",
      dataEmitida: item.dataEmitida || "",
      nivelLabel: item.nivelLabel || "",
      tipoLabel: item.tipoLabel || "",
      statusLabel: item.statusLabel || "",
      comissao: item.comissao || "",
      premioLiquido: item.premioLiquido || "",
      premioTotal: item.premioTotal || "",
      parcelas: item.parcelas || "",
      nomeCorretor: item.corretores?.[0]?.nome || "",
      nomeSegurado: item.segurado?.nome || "",
      tipoPessoa: item.segurado?.tipoPessoaLabel || "",
      qtdeProducoes: item.segurado?.qtdeRegistrosProducao || "",
      sexoSegurado: item.segurado?.sexoLabel || "",
      ramo: item.ramo?.nome || "",
      seguradora: item.companhia?.nome || "",
    });
  });

  // Salvar arquivo
  const fileName = `producao_${dataAnterior.replace(/-/g, "_")}.xlsx`;
  const filePath = path.join(process.cwd(), fileName);
  await workbook.xlsx.writeFile(filePath);

  return { filePath, fileName };
}

/**
 * Função centralizada para enviar e-mail de erro/notificação
 */
async function enviarEmailErro(titulo, mensagem, erro, contexto = {}) {
  try {
    const email = process.env.MAIL_EMAIL;
    const password = process.env.MAIL_PASSWORD;
    const emailDestinatario = process.env.DIRETOR_EMAIL || email;

    if (!email || !password) {
      console.error("⚠️ Não foi possível enviar e-mail de erro: MAIL_EMAIL ou MAIL_PASSWORD não configurados");
      return;
    }

    const transporter = nodemailer.createTransport({
      host: "smtp.dreamhost.com",
      port: 587,
      secure: false,
      auth: { user: email, pass: password },
      tls: { rejectUnauthorized: false },
    });

    const erroDetalhado = erro 
      ? `\n\n<strong>Erro:</strong>\n<pre style="white-space: pre-wrap; background: #f5f5f5; padding: 10px; border-radius: 4px;">${String(erro)}</pre>\n\n<strong>Stack Trace:</strong>\n<pre style="white-space: pre-wrap; background: #f5f5f5; padding: 10px; border-radius: 4px;">${erro.stack || 'N/A'}</pre>`
      : '';

    const contextoHtml = Object.keys(contexto).length > 0
      ? `\n\n<strong>Contexto:</strong>\n<ul>${Object.entries(contexto).map(([k, v]) => `<li><strong>${k}:</strong> ${v}</li>`).join('')}</ul>`
      : '';

    const html = `
      <!DOCTYPE html>
      <html>
      <head><meta charset="UTF-8"></head>
      <body style="font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5;">
        <div style="max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
          <h1 style="color: #d32f2f; margin-top: 0;">⚠️ ${titulo}</h1>
          <p style="font-size: 16px; line-height: 1.6;">${mensagem}</p>
          <p style="color: #666; font-size: 14px;"><strong>Data/Hora:</strong> ${new Date().toLocaleString("pt-BR")}</p>
          ${erroDetalhado}
          ${contextoHtml}
        </div>
      </body>
      </html>
    `;

    await transporter.sendMail({
      from: `"Sistema de Relatórios - Avantar" <${email}>`,
      to: emailDestinatario,
      subject: `🚨 ${titulo} - ${new Date().toLocaleDateString("pt-BR")}`,
      html,
    });

    console.log(`✅ E-mail de erro enviado para: ${emailDestinatario}`);
  } catch (emailErr) {
    console.error("❌ Erro crítico: Não foi possível enviar e-mail de notificação de erro:", emailErr);
    console.error("Erro original que deveria ser notificado:", erro);
  }
}

/**
 * Envia relatório por e-mail
 */
async function enviarEmail(
  transmissoes,
  apolicesEmitidas,
  sinistros,
  assistenciasUrgentes,
  arquivoExcel,
  dataAnterior
) {
  try {
    const email = process.env.MAIL_EMAIL;
    const password = process.env.MAIL_PASSWORD;
    const emailDestinatario = process.env.DIRETOR_EMAIL || email;

    // Verificar se as variáveis estão definidas
    if (!email || !password) {
      throw new Error(
        "MAIL_EMAIL ou MAIL_PASSWORD não estão configurados no .env"
      );
    }

    if (!emailDestinatario) {
      throw new Error("DIRETOR_EMAIL não está configurado no .env");
    }

    // Verificar se o arquivo existe antes de anexar
    if (!fs.existsSync(arquivoExcel.filePath)) {
      throw new Error(`Arquivo Excel não encontrado: ${arquivoExcel.filePath}`);
    }

    console.log(`Configurando envio de e-mail para: ${emailDestinatario}`);
    console.log(`Usando SMTP: smtp.dreamhost.com:587`);

    const transporter = nodemailer.createTransport({
      host: "smtp.dreamhost.com",
      port: 587,
      secure: false, // true para 465, false para outras portas
      auth: {
        user: email,
        pass: password,
      },
      tls: {
        // Não falhar em certificados inválidos
        rejectUnauthorized: false,
      },
    });

    // Verificar conexão antes de enviar
    await transporter.verify();
    console.log("Conexão SMTP verificada com sucesso");

    const html = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Diário - Centro de Operações</title>
</head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background-color: #f5f5f5;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
        <tr>
            <td style="padding: 30px 15px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
                    
                    <!-- Header simples -->
                    <tr>
                        <td style="border-bottom: 3px solid #4A04A5; flex: 1; justify-items: center">
                            <img src="https://iili.io/KZqR9cu.png" alt="Avantar" style="width: 200px; height: 200px; display: block;" />
                        </td>
                    </tr>

                    <!-- Título e data -->
                    <tr>
                        <td style="padding: 35px 40px 25px 40px;">
                            <h1 style="margin: 0; font-size: 24px; font-weight: 400; color: #170138; letter-spacing: -0.3px;">Relatório Diário</h1>
                            <p style="margin: 8px 0 0 0; font-size: 14px; color: #666666; font-weight: 400;">Centro de Operações</p>
                        </td>
                    </tr>

                    <!-- Informações de data -->
                    <tr>
                        <td style="padding: 0 40px 30px 40px;">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="padding: 0; border-bottom: 1px solid #e0e0e0;">
                                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 12px 0; font-size: 13px; color: #666666;">Data de referência</td>
                                                <td style="padding: 12px 0; font-size: 13px; color: #170138; text-align: right; font-weight: 500;">${dataAnterior}</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 0;">
                                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 12px 0; font-size: 13px; color: #666666;">Gerado em</td>
                                                <td style="padding: 12px 0; font-size: 13px; color: #170138; text-align: right; font-weight: 500;">${new Date().toLocaleString(
                                                  "pt-BR"
                                                )}</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- Espaçamento -->
                    <tr>
                        <td style="padding: 15px 0;"></td>
                    </tr>

                    <!-- Dados do relatório -->
                    <tr>
                        <td style="padding: 0 40px;">
                            <h2 style="margin: 0 0 20px 0; font-size: 16px; font-weight: 500; color: #170138; text-transform: uppercase; letter-spacing: 0.5px;">Resumo do dia anterior</h2>
                            
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="padding: 0; border-bottom: 1px solid #e0e0e0;">
                                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 18px 0; font-size: 14px; color: #666666;">Transmissões</td>
                                                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; font-weight: 300; letter-spacing: -0.5px;">${
                                                  transmissoes.length
                                                }</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 0; border-bottom: 1px solid #e0e0e0;">
                                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 18px 0; font-size: 14px; color: #666666;">Apólices emitidas</td>
                                                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; font-weight: 300; letter-spacing: -0.5px;">${
                                                  apolicesEmitidas.length
                                                }</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 0; border-bottom: 1px solid #e0e0e0;">
                                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 18px 0; font-size: 14px; color: #666666;">Sinistros abertos</td>
                                                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; font-weight: 300; letter-spacing: -0.5px;">${
                                                  sinistros.length
                                                }</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 0;">
                                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 18px 0;">
                                                    <p style="margin: 0; font-size: 14px; color: #666666;">Assistências urgentes</p>
                                                    <p style="margin: 3px 0 0 0; font-size: 12px; color: #999999;">(SULTS)</p>
                                                </td>
                                                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; font-weight: 300; letter-spacing: -0.5px;">${
                                                  assistenciasUrgentes.length
                                                }</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- Nota sobre anexo -->
                    <tr>
                        <td style="padding: 35px 40px 40px 40px;">
                            <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.6;">
                                Em anexo, segue a planilha completa com a produção do dia anterior.
                            </p>
                        </td>
                    </tr>

                    <!-- Footer -->
                    <tr>
                        <td style="padding: 25px 40px; background-color: #fafafa; border-top: 1px solid #e0e0e0;">
                            <p style="margin: 0; font-size: 12px; color: #999999; text-align: center;">
                                Tecnologia Rede Avantar
                            </p>
                        </td>
                    </tr>
                    
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
    `;

    const info = await transporter.sendMail({
      from: `"Tecnologia Avantar" <${email}>`,
      to: emailDestinatario,
      subject: `Relatório Diário - Centro de Operações - ${dataAnterior}`,
      html,
      attachments: [
        {
          filename: arquivoExcel.fileName,
          path: arquivoExcel.filePath,
        },
      ],
    });

    console.log("E-mail enviado com sucesso!");
    console.log("ID da mensagem:", info.messageId);
    console.log("Resposta do servidor:", info.response);
    console.log(`E-mail enviado de: ${email}`);
    console.log(`E-mail enviado para: ${emailDestinatario}`);
    console.log(
      `Assunto: Relatório Diário - Centro de Operações - ${dataAnterior}`
    );
    console.log(
      `Arquivo anexado: ${arquivoExcel.fileName} (${
        fs.statSync(arquivoExcel.filePath).size
      } bytes)`
    );

    // Adicionar informações sobre possíveis problemas
    console.log("\n⚠️  Se o e-mail não chegou, verifique:");
    console.log("  1. Pasta de SPAM/Lixo Eletrônico");
    console.log("  2. Aguarde alguns minutos (pode estar na fila do servidor)");
    console.log("  3. Verifique se o endereço de destino está correto");
  } catch (err) {
    console.error("Erro ao enviar e-mail:");
    console.error("Mensagem:", err.message);
    if (err.response) {
      console.error("Resposta do servidor:", err.response);
    }
    if (err.code) {
      console.error("Código do erro:", err.code);
    }
    throw err;
  }
}

// Envia relatório por WhatsApp

let whatsappClient = null;

async function inicializarWhatsApp() {
  return new Promise(async (resolve, reject) => {
    if (whatsappClient && whatsappClient.info) {
      // Cliente já está pronto
      resolve(whatsappClient);
      return;
    }

    // Se o cliente existe mas não está pronto, aguardar
    if (whatsappClient) {
      whatsappClient.once("ready", () => {
        console.log("WhatsApp conectado!");
        resolve(whatsappClient);
      });
      return;
    }

    let isResolved = false;

    // Timeout de segurança: se não conectar em 5 minutos, rejeita
    const timeout = setTimeout(async () => {
      if (!isResolved) {
        isResolved = true;
        const erro = new Error("Timeout: WhatsApp não conectou em 5 minutos");
        console.error(erro.message);
        await enviarEmailErro(
          "Timeout na Inicialização do WhatsApp",
          "O WhatsApp não conseguiu conectar em 5 minutos. Isso pode indicar problema de conexão ou necessidade de escanear QR Code novamente.",
          erro
        );
        reject(erro);
      }
    }, 5 * 60 * 1000); // 5 minutos

    whatsappClient = new Client({
      authStrategy: new LocalAuth({
        dataPath: "./whatsapp-session",
      }),
      puppeteer: {
        headless: true,
        args: [
          "--no-sandbox",
          "--disable-setuid-sandbox",
          "--disable-dev-shm-usage",
          "--disable-accelerated-2d-canvas",
          "--no-first-run",
          "--no-zygote",
          "--disable-gpu",
        ],
      },
    });

    whatsappClient.on("qr", (qr) => {
      console.log("QR Code gerado para WhatsApp. Escaneie com o WhatsApp:");
      qrcode.generate(qr, { small: true });
    });

    whatsappClient.on("ready", () => {
      console.log("WhatsApp conectado!");
      if (!isResolved) {
        isResolved = true;
        clearTimeout(timeout); // Limpa o timeout
        // Aguardar um pouco mais para garantir que tudo está carregado
        setTimeout(() => {
          resolve(whatsappClient);
        }, 2000);
      }
    });

    whatsappClient.on("authenticated", () => {
      console.log("WhatsApp autenticado!");
    });

    whatsappClient.on("auth_failure", async (msg) => {
      console.error("Falha na autenticação do WhatsApp:", msg);
      if (!isResolved) {
        isResolved = true;
        clearTimeout(timeout); // Limpa o timeout
        const erro = new Error("Falha na autenticação do WhatsApp");
        // Notificar por e-mail sobre falha de autenticação
        await enviarEmailErro(
          "Falha na Autenticação do WhatsApp",
          "O sistema não conseguiu autenticar no WhatsApp. É necessário escanear o QR Code novamente.",
          erro,
          { motivo: msg || "Desconhecido" }
        );
        reject(erro);
      }
    });

    whatsappClient.on("disconnected", async (reason) => {
      console.log("WhatsApp desconectado:", reason);
      // Notificar desconexão apenas se não foi intencional
      if (reason !== "LOGOUT" && reason !== "NAVIGATION") {
        await enviarEmailErro(
          "WhatsApp Desconectado",
          "O WhatsApp foi desconectado inesperadamente.",
          null,
          { motivo: reason || "Desconhecido" }
        );
      }
      whatsappClient = null;
    });

    whatsappClient.initialize().catch(async (err) => {
      if (!isResolved) {
        isResolved = true;
        clearTimeout(timeout); // Limpa o timeout
        // Notificar erro na inicialização
        await enviarEmailErro(
          "Erro ao Inicializar WhatsApp",
          "Ocorreu um erro ao tentar inicializar o cliente WhatsApp.",
          err
        );
        reject(err);
      }
    });
  });
}

async function enviarWhatsApp(
  transmissoes,
  apolicesEmitidas,
  sinistros,
  assistenciasUrgentes,
  arquivoExcel,
  dataAnterior
) {
  try {
    const numeroDestinatario = process.env.DIRETOR_WHATSAPP;

    if (!numeroDestinatario) {
      console.log(
        "Número de WhatsApp do diretor não configurado. Pulando envio por WhatsApp."
      );
      return;
    }

    console.log("Inicializando cliente WhatsApp...");
    const client = await inicializarWhatsApp();

    // Aguardar um pouco mais para garantir que o cliente está totalmente pronto
    await delay(3000);

    // Verificar se o cliente está realmente pronto
    if (!client.info) {
      console.warn(
        "Cliente WhatsApp pode não estar totalmente pronto. Aguardando mais..."
      );
      await delay(5000);
    }

    // Formatar número (remover caracteres não numéricos e adicionar código do país se necessário)
    let numeroFormatado = numeroDestinatario.replace(/\D/g, "");
    if (!numeroFormatado.startsWith("55")) {
      numeroFormatado = "55" + numeroFormatado;
    }
    numeroFormatado = numeroFormatado + "@c.us";

    console.log(`Enviando mensagem para ${numeroFormatado}...`);

    const mensagem = `
📊 *Relatório Diário - Centro de Operações*

📅 *Data de Referência:* ${dataAnterior}

📈 *Resumo do Dia Anterior:*
• Transmissões: *${transmissoes.length}*
• Apólices Emitidas: *${apolicesEmitidas.length}*
• Sinistros Abertos: *${sinistros.length}*
• Assistências Urgentes: *${assistenciasUrgentes.length}*

📎 Planilha completa em anexo.
    `.trim();

    // Função para enviar mensagem com retry
    async function enviarComRetry(fn, tentativas = 3) {
      for (let i = 0; i < tentativas; i++) {
        try {
          return await fn();
        } catch (err) {
          console.warn(`Tentativa ${i + 1} falhou:`, err.message);
          if (i < tentativas - 1) {
            await delay(2000 * (i + 1)); // Delay progressivo
          } else {
            throw err;
          }
        }
      }
    }

    // Enviar mensagem com retry
    await enviarComRetry(async () => {
      await client.sendMessage(numeroFormatado, mensagem);
    });

    console.log("Mensagem de texto enviada com sucesso!");

    // Aguardar um pouco antes de enviar o arquivo
    await delay(2000);

    // Verificar se o arquivo existe antes de enviar
    if (!fs.existsSync(arquivoExcel.filePath)) {
      throw new Error(`Arquivo Excel não encontrado: ${arquivoExcel.filePath}`);
    }

    // Obter informações do arquivo
    const fileStats = fs.statSync(arquivoExcel.filePath);
    const fileSizeInMB = (fileStats.size / (1024 * 1024)).toFixed(2);
    console.log(
      `Preparando arquivo para envio: ${arquivoExcel.fileName} (${fileSizeInMB} MB)`
    );

    // Verificar se o arquivo não é muito grande (limite do WhatsApp é ~64MB, mas recomendamos menos)
    if (fileStats.size > 64 * 1024 * 1024) {
      console.warn(
        `Aviso: Arquivo muito grande (${fileSizeInMB} MB). Pode falhar no envio.`
      );
    }

    // Ler arquivo e criar MessageMedia
    const fileBuffer = fs.readFileSync(arquivoExcel.filePath);
    const base64File = fileBuffer.toString("base64");

    console.log(
      `Arquivo convertido para base64 (${(base64File.length / 1024).toFixed(
        2
      )} KB)`
    );

    const media = new MessageMedia(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      base64File,
      arquivoExcel.fileName
    );

    console.log("MessageMedia criado com sucesso. Enviando arquivo...");

    // Enviar arquivo Excel com retry
    let arquivoEnviado = false;
    await enviarComRetry(async () => {
      const resultado = await client.sendMessage(numeroFormatado, media, {
        caption: arquivoExcel.fileName,
      });
      arquivoEnviado = true;
      return resultado;
    });

    if (arquivoEnviado) {
      console.log("Arquivo Excel enviado com sucesso!");
      // Aguardar um pouco mais para garantir que o arquivo foi processado pelo WhatsApp
      await delay(3000);
    } else {
      console.warn(
        "Aviso: Não foi possível confirmar o envio do arquivo Excel"
      );
    }

    console.log("WhatsApp enviado com sucesso!");
  } catch (err) {
    console.error("Erro ao enviar WhatsApp:", err);
    console.error("Stack trace:", err.stack);
    
    // Notificar erro por e-mail
    await enviarEmailErro(
      "Erro ao Enviar WhatsApp - Relatório Diário",
      "Ocorreu um erro ao tentar enviar o relatório diário via WhatsApp. O relatório foi gerado e enviado por e-mail, mas falhou no envio via WhatsApp.",
      err,
      {
        data: dataAnterior,
        tipo: "Relatório Diário"
      }
    );
    
    // Não lançar erro para não interromper o processo
  } finally {
    // Aguardar um tempo adicional antes de desconectar para garantir que os envios foram concluídos
    console.log("Aguardando processamento final antes de desconectar...");
    await delay(5000);

    // Desconectar o cliente WhatsApp após o envio para permitir que o processo termine
    if (whatsappClient) {
      try {
        await whatsappClient.destroy();
        whatsappClient = null;
        console.log("Cliente WhatsApp desconectado.");
      } catch (destroyErr) {
        console.warn(
          "Erro ao desconectar cliente WhatsApp:",
          destroyErr.message
        );
      }
    }
  }
}

/**
 * Função principal que gera e envia o relatório
 * @param {boolean} encerrarProcesso - Se true, encerra o processo após completar (padrão: true)
 */
async function gerarRelatorioDiario(encerrarProcesso = true) {
  try {
    console.log("Iniciando geração do relatório diário...");

    // Obter data do dia anterior
    const dataAnterior = formatDate(-1);
    console.log(`Buscando dados do dia: ${dataAnterior}`);

    // Autenticar
    const authData = await getAuth();
    const token = authData.data.token;
    console.log("Autenticação realizada com sucesso");

    // Buscar dados
    console.log("Buscando transmissões...");
    const transmissoes = await getTransmissoesDiaAnterior(token, dataAnterior);
    console.log(`Transmissões encontradas: ${transmissoes.length}`);

    console.log("Buscando apólices emitidas...");
    const apolicesEmitidas = await getApolicesEmitidasDiaAnterior(
      token,
      dataAnterior
    );
    console.log(`Apólices emitidas encontradas: ${apolicesEmitidas.length}`);

    console.log("Buscando sinistros...");
    const sinistros = await getSinistrosAbertosDiaAnterior(token, dataAnterior);
    console.log(`Sinistros encontrados: ${sinistros.length}`);

    console.log("Buscando assistências urgentes...");
    const assistenciasUrgentes = await getAssistenciasUrgentesDiaAnterior(
      dataAnterior
    );
    console.log(
      `Assistências urgentes encontradas: ${assistenciasUrgentes.length}`
    );

    // Buscar produção completa do dia anterior para a planilha
    console.log("Buscando produção completa do dia anterior...");
    const producaoCompleta = await getProducaoCompletaDiaAnterior(
      token,
      dataAnterior
    );
    console.log(
      `Produção completa encontrada: ${producaoCompleta.length} registros`
    );

    // Gerar planilha Excel completa com 3 abas
    console.log("Gerando planilha Excel completa (Produção, Sinistros, Assistências)...");
    const arquivoExcel = await gerarPlanilhaExcelCompleta(
      producaoCompleta,
      sinistros,
      assistenciasUrgentes,
      dataAnterior
    );
    console.log(`Planilha gerada: ${arquivoExcel.fileName}`);

    // Enviar por e-mail
    console.log("Enviando por e-mail...");
    try {
      await enviarEmail(
        transmissoes,
        apolicesEmitidas,
        sinistros,
        assistenciasUrgentes,
        arquivoExcel,
        dataAnterior
      );
    } catch (emailErr) {
      console.error("Erro ao enviar e-mail:", emailErr);
      await enviarEmailErro(
        "Erro ao Enviar E-mail - Relatório Diário",
        "Ocorreu um erro ao tentar enviar o relatório diário por e-mail. O relatório foi gerado, mas não foi enviado.",
        emailErr,
        {
          data: dataAnterior,
          tipo: "Relatório Diário"
        }
      );
      // Continua o processo mesmo se o e-mail falhar
    }

    // Enviar por WhatsApp
    console.log("Enviando por WhatsApp...");
    await enviarWhatsApp(
      transmissoes,
      apolicesEmitidas,
      sinistros,
      assistenciasUrgentes,
      arquivoExcel,
      dataAnterior
    );

    // Limpar arquivo temporário
    try {
      fs.unlinkSync(arquivoExcel.filePath);
      console.log("Arquivo temporário removido");
    } catch (err) {
      console.warn("Erro ao remover arquivo temporário:", err);
    }

    console.log("Relatório enviado com sucesso!");
  } catch (err) {
    console.error("Erro ao gerar relatório:", err);
    console.error("Stack trace:", err.stack);

    // Enviar e-mail de erro usando função centralizada
    await enviarEmailErro(
      "Erro ao Gerar Relatório Diário",
      "Ocorreu um erro crítico ao tentar gerar ou enviar o relatório diário. O processo foi interrompido.",
      err,
      {
        tipo: "Relatório Diário",
        data: dataAnterior
      }
    );

    throw err;
  } finally {
    // Garantir que o processo termine após a execução (apenas se solicitado)
    if (encerrarProcesso) {
      console.log("Finalizando processo...");
      setTimeout(() => {
        process.exit(0);
      }, 2000); // Aguarda 2 segundos para logs finais
    }
  }
}

/**
 * Função que gera e envia o relatório semanal
 * @param {boolean} encerrarProcesso - Se true, encerra o processo após completar (padrão: true)
 */
async function gerarRelatorioSemanal(encerrarProcesso = true) {
  try {
    console.log("Iniciando geração do relatório semanal...");

    // Obter intervalo da semana (segunda a sexta da semana anterior)
    const { inicio: dataInicial, fim: dataFinal } = getLastWeekRange();
    console.log(`Buscando dados da semana: ${dataInicial} até ${dataFinal}`);

    // Autenticar
    const authData = await getAuth();
    const token = authData.data.token;
    console.log("Autenticação realizada com sucesso");

    // Buscar dados acumulados da semana
    const datesInWeek = getDateRange(dataInicial, dataFinal);
    
    console.log("Buscando dados da semana...");
    let transmissoesTotais = [];
    let apolicesEmitidasTotais = [];
    let sinistrosTotais = [];
    let assistenciasUrgentesTotais = [];

    // Buscar dados dia por dia
    for (const data of datesInWeek) {
      console.log(`Processando dia: ${data}`);
      
      const transmissoes = await getTransmissoesDiaAnterior(token, data);
      transmissoesTotais.push(...transmissoes);
      
      const apolices = await getApolicesEmitidasDiaAnterior(token, data);
      apolicesEmitidasTotais.push(...apolices);
      
      const sinistros = await getSinistrosAbertosDiaAnterior(token, data);
      sinistrosTotais.push(...sinistros);
      
      const assistencias = await getAssistenciasUrgentesDiaAnterior(data);
      assistenciasUrgentesTotais.push(...assistencias);
    }

    console.log(`Transmissões da semana: ${transmissoesTotais.length}`);
    console.log(`Apólices emitidas da semana: ${apolicesEmitidasTotais.length}`);
    console.log(`Sinistros da semana: ${sinistrosTotais.length}`);
    console.log(`Assistências urgentes da semana: ${assistenciasUrgentesTotais.length}`);

    // Buscar produção completa da semana para a planilha
    console.log("Buscando produção completa da semana...");
    const producaoCompleta = await getProducaoCompletaPeriodo(
      token,
      dataInicial,
      dataFinal
    );
    console.log(
      `Produção completa encontrada: ${producaoCompleta.length} registros`
    );

    // Gerar planilha Excel completa com 3 abas
    console.log("Gerando planilha Excel completa (Produção, Sinistros, Assistências)...");
    const periodoSemanal = `semana_${dataInicial}_a_${dataFinal}`;
    const arquivoExcel = await gerarPlanilhaExcelCompleta(
      producaoCompleta,
      sinistrosTotais,
      assistenciasUrgentesTotais,
      periodoSemanal
    );
    console.log(`Planilha gerada: ${arquivoExcel.fileName}`);

    // Enviar por e-mail (adaptando a função existente)
    console.log("Enviando relatório semanal por e-mail...");
    try {
      await enviarEmailSemanal(
        transmissoesTotais,
        apolicesEmitidasTotais,
        sinistrosTotais,
        assistenciasUrgentesTotais,
        arquivoExcel,
        dataInicial,
        dataFinal
      );
    } catch (emailErr) {
      console.error("Erro ao enviar e-mail semanal:", emailErr);
      await enviarEmailErro(
        "Erro ao Enviar E-mail - Relatório Semanal",
        "Ocorreu um erro ao tentar enviar o relatório semanal por e-mail. O relatório foi gerado, mas não foi enviado.",
        emailErr,
        {
          periodo: `${dataInicial} até ${dataFinal}`,
          tipo: "Relatório Semanal"
        }
      );
      // Continua o processo mesmo se o e-mail falhar
    }

    // Enviar por WhatsApp
    console.log("Enviando relatório semanal por WhatsApp...");
    await enviarWhatsAppSemanal(
      transmissoesTotais,
      apolicesEmitidasTotais,
      sinistrosTotais,
      assistenciasUrgentesTotais,
      arquivoExcel,
      dataInicial,
      dataFinal
    );

    // Limpar arquivo temporário
    try {
      fs.unlinkSync(arquivoExcel.filePath);
      console.log("Arquivo temporário removido");
    } catch (err) {
      console.warn("Erro ao remover arquivo temporário:", err);
    }

    console.log("Relatório semanal enviado com sucesso!");
  } catch (err) {
    console.error("Erro ao gerar relatório semanal:", err);
    console.error("Stack trace:", err.stack);

    // Enviar e-mail de erro usando função centralizada
    await enviarEmailErro(
      "Erro ao Gerar Relatório Semanal",
      "Ocorreu um erro crítico ao tentar gerar ou enviar o relatório semanal. O processo foi interrompido.",
      err,
      {
        tipo: "Relatório Semanal",
        periodo: `${dataInicial} até ${dataFinal}`
      }
    );

    throw err;
  } finally {
    // Garantir que o processo termine após a execução (apenas se solicitado)
    if (encerrarProcesso) {
      console.log("Finalizando processo...");
      setTimeout(() => {
        process.exit(0);
      }, 2000);
    }
  }
}

/**
 * Função que gera e envia o relatório mensal
 * @param {boolean} encerrarProcesso - Se true, encerra o processo após completar (padrão: true)
 */
async function gerarRelatorioMensal(encerrarProcesso = true) {
  try {
    console.log("Iniciando geração do relatório mensal...");

    // Obter intervalo do mês anterior
    const { inicio, fim, mesNome } = getLastMonthRange();
    console.log(`Buscando dados do mês: ${mesNome} (${inicio} até ${fim})`);

    // Autenticar
    const authData = await getAuth();
    const token = authData.data.token;
    console.log("Autenticação realizada com sucesso");

    // Buscar dados acumulados do mês
    const datesInMonth = getDateRange(inicio, fim);
    
    console.log("Buscando dados do mês...");
    let transmissoesTotais = [];
    let apolicesEmitidasTotais = [];
    let sinistrosTotais = [];
    let assistenciasUrgentesTotais = [];

    // Buscar dados dia por dia (pode levar tempo)
    for (const data of datesInMonth) {
      console.log(`Processando dia: ${data}`);
      
      const transmissoes = await getTransmissoesDiaAnterior(token, data);
      transmissoesTotais.push(...transmissoes);
      
      const apolices = await getApolicesEmitidasDiaAnterior(token, data);
      apolicesEmitidasTotais.push(...apolices);
      
      const sinistros = await getSinistrosAbertosDiaAnterior(token, data);
      sinistrosTotais.push(...sinistros);
      
      const assistencias = await getAssistenciasUrgentesDiaAnterior(data);
      assistenciasUrgentesTotais.push(...assistencias);
    }

    console.log(`Transmissões do mês: ${transmissoesTotais.length}`);
    console.log(`Apólices emitidas do mês: ${apolicesEmitidasTotais.length}`);
    console.log(`Sinistros do mês: ${sinistrosTotais.length}`);
    console.log(`Assistências urgentes do mês: ${assistenciasUrgentesTotais.length}`);

    // Buscar produção completa do mês para a planilha
    console.log("Buscando produção completa do mês...");
    const producaoCompleta = await getProducaoCompletaPeriodo(
      token,
      inicio,
      fim
    );
    console.log(
      `Produção completa encontrada: ${producaoCompleta.length} registros`
    );

    // Gerar planilha Excel completa com 3 abas
    console.log("Gerando planilha Excel completa (Produção, Sinistros, Assistências)...");
    const arquivoExcel = await gerarPlanilhaExcelCompleta(
      producaoCompleta,
      sinistrosTotais,
      assistenciasUrgentesTotais,
      mesNome
    );
    console.log(`Planilha gerada: ${arquivoExcel.fileName}`);

    // Enviar por e-mail
    console.log("Enviando relatório mensal por e-mail...");
    try {
      await enviarEmailMensal(
        transmissoesTotais,
        apolicesEmitidasTotais,
        sinistrosTotais,
        assistenciasUrgentesTotais,
        arquivoExcel,
        mesNome
      );
    } catch (emailErr) {
      console.error("Erro ao enviar e-mail mensal:", emailErr);
      await enviarEmailErro(
        "Erro ao Enviar E-mail - Relatório Mensal",
        "Ocorreu um erro ao tentar enviar o relatório mensal por e-mail. O relatório foi gerado, mas não foi enviado.",
        emailErr,
        {
          periodo: mesNome,
          tipo: "Relatório Mensal"
        }
      );
      // Continua o processo mesmo se o e-mail falhar
    }

    // Enviar por WhatsApp
    console.log("Enviando relatório mensal por WhatsApp...");
    await enviarWhatsAppMensal(
      transmissoesTotais,
      apolicesEmitidasTotais,
      sinistrosTotais,
      assistenciasUrgentesTotais,
      arquivoExcel,
      mesNome
    );

    // Limpar arquivo temporário
    try {
      fs.unlinkSync(arquivoExcel.filePath);
      console.log("Arquivo temporário removido");
    } catch (err) {
      console.warn("Erro ao remover arquivo temporário:", err);
    }

    console.log("Relatório mensal enviado com sucesso!");
  } catch (err) {
    console.error("Erro ao gerar relatório mensal:", err);
    console.error("Stack trace:", err.stack);

    // Enviar e-mail de erro usando função centralizada
    await enviarEmailErro(
      "Erro ao Gerar Relatório Mensal",
      "Ocorreu um erro crítico ao tentar gerar ou enviar o relatório mensal. O processo foi interrompido.",
      err,
      {
        tipo: "Relatório Mensal",
        periodo: mesNome
      }
    );

    throw err;
  } finally {
    // Garantir que o processo termine após a execução (apenas se solicitado)
    if (encerrarProcesso) {
      console.log("Finalizando processo...");
      setTimeout(() => {
        process.exit(0);
      }, 2000);
    }
  }
}

// Funções auxiliares de envio semanal e mensal
async function enviarEmailSemanal(transmissoes, apolices, sinistros, assistencias, arquivo, dataInicio, dataFim) {
  // Implementação similar ao enviarEmail, mas com texto adaptado para semanal
  const email = process.env.MAIL_EMAIL;
  const password = process.env.MAIL_PASSWORD;
  const emailDestinatario = process.env.DIRETOR_EMAIL || email;

  const transporter = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 587,
    secure: false,
    auth: { user: email, pass: password },
    tls: { rejectUnauthorized: false },
  });

  await transporter.verify();
  
  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>Relatório Semanal</title></head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; background-color: #f5f5f5;">
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
    <tr><td style="padding: 30px 15px;">
      <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
        <tr><td style="border-bottom: 3px solid #4A04A5; flex: 1; justify-items: center">
          <img src="https://iili.io/KZqR9cu.png" alt="Avantar" style="width: 200px; height: 200px; display: block;" />
        </td></tr>
        <tr><td style="padding: 35px 40px 25px 40px;">
          <h1 style="margin: 0; font-size: 24px; font-weight: 400; color: #170138;">Relatório Semanal</h1>
          <p style="margin: 8px 0 0 0; font-size: 14px; color: #666666;">Centro de Operações</p>
        </td></tr>
        <tr><td style="padding: 0 40px 30px 40px;">
          <p style="margin: 0; font-size: 13px; color: #666666;">Período: ${dataInicio} até ${dataFim}</p>
        </td></tr>
        <tr><td style="padding: 0 40px;">
          <h2 style="margin: 0 0 20px 0; font-size: 16px; font-weight: 500; color: #170138; text-transform: uppercase;">Resumo da Semana</h2>
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666; border-bottom: 1px solid #e0e0e0;">Transmissões</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; border-bottom: 1px solid #e0e0e0;">${transmissoes.length}</td></tr>
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666; border-bottom: 1px solid #e0e0e0;">Apólices Emitidas</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; border-bottom: 1px solid #e0e0e0;">${apolices.length}</td></tr>
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666; border-bottom: 1px solid #e0e0e0;">Sinistros Abertos</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; border-bottom: 1px solid #e0e0e0;">${sinistros.length}</td></tr>
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666;">Assistências Urgentes</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right;">${assistencias.length}</td></tr>
          </table>
        </td></tr>
        <tr><td style="padding: 35px 40px 40px 40px;">
          <p style="margin: 0; font-size: 13px; color: #666666;">Em anexo, planilha completa da semana.</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

  await transporter.sendMail({
    from: `"Tecnologia Avantar" <${email}>`,
    to: emailDestinatario,
    subject: `Relatório Semanal - Centro de Operações - ${dataInicio} a ${dataFim}`,
    html,
    attachments: [{ filename: arquivo.fileName, path: arquivo.filePath }],
  });

  console.log("E-mail semanal enviado com sucesso!");
}

async function enviarEmailMensal(transmissoes, apolices, sinistros, assistencias, arquivo, mesNome) {
  const email = process.env.MAIL_EMAIL;
  const password = process.env.MAIL_PASSWORD;
  const emailDestinatario = process.env.DIRETOR_EMAIL || email;

  const transporter = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 587,
    secure: false,
    auth: { user: email, pass: password },
    tls: { rejectUnauthorized: false },
  });

  await transporter.verify();
  
  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>Relatório Mensal</title></head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; background-color: #f5f5f5;">
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
    <tr><td style="padding: 30px 15px;">
      <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
        <tr><td style="border-bottom: 3px solid #4A04A5; flex: 1; justify-items: center">
          <img src="https://iili.io/KZqR9cu.png" alt="Avantar" style="width: 200px; height: 200px; display: block;" />
        </td></tr>
        <tr><td style="padding: 35px 40px 25px 40px;">
          <h1 style="margin: 0; font-size: 24px; font-weight: 400; color: #170138;">Relatório Mensal</h1>
          <p style="margin: 8px 0 0 0; font-size: 14px; color: #666666;">Centro de Operações</p>
        </td></tr>
        <tr><td style="padding: 0 40px 30px 40px;">
          <p style="margin: 0; font-size: 13px; color: #666666;">Período: ${mesNome}</p>
        </td></tr>
        <tr><td style="padding: 0 40px;">
          <h2 style="margin: 0 0 20px 0; font-size: 16px; font-weight: 500; color: #170138; text-transform: uppercase;">Resumo do Mês</h2>
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666; border-bottom: 1px solid #e0e0e0;">Transmissões</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; border-bottom: 1px solid #e0e0e0;">${transmissoes.length}</td></tr>
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666; border-bottom: 1px solid #e0e0e0;">Apólices Emitidas</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; border-bottom: 1px solid #e0e0e0;">${apolices.length}</td></tr>
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666; border-bottom: 1px solid #e0e0e0;">Sinistros Abertos</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right; border-bottom: 1px solid #e0e0e0;">${sinistros.length}</td></tr>
            <tr><td style="padding: 18px 0; font-size: 14px; color: #666666;">Assistências Urgentes</td>
                <td style="padding: 18px 0; font-size: 28px; color: #4A04A5; text-align: right;">${assistencias.length}</td></tr>
          </table>
        </td></tr>
        <tr><td style="padding: 35px 40px 40px 40px;">
          <p style="margin: 0; font-size: 13px; color: #666666;">Em anexo, planilha completa do mês.</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

  await transporter.sendMail({
    from: `"Tecnologia Avantar" <${email}>`,
    to: emailDestinatario,
    subject: `Relatório Mensal - Centro de Operações - ${mesNome}`,
    html,
    attachments: [{ filename: arquivo.fileName, path: arquivo.filePath }],
  });

  console.log("E-mail mensal enviado com sucesso!");
}

async function enviarWhatsAppSemanal(transmissoes, apolices, sinistros, assistencias, arquivo, dataInicio, dataFim) {
  try {
    const numeroDestinatario = process.env.DIRETOR_WHATSAPP;
    if (!numeroDestinatario) return;

    const client = await inicializarWhatsApp();
    await delay(3000);

    let numeroFormatado = numeroDestinatario.replace(/\D/g, "");
    if (!numeroFormatado.startsWith("55")) numeroFormatado = "55" + numeroFormatado;
    numeroFormatado = numeroFormatado + "@c.us";

    const mensagem = `📊 *Relatório Semanal - Centro de Operações*

📅 *Período:* ${dataInicio} até ${dataFim}

📈 *Resumo da Semana:*
• Transmissões: *${transmissoes.length}*
• Apólices Emitidas: *${apolices.length}*
• Sinistros Abertos: *${sinistros.length}*
• Assistências Urgentes: *${assistencias.length}*

📎 Planilha completa em anexo.`.trim();

    await client.sendMessage(numeroFormatado, mensagem);
    console.log("Mensagem semanal enviada!");
    await delay(2000);

    const fileBuffer = fs.readFileSync(arquivo.filePath);
    const base64File = fileBuffer.toString("base64");
    const media = new MessageMedia(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      base64File,
      arquivo.fileName
    );

    await client.sendMessage(numeroFormatado, media);
    console.log("Arquivo semanal enviado!");
    await delay(3000);
  } catch (err) {
    console.error("Erro ao enviar WhatsApp semanal:", err);
    console.error("Stack trace:", err.stack);
    
    // Notificar erro por e-mail
    await enviarEmailErro(
      "Erro ao Enviar WhatsApp - Relatório Semanal",
      "Ocorreu um erro ao tentar enviar o relatório semanal via WhatsApp. O relatório foi gerado e enviado por e-mail, mas falhou no envio via WhatsApp.",
      err,
      {
        periodo: `${dataInicio} até ${dataFim}`,
        tipo: "Relatório Semanal"
      }
    );
  } finally {
    if (whatsappClient) {
      try {
        await delay(5000);
        await whatsappClient.destroy();
        whatsappClient = null;
      } catch (destroyErr) {
        console.warn("Erro ao desconectar cliente WhatsApp:", destroyErr.message);
      }
    }
  }
}

async function enviarWhatsAppMensal(transmissoes, apolices, sinistros, assistencias, arquivo, mesNome) {
  try {
    const numeroDestinatario = process.env.DIRETOR_WHATSAPP;
    if (!numeroDestinatario) return;

    const client = await inicializarWhatsApp();
    await delay(3000);

    let numeroFormatado = numeroDestinatario.replace(/\D/g, "");
    if (!numeroFormatado.startsWith("55")) numeroFormatado = "55" + numeroFormatado;
    numeroFormatado = numeroFormatado + "@c.us";

    const mensagem = `📊 *Relatório Mensal - Centro de Operações*

📅 *Período:* ${mesNome}

📈 *Resumo do Mês:*
• Transmissões: *${transmissoes.length}*
• Apólices Emitidas: *${apolices.length}*
• Sinistros Abertos: *${sinistros.length}*
• Assistências Urgentes: *${assistencias.length}*

📎 Planilha completa em anexo.`.trim();

    await client.sendMessage(numeroFormatado, mensagem);
    console.log("Mensagem mensal enviada!");
    await delay(2000);

    const fileBuffer = fs.readFileSync(arquivo.filePath);
    const base64File = fileBuffer.toString("base64");
    const media = new MessageMedia(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      base64File,
      arquivo.fileName
    );

    await client.sendMessage(numeroFormatado, media);
    console.log("Arquivo mensal enviado!");
    await delay(3000);
  } catch (err) {
    console.error("Erro ao enviar WhatsApp mensal:", err);
    console.error("Stack trace:", err.stack);
    
    // Notificar erro por e-mail
    await enviarEmailErro(
      "Erro ao Enviar WhatsApp - Relatório Mensal",
      "Ocorreu um erro ao tentar enviar o relatório mensal via WhatsApp. O relatório foi gerado e enviado por e-mail, mas falhou no envio via WhatsApp.",
      err,
      {
        periodo: mesNome,
        tipo: "Relatório Mensal"
      }
    );
  } finally {
    if (whatsappClient) {
      try {
        await delay(5000);
        await whatsappClient.destroy();
        whatsappClient = null;
      } catch (destroyErr) {
        console.warn("Erro ao desconectar cliente WhatsApp:", destroyErr.message);
      }
    }
  }
}


// ===== HANDLERS GLOBAIS DE ERRO =====
// Captura erros não tratados
process.on("uncaughtException", async (err) => {
  console.error("❌ ERRO NÃO TRATADO (uncaughtException):", err);
  console.error("Stack trace:", err.stack);
  
  // Tentar notificar por e-mail
  try {
    await enviarEmailErro(
      "Erro Crítico Não Tratado - Sistema de Relatórios",
      "Ocorreu um erro crítico não tratado que pode ter interrompido o sistema de relatórios.",
      err,
      {
        tipo: "uncaughtException",
        timestamp: new Date().toISOString()
      }
    );
  } catch (emailErr) {
    console.error("❌ Falha ao enviar e-mail de erro crítico:", emailErr);
  }
  
  // Dar tempo para o e-mail ser enviado antes de encerrar
  setTimeout(() => {
    process.exit(1);
  }, 5000);
});

// Captura promises rejeitadas não tratadas
process.on("unhandledRejection", async (reason, promise) => {
  console.error("❌ PROMISE REJEITADA NÃO TRATADA:", reason);
  console.error("Promise:", promise);
  
  // Tentar notificar por e-mail
  try {
    const erro = reason instanceof Error ? reason : new Error(String(reason));
    await enviarEmailErro(
      "Promise Rejeitada Não Tratada - Sistema de Relatórios",
      "Uma promise foi rejeitada e não foi tratada adequadamente. Isso pode indicar um problema no código assíncrono.",
      erro,
      {
        tipo: "unhandledRejection",
        timestamp: new Date().toISOString()
      }
    );
  } catch (emailErr) {
    console.error("❌ Falha ao enviar e-mail de erro crítico:", emailErr);
  }
});

// Agendar execução diária às 6h (terça a sábado - dias úteis)
console.log("Agendando relatório diário para 6h da manhã (terça a sábado)...");
cron.schedule("0 6 * * 2-6", async () => {
  try {
    console.log("Executando relatório diário agendado...");
    await gerarRelatorioDiario();
  } catch (err) {
    console.error("Erro não tratado no relatório diário agendado:", err);
    // O erro já foi notificado dentro da função gerarRelatorioDiario
  }
}, {
  timezone: "America/Sao_Paulo"
});

// Agendar execução semanal aos sábados às 6h15
console.log("Agendando relatório semanal para sábados às 6h15...");
cron.schedule("15 6 * * 6", async () => {
  try {
    console.log("Executando relatório semanal agendado...");
    await gerarRelatorioSemanal();
  } catch (err) {
    console.error("Erro não tratado no relatório semanal agendado:", err);
    // O erro já foi notificado dentro da função gerarRelatorioSemanal
  }
}, {
  timezone: "America/Sao_Paulo"
}); 

// Agendar execução mensal no primeiro dia do mês às 6h00
console.log("Agendando relatório mensal para o primeiro dia de cada mês às 6h00...");
cron.schedule("0 6 1 * *", async () => {
  try {
    console.log("Executando relatório mensal agendado...");
    await gerarRelatorioMensal();
  } catch (err) {
    console.error("Erro não tratado no relatório mensal agendado:", err);
    // O erro já foi notificado dentro da função gerarRelatorioMensal
  }
}, {
  timezone: "America/Sao_Paulo"
});

console.log("\n📊 Sistema de Relatórios - Centro de Operações");
console.log("===============================================");
console.log("✅ Relatório Diário: Terça a Sábado às 6h00");
console.log("✅ Relatório Semanal: Sábados às 6h15");
console.log("✅ Relatório Mensal: Dia 1 de cada mês às 6h00");
console.log("===============================================\n");

// Executar imediatamente se necessário (para testes)

/*
// Descomente as linhas abaixo para testar
console.log("Executando relatórios imediatamente (modo teste)...");

// Executar todos em sequência (sem encerrar o processo entre eles)
(async () => {
  try {
    await gerarRelatorioDiario(false); // false = não encerra o processo
    await gerarRelatorioSemanal(false); // false = não encerra o processo
    await gerarRelatorioMensal(false); // false = não encerra o processo
    console.log("\n✅ Todos os relatórios foram executados com sucesso!");
    console.log("Finalizando processo...");
    setTimeout(() => {
      process.exit(0);
    }, 2000);
  } catch (err) {
    console.error("Erro ao executar relatórios:", err);
    process.exit(1);
  }
})();
*/