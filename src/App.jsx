import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Search, UploadCloud, CheckCircle2, XCircle, AlertTriangle,
  Copy, Receipt, Info, Download, FileSpreadsheet, Pencil, X, RotateCcw,
  Save, History, Clock, CornerDownLeft,
} from "lucide-react";

/* ============================== Paleta ============================== */
const CORES = {
  bgPage: "#020508",
  bgCard: "#080E1A",
  bgCardAlt: "#060C17",
  bgCardHover: "#0D1525",
  borda: "#141C2E",
  bordaForte: "#1D2840",
  texto: "#EDEFF7",
  textoSub: "#94A0BD",
  textoFraco: "#5C6786",

  accent: "#FF5A36",
  accentDark: "#D9431F",
  accentSoft: "rgba(255,90,54,0.12)",
  accentBorder: "rgba(255,90,54,0.35)",

  ok: "#3DDC97",
  okSoft: "rgba(61,220,151,0.12)",
  alerta: "#FFC857",
  alertaSoft: "rgba(255,200,87,0.12)",
  erro: "#FF5D7A",
  erroSoft: "rgba(255,93,122,0.12)",
  duplicada: "#B083FF",
  duplicadaSoft: "rgba(176,131,255,0.12)",
  parcela: "#FFFFFF",
  parcelaSoft: "rgba(255,255,255,0.10)",
  devolucao: "#6FD6FF",
  devolucaoSoft: "rgba(111,214,255,0.10)",
};

const STATUS_CONFIG = {
  ok: { label: "OK", color: CORES.ok, bg: CORES.okSoft, Icon: CheckCircle2 },
  nao_encontrada: { label: "Não lançada", color: CORES.erro, bg: CORES.erroSoft, Icon: XCircle },
  divergente: { label: "Divergência", color: CORES.alerta, bg: CORES.alertaSoft, Icon: AlertTriangle },
  duplicada: { label: "Duplicada", color: CORES.duplicada, bg: CORES.duplicadaSoft, Icon: Copy },
  parcela: { label: "Parcela", color: CORES.parcela, bg: CORES.parcelaSoft, Icon: Receipt },
  devolucao: { label: "Devolução", color: CORES.devolucao, bg: CORES.devolucaoSoft, Icon: CornerDownLeft },
};

// Status efetivo de um resultado: o status manual (override do usuário), se houver,
// senão o status calculado automaticamente pela conferência.
function statusEfetivo(r) {
  return r.statusManual || r.tipo;
}

// Indica se o item deve ser tratado como parcela para fins de contagem/exibição:
// segue o override manual quando presente, senão a detecção automática (eraParcela).
function isParcelaEfetiva(r) {
  return r.statusManual ? r.statusManual === "parcela" : !!r.eraParcela;
}

/* ========================= Funções utilitárias ========================= */
function normalizarTexto(t) {
  if (!t) return "";
  return String(t).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function extrairNumeroNF(descricao) {
  if (!descricao) return null;
  const match = String(descricao).match(/NF[\s\-:]*(\d+)/i);
  return match ? String(parseInt(match[1])) : null;
}

// Palavras genéricas que não ajudam a identificar um fornecedor específico
const STOPWORDS = new Set([
  "ltda", "me", "epp", "sa", "eireli", "cia", "grupo",
  "comercio", "comercial", "industria", "industrial",
  "servicos", "servico", "importacao", "exportacao",
  "dos", "das", "do", "da", "de",
]);

function tokens(str) {
  return normalizarTexto(str)
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(t => t.length >= 3 && !STOPWORDS.has(t));
}

// Compara dois nomes de fornecedor por sobreposição de palavras significativas.
// Usado para distinguir notas de fornecedores diferentes que, por coincidência,
// têm o mesmo número de NF (numeração de nota é por fornecedor, não global).
function mesmoFornecedor(a, b) {
  const ta = tokens(a), tb = tokens(b);
  if (ta.length === 0 || tb.length === 0) return true; // dados insuficientes: não filtra
  return ta.some(t => tb.includes(t));
}

// Verifica se o valor lançado corresponde a 1/N do valor da nota (parcelamento
// em que apenas uma parcela aparece no período exportado).
function verificarParcela(valorSIEG, valorCA) {
  if (!valorCA || valorCA <= 0) return null;
  for (let n = 2; n <= 24; n++) {
    const total = valorCA * n;
    const tolerancia = Math.max(0.05, n * 0.015);
    if (Math.abs(valorSIEG - total) <= tolerancia) return n;
  }
  return null;
}

// Classifica o lançamento como CUSTO ou DESPESA a partir do texto da categoria
// do Conta Azul (que normalmente traz a tag "(Custo)" ou "(Despesa)" no final,
// ex: "Manutenção de Veículos (CUSTO)"). Quando a categoria não traz a tag
// explícita, usa o centro de custo como critério auxiliar: centros de custo
// operacionais são tratados como CUSTO.
function extrairCustoDespesa(categoria, centroCusto) {
  const catNorm = normalizarTexto(categoria);
  if (/\(\s*custo\s*\)/.test(catNorm)) return "CUSTO";
  if (/\(\s*despesa\s*\)/.test(catNorm)) return "DESPESA";
  const centroNorm = normalizarTexto(centroCusto);
  if (centroNorm.includes("operacional")) return "CUSTO";
  return "";
}

/* ========================= Leitura: NFS-e (notas de serviço) =========================
 * A planilha de NFS-e tem uma característica diferente da SIEG (NF-e): cada aba pode
 * vir de uma fonte/exportação diferente (prefeitura, sistema nacional, etc.), então o
 * layout de colunas varia de aba para aba dentro do mesmo arquivo. Por isso a leitura
 * detecta o layout pelo cabeçalho de cada aba individualmente, em vez de assumir uma
 * posição fixa de coluna como na SIEG.
 *
 * Além disso, a planilha de NFS-e já vem com colunas "Custo" e "Despesa" preenchidas
 * com valores (não uma tag de texto como na Categoria do Conta Azul) — esses valores
 * são usados como a classificação primária, mas são comparados com o que o Conta Azul
 * indica para sinalizar divergência (ver compararCustoDespesaNFSE).
 */
function lerPlanilhaNFSE(workbook) {
  const notas = [];
  for (const aba of workbook.SheetNames) {
    const ws = workbook.Sheets[aba];
    const dados = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (dados.length === 0) continue;

    // Localiza a linha de cabeçalho procurando por uma coluna que identifique o número
    // da nota (varia entre "Número NFS-e", "Nº NFS-e", "Nº NFSE", etc.)
    let headerRow = -1, idxNum = -1;
    for (let i = 0; i < Math.min(5, dados.length); i++) {
      const row = dados[i] || [];
      const idx = row.findIndex(c => {
        const h = normalizarTexto(String(c || ""));
        return /^n[ºo°]?\s*nfs[\s-]?e$|^numero nfs[\s-]?e$/.test(h.replace(/\s+/g, " ").trim());
      });
      if (idx !== -1) { headerRow = i; idxNum = idx; break; }
    }
    if (headerRow === -1) continue;
    const headers = dados[headerRow].map(h => normalizarTexto(String(h || "")));

    const findIdx = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        if (keywords.some(k => headers[i].includes(k))) return i;
      }
      return -1;
    };

    // Fornecedor = "Prestador" (quem prestou o serviço), nunca "Tomador" (a própria
    // empresa, que aparece em alguns layouts mas não serve para comparação).
    const idxRazaoPrestador = findIdx(["razao social do prestador", "nome prestador", "nome/ nome empresarial", "nome empresarial"]);
    const idxCNPJPrestador = findIdx(["cnpj/cpf prestador", "cpf/cnpj do prestador"]);
    const idxValor = findIdx(["valor do servico", "valor dos servicos", "vr. da nfse", "vr da nfse"]);
    const idxData = findIdx(["data geracao", "data hora nfe", "data de emissao", "data do fato gerador"]);
    const idxCusto = findIdx(["custo"]);
    const idxDespesa = findIdx(["despesa"]);
    const idxValidacao = findIdx(["validacao"]);

    for (let i = headerRow + 1; i < dados.length; i++) {
      const row = dados[i];
      if (!row) continue;
      const numNF = row[idxNum];
      // Pula linhas de total/resumo (ex.: "TOTAL (26 notas)" na última linha da aba),
      // que não têm número de nota mas têm valor.
      if (numNF === null || numNF === undefined || numNF === "") continue;

      const custo = idxCusto >= 0 ? parseFloat(row[idxCusto]) || 0 : 0;
      const despesa = idxDespesa >= 0 ? parseFloat(row[idxDespesa]) || 0 : 0;

      notas.push({
        aba,
        numNF: String(numNF),
        valor: idxValor >= 0 ? (parseFloat(row[idxValor]) || 0) : 0,
        dataEmissao: idxData >= 0 && row[idxData] ? String(row[idxData]) : "",
        fornecedor: idxRazaoPrestador >= 0 ? String(row[idxRazaoPrestador] || "") : "",
        cnpj: idxCNPJPrestador >= 0 ? String(row[idxCNPJPrestador] || "") : "",
        status: idxValidacao >= 0 ? String(row[idxValidacao] || "") : "",
        custoPlanilha: custo > 0 ? custo : null,
        despesaPlanilha: despesa > 0 ? despesa : null,
      });
    }
  }
  return notas;
}

// Compara a classificação Custo/Despesa que já vem na planilha de NFS-e com o que o
// Conta Azul indica (via extrairCustoDespesa). Quando os dois existem mas divergem,
// sinaliza para o usuário em vez de decidir silenciosamente qual prevalece.
function compararCustoDespesaNFSE(nf, lancamentosRelacionados) {
  const daPlanilha = nf.custoPlanilha != null && nf.despesaPlanilha != null
    ? "CUSTO + DESPESA" // nota dividida entre as duas classificações
    : nf.custoPlanilha != null ? "CUSTO"
    : nf.despesaPlanilha != null ? "DESPESA" : "";

  const doContaAzul = Array.from(new Set(
    (lancamentosRelacionados || []).map(l => l.custoDespesa).filter(Boolean)
  )).join(" / ");

  if (!daPlanilha) return { valor: doContaAzul, divergente: false };
  if (!doContaAzul) return { valor: daPlanilha, divergente: false };

  const divergente = daPlanilha !== doContaAzul && !(daPlanilha === "CUSTO + DESPESA");
  return { valor: daPlanilha, doContaAzul, divergente };
}

function lerPlanilhaSIEG(workbook) {
  const notas = [];
  const abas = workbook.SheetNames;
  for (const aba of abas) {
    const ws = workbook.Sheets[aba];
    const dados = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    let headerRow = -1;
    for (let i = 0; i < dados.length; i++) {
      if (dados[i] && dados[i][0] === "Num NFe") { headerRow = i; break; }
    }
    if (headerRow === -1) continue;
    const headers = dados[headerRow];
    const idxNumNF = headers.indexOf("Num NFe");
    const idxValor = headers.indexOf("Valor");
    const idxData = headers.indexOf("Data Emissão");
    const idxFornecedor = headers.indexOf("Nome Fant. Emit");
    const idxRazao = headers.indexOf("Razão Soc. Emit");
    const idxCNPJ = headers.indexOf("CNPJ Emit");
    const idxStatus = headers.indexOf("Status");
    for (let i = headerRow + 1; i < dados.length; i++) {
      const row = dados[i];
      if (!row || !row[idxNumNF]) continue;
      notas.push({
        aba,
        numNF: String(row[idxNumNF]),
        valor: parseFloat(row[idxValor]) || 0,
        dataEmissao: row[idxData] ? String(row[idxData]) : "",
        fornecedor: row[idxFornecedor] || row[idxRazao] || "",
        cnpj: row[idxCNPJ] || "",
        status: row[idxStatus] || "",
      });
    }
  }
  return notas;
}

function lerContaAzul(workbook) {
  const lancamentos = [];
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  const dados = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  let headerRow = 0;
  for (let i = 0; i < Math.min(10, dados.length); i++) {
    const row = dados[i];
    if (row && row.some(c => c && normalizarTexto(String(c)).includes("descri"))) {
      headerRow = i; break;
    }
  }
  const headers = dados[headerRow] || [];
  const findIdx = (keywords) => {
    for (let i = 0; i < headers.length; i++) {
      const h = normalizarTexto(String(headers[i] || ""));
      if (keywords.some(k => h.includes(k))) return i;
    }
    return -1;
  };
  const idxDesc = findIdx(["descri"]);
  const idxValor = findIdx(["valor"]);
  const idxFornecedor = findIdx(["fornecedor", "credor", "nome"]);
  const idxVenc = findIdx(["vencimento", "venc"]);
  const idxComp = findIdx(["competencia", "competência", "comp"]);
  const idxCategoria = findIdx(["categoria"]);
  const idxCentroCusto = findIdx(["centro de custo", "centro custo"]);

  for (let i = headerRow + 1; i < dados.length; i++) {
    const row = dados[i];
    if (!row) continue;
    const descricao = row[idxDesc] ? String(row[idxDesc]) : "";
    if (!descricao) continue;
    const numNF = extrairNumeroNF(descricao);
    const categoria = idxCategoria >= 0 && row[idxCategoria] ? String(row[idxCategoria]) : "";
    const centroCusto = idxCentroCusto >= 0 && row[idxCentroCusto] ? String(row[idxCentroCusto]) : "";
    lancamentos.push({
      numNF,
      descricao,
      valor: parseFloat(String(row[idxValor] || "0").replace(",", ".")) || 0,
      fornecedor: row[idxFornecedor] ? String(row[idxFornecedor]) : "",
      vencimento: row[idxVenc] ? String(row[idxVenc]) : "",
      competencia: row[idxComp] ? String(row[idxComp]) : "",
      categoria,
      centroCusto,
      custoDespesa: extrairCustoDespesa(categoria, centroCusto),
      linha: i + 1,
    });
  }
  return lancamentos;
}

/* ============================== Comparação ============================== */
function compararDados(notasSIEG, lancamentosCA) {
  const resultados = [];
  const caByNF = {};
  for (const l of lancamentosCA) {
    if (l.numNF) {
      if (!caByNF[l.numNF]) caByNF[l.numNF] = [];
      caByNF[l.numNF].push(l);
    }
  }

  for (const nf of notasSIEG) {
    const candidatos = caByNF[nf.numNF] || [];

    if (candidatos.length === 0) {
      resultados.push({ tipo: "nao_encontrada", nf, lancamento: null, detalhes: "NF não encontrada no Conta Azul" });
      continue;
    }

    // Numeração de NF é por fornecedor, não global: se houver mais de um lançamento
    // com o mesmo número, restringe pelo nome do fornecedor antes de qualquer outra
    // verificação, para não confundir notas de empresas diferentes.
    let matches = candidatos;
    if (candidatos.length > 1) {
      const doMesmoFornecedor = candidatos.filter(c => mesmoFornecedor(nf.fornecedor, c.fornecedor));
      if (doMesmoFornecedor.length > 0) matches = doMesmoFornecedor;
    }

    if (matches.length === 1) {
      const l = matches[0];
      const diffValor = Math.abs(nf.valor - l.valor);
      if (diffValor > 0.05) {
        const numParcelas = verificarParcela(nf.valor, l.valor);
        if (numParcelas) {
          resultados.push({
            tipo: "ok", nf, lancamento: l, eraParcela: true, numParcelas,
            detalhes: `Valor lançado equivale a 1/${numParcelas} do valor da NF — parcelamento confirmado (demais parcelas podem estar fora do período exportado)`,
          });
        } else {
          resultados.push({
            tipo: "divergente", nf, lancamento: l,
            detalhes: `Valor: nota R$${nf.valor.toFixed(2)} × CA R$${l.valor.toFixed(2)}`,
          });
        }
      } else {
        resultados.push({ tipo: "ok", nf, lancamento: l, detalhes: "OK" });
      }
      continue;
    }

    // Mais de um lançamento do mesmo fornecedor para a mesma NF: antes de marcar
    // como duplicidade real, verifica se a soma dos lançamentos corresponde ao
    // valor total da nota — isso indica parcelamento ou divisão do lançamento
    // (ex: parcela inicial + provisionada, ou valor dividido entre lançamentos),
    // não um erro de duplicação. Confirmado, é tratado como OK.
    const soma = matches.reduce((acc, m) => acc + m.valor, 0);
    const diffSoma = Math.abs(soma - nf.valor);
    const tolerancia = Math.max(0.05, matches.length * 0.02);

    if (diffSoma <= tolerancia) {
      const partes = matches.map(m => `R$${m.valor.toFixed(2)}`).join(" + ");
      resultados.push({
        tipo: "ok", nf, lancamento: matches[0], lancamentosRelacionados: matches,
        eraParcela: true, numParcelas: matches.length,
        detalhes: `Nota dividida em ${matches.length} lançamentos (${partes} = R$${soma.toFixed(2)}) — soma confere com o valor da NF`,
      });
    } else {
      resultados.push({
        tipo: "duplicada", nf, lancamento: matches[0], lancamentosRelacionados: matches,
        detalhes: `NF lançada ${matches.length}x no Conta Azul (soma R$${soma.toFixed(2)} ≠ valor da NF R$${nf.valor.toFixed(2)})`,
      });
    }
  }

  // Enriquece os resultados com a comparação de Custo/Despesa quando a nota vem da
  // planilha de NFS-e (que já traz essa classificação pronta) — sem alterar o tipo/
  // status do resultado, apenas anexando a informação para exibição e para a aba de
  // divergências do Excel.
  for (const r of resultados) {
    if (r.nf.custoPlanilha == null && r.nf.despesaPlanilha == null) continue;
    const relacionados = r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
    const comp = compararCustoDespesaNFSE(r.nf, relacionados);
    r.custoDespesaPlanilha = comp.valor;
    r.custoDespesaDivergente = comp.divergente;
    if (comp.divergente) {
      r.detalhes += ` | Custo/Despesa divergente: planilha indica "${comp.valor}", Conta Azul indica "${comp.doContaAzul}"`;
    }
  }

  return resultados.map((r, i) => ({ ...r, id: i }));
}
function gerarExcel(resultados, tipoPlanilha) {
  const wb = XLSX.utils.book_new();
  const TIPO_LABEL = {
    nao_encontrada: "NF Não Lançada",
    duplicada: "Duplicidade",
    divergente: "Divergência",
    ok: "OK",
    parcela: "Parcela",
    devolucao: "Devolução",
  };
  const rotuloFonte = tipoPlanilha === "nfse" ? "NFS-e" : "SIEG";

  const relacionadosDe = (r) => r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
  const getValoresCA = (r) => relacionadosDe(r).map(l => `R$${l.valor.toFixed(2)}`).join(" + ");
  const getFornecedoresCA = (r) => Array.from(new Set(relacionadosDe(r).map(l => l.fornecedor).filter(Boolean))).join(" / ");
  const getCustoDespesa = (r) => r.custoDespesaManual || r.custoDespesaPlanilha || Array.from(new Set(relacionadosDe(r).map(l => l.custoDespesa).filter(Boolean))).join(" / ");

  // Aba "Divergências": apenas problemas reais (parcelas confirmadas não entram aqui),
  // considerando o status efetivo (override manual tem prioridade sobre o automático).
  const linhasDiverg = [
    ["Tipo", "Filial", "Num NF", `Fornecedor ${rotuloFonte}`, `Valor ${rotuloFonte}`, "Data Emissão", "Valor(es) lançado(s) CA", "Fornecedor(es) CA", "Custo/Despesa", "Custo/Despesa divergente?", "Detalhes", "Observação"]
  ];
  for (const r of resultados.filter(r => statusEfetivo(r) !== "ok")) {
    linhasDiverg.push([
      TIPO_LABEL[statusEfetivo(r)] || statusEfetivo(r),
      r.nf.aba, r.nf.numNF, r.nf.fornecedor, r.nf.valor, r.nf.dataEmissao,
      getValoresCA(r), getFornecedoresCA(r), getCustoDespesa(r),
      r.custoDespesaDivergente ? "Sim" : "", r.detalhes, r.observacao || "",
    ]);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(linhasDiverg), "Divergências");

  // Aba "Parcelas": notas pagas em parcelas ou divididas em mais de um lançamento,
  // confirmadas matematicamente (ou marcadas manualmente como parcela) — mantidas
  // aqui apenas para documentação/auditoria.
  const linhasParcelas = [
    [`Filial`, "Num NF", `Fornecedor ${rotuloFonte}`, `Valor da Nota (${rotuloFonte})`, "Lançamento(s) no Conta Azul", "Custo/Despesa", "Observação", "Detalhes"]
  ];
  for (const r of resultados.filter(isParcelaEfetiva)) {
    linhasParcelas.push([
      r.nf.aba, r.nf.numNF, r.nf.fornecedor, r.nf.valor,
      getValoresCA(r), getCustoDespesa(r), r.observacao || "", r.detalhes,
    ]);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(linhasParcelas), "Parcelas");

  const linhasResumo = [
    ["Resumo da Conferência"], [],
    [`Fonte das notas`, rotuloFonte],
    [`Total de NFs na ${rotuloFonte}`, resultados.length],
    ["NFs OK", resultados.filter(r => statusEfetivo(r) === "ok").length],
    ["  das quais parceladas/divididas", resultados.filter(isParcelaEfetiva).length],
    ["NFs não lançadas", resultados.filter(r => statusEfetivo(r) === "nao_encontrada").length],
    ["NFs com valor divergente", resultados.filter(r => statusEfetivo(r) === "divergente").length],
    ["NFs duplicadas", resultados.filter(r => statusEfetivo(r) === "duplicada").length],
    ["NFs com Custo/Despesa divergente (planilha × Conta Azul)", resultados.filter(r => r.custoDespesaDivergente).length],
    [],
    ["Itens com edição manual", resultados.filter(r => r.statusManual || r.custoDespesaManual || r.observacao).length],
  ];
  const wsResumo = XLSX.utils.aoa_to_sheet(linhasResumo);
  XLSX.utils.book_append_sheet(wb, wsResumo, "Resumo");

  XLSX.writeFile(wb, `Relatorio_Divergencias_${rotuloFonte}.xlsx`);
}

/* ============= Exportar planilha original com Custo/Despesa/Aceite preenchidos =============
 * Lê o ArrayBuffer original do arquivo (preservando tudo: cores, estilos, merged cells,
 * títulos, abas), localiza apenas as células alvo pelo cabeçalho, modifica só o valor
 * dessas células e devolve o arquivo praticamente intacto.
 */
function gerarPlanilhaPreenchida(buffer, resultados, tipoPlanilha, nomeArquivo) {
  // Re-lê direto do buffer original com suporte a estilos
  const wb = XLSX.read(buffer, { type: "array", cellStyles: true, cellNF: true, cellDates: true });

  const mapa = {};
  for (const r of resultados) {
    if (!mapa[r.nf.aba]) mapa[r.nf.aba] = {};
    mapa[r.nf.aba][String(r.nf.numNF)] = r;
  }

  const relacionadosDe = (r) => r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);

  for (const aba of wb.SheetNames) {
    const ws = wb.Sheets[aba];
    const dados = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!dados || dados.length === 0) continue;

    let headerRow = -1, idxNum = -1, idxCusto = -1, idxDespesa = -1, idxAceite = -1;

    for (let i = 0; i < Math.min(5, dados.length); i++) {
      const row = dados[i] || [];
      for (let j = 0; j < row.length; j++) {
        const h = normalizarTexto(String(row[j] || "")).replace(/\s+/g, " ").trim();
        if (/^num\s*nf[e-]?$|^numero nfs[\s-]?e$|^n[ºo]?\s*nfs[\s-]?e$|^n[ºo]?\s*nfse$/.test(h)) {
          headerRow = i; idxNum = j;
        }
        if (h === "custo" || h === "custo ") idxCusto = j;
        if (h === "despesa") idxDespesa = j;
        if (h.startsWith("aceite")) idxAceite = j;
      }
      if (headerRow === i) break;
    }
    if (headerRow === -1 || (idxCusto === -1 && idxDespesa === -1 && idxAceite === -1)) continue;

    const abaMapa = mapa[aba] || {};

    for (let rowIdx = headerRow + 1; rowIdx < dados.length; rowIdx++) {
      const row = dados[rowIdx];
      if (!row) continue;
      const numNF = row[idxNum];
      if (numNF === null || numNF === undefined || numNF === "") continue;
      if (String(row[0] || "").toLowerCase().includes("total")) continue;

      const r = abaMapa[String(numNF)];
      if (!r) continue;

      const st = statusEfetivo(r);
      const relacionados = relacionadosDe(r);
      const classif = r.custoDespesaManual || r.custoDespesaPlanilha ||
        Array.from(new Set(relacionados.map(l => l.custoDespesa).filter(Boolean))).join("/");

      // rowIdx já é 0-based (sheet_to_json header:1 começa em 0)
      // mas a linha real na sheet é rowIdx (0-based = linha real já que header:1 não offset)
      const sheetRow = rowIdx;

      const setCelula = (colIdx, valor) => {
        if (colIdx < 0) return;
        const addr = XLSX.utils.encode_cell({ r: sheetRow, c: colIdx });
        // Preserva célula existente (estilo incluído) — troca apenas valor e tipo
        const existente = ws[addr] ? { ...ws[addr] } : {};
        ws[addr] = {
          ...existente,
          v: valor === "" ? null : valor,
          t: typeof valor === "number" ? "n" : "s",
        };
      };

      if (idxCusto >= 0) {
        const val = classif === "CUSTO" ? r.nf.valor
          : classif === "CUSTO + DESPESA" ? (r.nf.custoPlanilha ?? "") : "";
        setCelula(idxCusto, val);
      }
      if (idxDespesa >= 0) {
        const val = classif === "DESPESA" ? r.nf.valor
          : classif === "CUSTO + DESPESA" ? (r.nf.despesaPlanilha ?? "") : "";
        setCelula(idxDespesa, val);
      }
      if (idxAceite >= 0) {
        setCelula(idxAceite, (st === "ok" || st === "parcela") ? "ACEITE" : "RECUSA");
      }
    }
  }

  const rotulo = tipoPlanilha === "nfse" ? "NFS-e" : "SIEG";
  const nome = nomeArquivo
    ? nomeArquivo.replace(/\.(xlsx?|xls)$/i, "") + "_preenchida.xlsx"
    : `Planilha_${rotulo}_preenchida.xlsx`;

  XLSX.writeFile(wb, nome, { bookSST: true, cellStyles: true });
}

/* ============================ Componentes UI ============================ */
function DropZone({ label, sublabel, onFile, fileName }) {
  const [drag, setDrag] = useState(false);
  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDrag(false);
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  }, [onFile]);

  return (
    <div
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={handleDrop}
      onClick={() => document.getElementById(`file-${label}`).click()}
      style={{
        flex: 1, minWidth: 220, cursor: "pointer",
        border: `1.5px dashed ${drag ? CORES.accent : CORES.borda}`,
        borderRadius: 14,
        padding: "26px 20px",
        textAlign: "center",
        background: drag ? CORES.accentSoft : CORES.bgCard,
        transition: "all 0.18s ease",
      }}
    >
      <input id={`file-${label}`} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
        onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
      <div style={{
        width: 44, height: 44, borderRadius: 12, margin: "0 auto 12px",
        background: fileName ? CORES.okSoft : CORES.accentSoft,
        display: "flex", alignItems: "center", justifyContent: "center",
      }}>
        {fileName
          ? <CheckCircle2 size={22} color={CORES.ok} />
          : <UploadCloud size={22} color={CORES.accent} />}
      </div>
      <div style={{ fontWeight: 700, color: CORES.texto, marginBottom: 3, fontSize: 14.5 }}>{label}</div>
      <div style={{ fontSize: 12.5, color: CORES.textoSub, marginBottom: 10 }}>{sublabel}</div>
      {fileName
        ? <div style={{ fontSize: 12.5, color: CORES.ok, fontWeight: 600, background: CORES.okSoft, borderRadius: 7, padding: "4px 10px", display: "inline-block" }}>{fileName}</div>
        : <div style={{ fontSize: 12, color: CORES.textoFraco }}>Clique ou arraste o arquivo aqui</div>}
    </div>
  );
}

function Badge({ tipo }) {
  const c = STATUS_CONFIG[tipo] || STATUS_CONFIG.ok;
  const Icon = c.Icon;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5,
      background: c.bg, color: c.color, borderRadius: 999,
      padding: "4px 10px 4px 8px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap",
      border: `1px solid ${c.color}33`,
    }}>
      <Icon size={13} />
      {c.label}
    </span>
  );
}

function InfoTooltip({ texto }) {
  const [aberto, setAberto] = useState(false);
  return (
    <span
      style={{ position: "relative", display: "inline-flex", verticalAlign: "middle", marginLeft: 6 }}
      onMouseEnter={() => setAberto(true)}
      onMouseLeave={() => setAberto(false)}
    >
      <Info size={15} color={CORES.parcela} style={{ cursor: "help" }} />
      {aberto && (
        <span style={{
          position: "absolute", bottom: "150%", left: "50%", transform: "translateX(-50%)",
          background: CORES.bgCardAlt, color: CORES.texto, fontSize: 12, fontWeight: 400, lineHeight: 1.45,
          padding: "10px 12px", borderRadius: 10, width: 230, zIndex: 20,
          border: `1px solid ${CORES.bordaForte}`,
          boxShadow: "0 8px 24px rgba(0,0,0,0.45)",
        }}>
          {texto}
          <span style={{
            position: "absolute", top: "100%", left: "50%", transform: "translateX(-50%)",
            width: 0, height: 0, borderWidth: "5px", borderStyle: "solid",
            borderColor: `${CORES.bgCardAlt} transparent transparent transparent`,
          }} />
        </span>
      )}
    </span>
  );
}

function StatCard({ label, valor, color, Icon, tooltip }) {
  const [hover, setHover] = useState(false);
  return (
    <div
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        flex: "1 1 140px", minWidth: 120,
        background: CORES.bgCard, border: `1px solid ${hover && tooltip ? CORES.bordaForte : CORES.borda}`,
        borderRadius: 14, padding: "16px 16px 14px", borderLeft: `3px solid ${color}`,
        transition: "border-color 0.15s",
      }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <span style={{ fontSize: 11.5, color: CORES.textoSub, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>{label}</span>
        <Icon size={15} color={color} />
      </div>
      <div style={{ fontSize: 26, fontWeight: 700, color: CORES.texto, fontFamily: "'IBM Plex Mono', monospace" }}>{valor}</div>
      {tooltip && hover && (
        <div style={{
          marginTop: 10, paddingTop: 10,
          borderTop: `1px solid ${CORES.borda}`,
          fontSize: 12, color: CORES.textoSub, lineHeight: 1.8,
          whiteSpace: "pre-line",
        }}>
          {tooltip}
        </div>
      )}
    </div>
  );
}

function TagCustoDespesa({ valor, divergente }) {
  if (!valor) return <span style={{ color: CORES.textoFraco, fontSize: 12 }}>—</span>;
  const cor = divergente ? CORES.alerta : (valor === "CUSTO" ? "#7DD3FC" : valor === "DESPESA" ? "#FDBA74" : "#C4B5FD");
  return (
    <span
      title={divergente ? "Classificação diverge entre a planilha e o Conta Azul" : undefined}
      style={{
        display: "inline-flex", alignItems: "center", gap: 4,
        fontSize: 11.5, fontWeight: 700, color: cor,
        background: `${cor}1A`, borderRadius: 6, padding: "3px 8px",
        letterSpacing: "0.02em", whiteSpace: "nowrap",
      }}>
      {divergente && <AlertTriangle size={11} />}
      {valor}
    </span>
  );
}

const STATUS_MANUAL_OPCOES = [
  { val: "ok", label: "OK" },
  { val: "divergente", label: "Divergência" },
  { val: "nao_encontrada", label: "Não lançado" },
  { val: "parcela", label: "Parcela" },
  { val: "duplicada", label: "Duplicada" },
  { val: "devolucao", label: "Devolução" },
];

const CENTROS_DE_CUSTO = [
  "CMFOR OPERACIONAL 1 (Nº 309/2025)",
  "SEHAB-SP OPERACIONAL 1",
  "Call Center Detran-CE OPERACIONAL 1 (Nº 103/2025)",
  "Galpão de Itupeva OPERACIONAL 1",
  "AMC OPERACIONAL 1 (09/2022 4 ADT)",
  "Galpao/SEHAB-SP OPERACIONAL 1",
  "Galpao/Detran-ce OPERACIONAL 1",
  "CUSTO OPERACIONAL 1",
  "Galpao/MEC OPERACIONAL 1",
  "CONAB - DF OPERACIONAL 1 (Nº: 003/2026)",
  "Arrais Veiculos OPERACIONAL 1",
  "CRN OPERACIONAL 1",
  "Galpao/SEHAB OPERACIONAL 1",
  "IPME OPERACIONAL 1 (2024.05.28.001 ADT)",
  "SESC/SENAC OPERACIONAL 1 (N° 407/2022 3° ADT)",
  "APACEFOR OPERACIONAL 1",
  "COMERCIAL NOVO CFC (DESPESA)",
  "Call Center Detran-CE OPERACIONAL 1",
  "MEC-DF OPERACIONAL 1",
  "GALPAO CE OPERACIONAL 1",
  "CONFEA OPERACIONAL 1 (Nº 309/2025)",
  "Galpao/AMC Atendimento OPERACIONAL 1",
  "UFC OPERACIONAL 1",
  "DETRAN PATIO OPERACIONAL 1 (N° 334/2023 2 ADT)",
  "CGE OPERACIONAL 1",
  "PLANEJAMENTO (DESPESA)",
  "ESP CE OPERACIONAL 1 (N°: 15/2024 2° ADT)",
  "Inovação OPERACIONAL",
  "JOSE MURILO CIRINO OPERACIONAL 1",
  "SEGER OPERACIONAL 1 (N° 25/2022 3 ADT)",
  "Compras",
  "FWA OPERACIONAL",
  "CREA-SP OPERACIONAL 1 (N° 19/2024 ADT)",
  "Galpao/GARDEN OPERACIONAL 1",
  "FACILITIES (DESPESA)",
  "ESS VEICULOS OPERACIONAL 1",
  "ETICE - OPERACIONAL 1 (Nº 06/2023 ADT)",
  "SEFAZ OPERACIONAL 1 (N°: 053/2024 ADT)",
  "Galpao/PMA OPERACIONAL 1",
  "HEMOCE OPERACIONAL 1 (Nº 1313/2025)",
  "TRF OPERACIONAL 1 (N° 0405210/2023 3 ADT)",
  "CTC OPERACIONAL 1",
  "Galpao/Omnimagem OPERACIONAL 1",
  "Transito OPERACIONAL",
  "ETUFOR OPERACIONAL 1 (N°: 002/2026)",
  "CEARAPREV OPERACIONAL 1 (N° 002/2024 3° ADT)",
  "FORT GLASS OPERACIONAL 1 (N° 1013/2025)",
  "DIRETORIA",
  "Departamento Pessoal (DESPESA)",
  "Sigam Operacional",
  "OMINIMAGEM OPERACIONAL 1",
  "GODOCS OPERACIONAL",
  "Prefeitura Eusebio OPERACIONAL 1",
  "Sos docs OPERACIONAL 1",
  "EXITO SERVICOS DE APOIO ADMINISTRATIVO OPERACIONAL 1",
  "AGROPECUARIA OPERACIONAL 1",
  "Agile Locação de Veiculos OPERACIONAL 1",
  "Diretoria 2",
  "NOVOCFC OPERACIONAL 1",
  "SINGSERV PRESTACAO DE SERVICOS OPERACIONAL 1",
  "CARTORIO QUERENCIA OPERACIONAL 1",
  "SP LOCACAO DE MAO DE OBRA LTDA OPERACIONAL 1",
  "ÚNICA OPERACIONAL 1",
  "Galpão Eusebio OPERACIONAL 1",
  "Despesas Patrimonial",
  "SME-SP OPERACIONAL 1",
  "Suporte Tecnico OPERACIONAL",
  "TJ RS OPERACIONAL 1",
  "e-Identidade OPERACIONAL",
  "MARKETING (DESPESA)",
  "RH e DEPARTAMENTO PESSOAL (DESPESA)",
  "SDHDS OPERACIONAL 1",
  "Detran-CE Atendimento OPERACIONAL 1",
  "AMC PATIO VEICULAR OPERACIONAL 1",
  "DESPESAS COMERCIAL",
  "DIRETORIA 1",
  "CLUBFUT OPERACIONAL",
  "IHGF OPERACIONAL 1",
  "IBF",
  "AMC DIVIDA ATIVA OPERACIONAL 1",
  "PATIO DETRAN",
  "PATIO AMC",
  "DESENVOLVIMENTO OPERACIONAL",
  "Administrativo (DESPESA)",
  "GARDEN OPERACIONAL 1",
  "Galpao/CMFOR OPERACIONAL 1",
  "Prefeitura Aquiraz OPERACIONAL 1",
  "PRODESP OPERACIONAL 1",
  "PROCON - GO OPERACIONAL 1",
  "Galpão Crea-Sp OPERACIONAL 1",
  "QUALIDADE / PCP OPERACIONAL",
  "Prefeitura Aracati OPERACIONAL 1",
  "Desenvolvimento NOVOCFC OPERACIONAL",
  "Galpao de Brasilia OPERACIONAL 1",
  "CENTEC OPERACIONAL 1 (N° 080/2024)",
  "ArcelorMittal OPERACIONAL 1",
  "DESENVOLVIMENTO CMFOR OPERACIONAL",
  "MEC-DF OPERACIONAL 1 (15/2024 ADT)",
  "SEHAB-SP OPERACIONAL 1 (01/12/2024 ADT)",
  "ALECE OPERACIONAL 1",
  "Desenvolvimento ALECE OPERACIONAL",
  "AMC PATIO VEICULAR OPERACIONAL 1 (N° 36/2023 2 ADT)",
  "Sergio Telerman OPERACIONAL",
  "Galpao/PME OPERACIONAL 1",
  "Galpão SME-SP OPERACIONAL 1",
  "POLÍCIA CIVIL CE OPERACIONAL 1 (Nº 077/2025)",
  "Detran-CE OPERACIONAL 1 (N° 318/2023 2 ADT)",
].sort((a, b) => a.localeCompare(b, "pt-BR"));

const CUSTO_DESPESA_OPCOES = [
  { val: "", label: "Não classificado" },
  { val: "CUSTO", label: "Custo" },
  { val: "DESPESA", label: "Despesa" },
];

function PainelEdicao({ r, onSalvar, onRestaurar, onFechar }) {
  const [statusManual, setStatusManual] = useState(r.statusManual || "");
  const [custoDespesaManual, setCustoDespesaManual] = useState(r.custoDespesaManual || "");
  const [centroCustoManual, setCentroCustoManual] = useState(r.centroCustoManual || "");
  const [observacao, setObservacao] = useState(r.observacao || "");
  const [buscaCentro, setBuscaCentro] = useState("");

  const relacionados = r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
  const valoresCA = relacionados.length
    ? relacionados.map(l => `R$ ${l.valor.toFixed(2)}`).join(" + ")
    : "—";

  const temOverride = !!(r.statusManual || r.custoDespesaManual || r.centroCustoManual || r.observacao);

  const centroAtual = centroCustoManual || (relacionados[0]?.centroCusto || "");
  const centrosFiltrados = buscaCentro.trim()
    ? CENTROS_DE_CUSTO.filter(c => normalizarTexto(c).includes(normalizarTexto(buscaCentro)))
    : CENTROS_DE_CUSTO;

  return (
    <div
      onClick={onFechar}
      style={{
        position: "fixed", inset: 0, background: "rgba(4,7,14,0.65)",
        display: "flex", alignItems: "center", justifyContent: "center",
        zIndex: 100, padding: 18,
      }}
    >
      <div
        onClick={e => e.stopPropagation()}
        style={{
          width: "100%", maxWidth: 460, maxHeight: "88vh", overflow: "auto",
          background: CORES.bgCard, border: `1px solid ${CORES.bordaForte}`, borderRadius: 16,
          padding: 22, boxShadow: "0 24px 60px rgba(0,0,0,0.5)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 800 }}>NF {r.nf.numNF}</div>
            <div style={{ fontSize: 12.5, color: CORES.textoSub, marginTop: 2 }}>{r.nf.fornecedor}</div>
          </div>
          <button onClick={onFechar} style={{
            border: "none", background: CORES.bgCardAlt, borderRadius: 8, padding: 6,
            cursor: "pointer", display: "flex", color: CORES.textoSub,
          }}>
            <X size={16} />
          </button>
        </div>

        <div style={{
          display: "flex", justifyContent: "space-between", fontSize: 13, color: CORES.textoSub,
          background: CORES.bgCardAlt, borderRadius: 10, padding: "10px 12px", margin: "14px 0 18px",
        }}>
          <span>Valor da nota: <strong style={{ color: CORES.texto }}>R$ {r.nf.valor.toFixed(2)}</strong></span>
          <span>CA: <strong style={{ color: CORES.texto }}>{valoresCA}</strong></span>
        </div>

        <div style={{ fontSize: 12, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 8 }}>
          Status
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginBottom: 18 }}>
          {STATUS_MANUAL_OPCOES.map(op => {
            const ativo = statusManual === op.val;
            const cor = STATUS_CONFIG[op.val]?.color || CORES.texto;
            return (
              <button
                key={op.val}
                onClick={() => setStatusManual(ativo ? "" : op.val)}
                style={{
                  padding: "7px 13px", borderRadius: 999, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                  border: `1px solid ${ativo ? cor : CORES.borda}`,
                  background: ativo ? `${cor}1F` : "transparent",
                  color: ativo ? cor : CORES.textoSub,
                }}
              >{op.label}</button>
            );
          })}
        </div>
        {!statusManual && (
          <div style={{ fontSize: 11.5, color: CORES.textoFraco, marginTop: -12, marginBottom: 18 }}>
            Sem seleção = mantém a classificação automática ({STATUS_CONFIG[r.tipo]?.label || r.tipo}).
          </div>
        )}

        <div style={{ fontSize: 12, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 8 }}>
          Custo / Despesa
        </div>
        <div style={{ display: "flex", gap: 7, marginBottom: 18 }}>
          {CUSTO_DESPESA_OPCOES.map(op => {
            const ativo = custoDespesaManual === op.val;
            return (
              <button
                key={op.label}
                onClick={() => setCustoDespesaManual(op.val)}
                style={{
                  flex: 1, padding: "8px 10px", borderRadius: 9, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                  border: `1px solid ${ativo ? CORES.accentBorder : CORES.borda}`,
                  background: ativo ? CORES.accentSoft : "transparent",
                  color: ativo ? CORES.accent : CORES.textoSub,
                }}
              >{op.label}</button>
            );
          })}
        </div>

        <div style={{ fontSize: 12, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 8 }}>
          Centro de Custo
        </div>
        {centroAtual && (
          <div style={{ fontSize: 12, color: CORES.textoSub, marginBottom: 6 }}>
            Atual: <strong style={{ color: CORES.texto }}>{centroAtual}</strong>
            {centroCustoManual && (
              <button onClick={() => { setCentroCustoManual(""); setBuscaCentro(""); }} style={{
                marginLeft: 8, fontSize: 11, color: CORES.textoFraco, background: "none",
                border: "none", cursor: "pointer", textDecoration: "underline",
              }}>limpar</button>
            )}
          </div>
        )}
        <input
          value={buscaCentro}
          onChange={e => setBuscaCentro(e.target.value)}
          placeholder="Buscar centro de custo…"
          style={{
            width: "100%", borderRadius: 8, padding: "8px 12px", marginBottom: 6,
            background: CORES.bgCardAlt, border: `1px solid ${CORES.borda}`,
            color: CORES.texto, fontSize: 13, fontFamily: "inherit", outline: "none",
          }}
        />
        <div style={{
          maxHeight: 160, overflowY: "auto", borderRadius: 8,
          border: `1px solid ${CORES.borda}`, background: CORES.bgCardAlt, marginBottom: 18,
        }}>
          {centrosFiltrados.length === 0 ? (
            <div style={{ padding: "10px 12px", fontSize: 12.5, color: CORES.textoFraco }}>Nenhum resultado.</div>
          ) : centrosFiltrados.map(c => (
            <div
              key={c}
              onClick={() => { setCentroCustoManual(c); setBuscaCentro(""); }}
              style={{
                padding: "7px 12px", fontSize: 12.5, cursor: "pointer",
                color: centroCustoManual === c ? CORES.accent : CORES.texto,
                background: centroCustoManual === c ? CORES.accentSoft : "transparent",
                borderBottom: `1px solid ${CORES.borda}`,
              }}
              onMouseEnter={e => { if (centroCustoManual !== c) e.currentTarget.style.background = CORES.bgCardHover; }}
              onMouseLeave={e => { if (centroCustoManual !== c) e.currentTarget.style.background = "transparent"; }}
            >{c}</div>
          ))}
        </div>

        <div style={{ fontSize: 12, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 8 }}>
          Observação
        </div>
        <textarea
          value={observacao}
          onChange={e => setObservacao(e.target.value)}
          placeholder="Anotações sobre esta nota…"
          rows={3}
          style={{
            width: "100%", resize: "vertical", borderRadius: 10, padding: "10px 12px",
            background: CORES.bgCardAlt, border: `1px solid ${CORES.borda}`, color: CORES.texto,
            fontSize: 13, fontFamily: "inherit", marginBottom: 20, outline: "none",
          }}
        />

        <div style={{ display: "flex", gap: 10 }}>
          {temOverride && (
            <button
              onClick={() => onRestaurar(r.id)}
              title="Remove o status e a classificação manual de custo/despesa. A observação é preservada."
              style={{
                display: "flex", alignItems: "center", gap: 6, padding: "11px 14px", borderRadius: 10,
                border: `1px solid ${CORES.borda}`, background: "transparent", color: CORES.textoSub,
                fontWeight: 600, fontSize: 13, cursor: "pointer",
              }}
            >
              <RotateCcw size={14} /> Restaurar automático
            </button>
          )}
          <button
            onClick={() => onSalvar(r.id, { statusManual, custoDespesaManual, centroCustoManual, observacao })}
            style={{
              flex: 1, padding: "11px 14px", borderRadius: 10, border: "none",
              background: `linear-gradient(135deg, ${CORES.accent}, ${CORES.accentDark})`,
              color: "#fff", fontWeight: 700, fontSize: 13.5, cursor: "pointer",
            }}
          >
            Salvar
          </button>
        </div>
      </div>
    </div>
  );
}

function ColunaFiltro({ label, col, todos, selecionados, onToggle, onLimpar, ordemCol, ordemDir, onOrdem, dropdownAberto, onAbrirDropdown }) {
  const ativo = ordemCol === col;
  const temFiltro = selecionados !== null;
  const seta = ativo ? (ordemDir === "asc" ? "↑" : "↓") : null;
  const todosValores = [...todos].sort((a, b) => {
    const na = parseInt(a, 10), nb = parseInt(b, 10);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b, "pt-BR");
  });
  const aberto = dropdownAberto === col;
  const todosMarcados = !selecionados || selecionados.size === todos.size;

  return (
    <th style={{
      padding: "12px 18px", textAlign: "left", fontWeight: 600, fontSize: 11.5,
      color: (ativo || temFiltro) ? CORES.accent : CORES.textoSub,
      textTransform: "uppercase", letterSpacing: "0.03em",
      whiteSpace: "nowrap", borderBottom: `1px solid ${CORES.borda}`,
      position: "relative", userSelect: "none",
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
        <span
          onClick={() => onOrdem && onOrdem(col)}
          title={!onOrdem ? undefined : ativo ? (ordemDir === "asc" ? "Ordem decrescente" : "Ordem crescente") : "Ordenar"}
          style={{ cursor: onOrdem ? "pointer" : "default", display: "flex", alignItems: "center", gap: 3 }}
        >
          {label}
          {onOrdem && <span style={{ fontSize: 11, opacity: ativo ? 1 : 0.35 }}>{seta || "↕"}</span>}
        </span>
        <span
          onClick={(e) => { e.stopPropagation(); onAbrirDropdown(aberto ? null : col); }}
          title="Filtrar"
          style={{
            cursor: "pointer", display: "inline-flex", alignItems: "center", justifyContent: "center",
            width: 18, height: 18, borderRadius: 4, marginLeft: 2,
            background: temFiltro ? CORES.accentSoft : "transparent",
            border: `1px solid ${temFiltro ? CORES.accentBorder : CORES.borda}`,
            color: temFiltro ? CORES.accent : CORES.textoSub,
            fontSize: 11, fontWeight: 700,
          }}
        >▾</span>
      </div>

      {aberto && (
        <div
          onClick={e => e.stopPropagation()}
          style={{
            position: "absolute", top: "100%", left: 0, zIndex: 200,
            minWidth: 220, maxWidth: 320, maxHeight: 320, overflowY: "auto",
            background: CORES.bgCard, border: `1px solid ${CORES.bordaForte}`,
            borderRadius: 10, boxShadow: "0 12px 36px rgba(0,0,0,0.5)",
            padding: "8px 0",
          }}
        >
          {/* Selecionar tudo / Limpar */}
          <div style={{ padding: "6px 12px 8px", borderBottom: `1px solid ${CORES.borda}`, display: "flex", gap: 8 }}>
            <button
              onClick={() => onLimpar(col)}
              style={{
                flex: 1, padding: "5px 8px", borderRadius: 7, fontSize: 11.5, fontWeight: 600, cursor: "pointer",
                border: `1px solid ${CORES.borda}`, background: "transparent", color: CORES.textoSub,
              }}
            >Limpar filtro</button>
            <button
              onClick={() => onToggle(col, null, "todos")}
              style={{
                flex: 1, padding: "5px 8px", borderRadius: 7, fontSize: 11.5, fontWeight: 600, cursor: "pointer",
                border: `1px solid ${CORES.borda}`, background: "transparent", color: CORES.textoSub,
              }}
            >Selec. tudo</button>
          </div>

          {/* Lista de valores */}
          {todosValores.map(val => {
            const marcado = !selecionados || selecionados.has(val);
            return (
              <label
                key={val}
                style={{
                  display: "flex", alignItems: "center", gap: 9,
                  padding: "6px 14px", cursor: "pointer", fontSize: 12.5,
                  color: marcado ? CORES.texto : CORES.textoFraco,
                  background: "transparent",
                }}
                onMouseEnter={e => e.currentTarget.style.background = CORES.bgCardHover}
                onMouseLeave={e => e.currentTarget.style.background = "transparent"}
              >
                <input
                  type="checkbox"
                  checked={marcado}
                  onChange={() => onToggle(col, val)}
                  style={{ accentColor: CORES.accent, width: 14, height: 14, cursor: "pointer" }}
                />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{val}</span>
              </label>
            );
          })}
        </div>
      )}
    </th>
  );
}

function TagCentroCusto({ relacionados, override }) {
  const [aberto, setAberto] = useState(false);

  // Se houver override manual, exibe só ele com indicador de edição
  if (override) {
    return (
      <span style={{
        fontSize: 11.5, color: CORES.accent, whiteSpace: "nowrap",
        overflow: "hidden", textOverflow: "ellipsis", display: "block", maxWidth: 180,
        fontStyle: "italic",
      }} title={override}>
        {override}
      </span>
    );
  }

  const centros = relacionados
    .filter(l => l.centroCusto)
    .map(l => ({ nome: l.centroCusto, valor: l.valor }));

  const agrupado = Object.values(
    centros.reduce((acc, { nome, valor }) => {
      acc[nome] = acc[nome] || { nome, valor: 0 };
      acc[nome].valor += valor;
      return acc;
    }, {})
  );

  if (agrupado.length === 0) return <span style={{ color: CORES.textoFraco, fontSize: 12 }}>—</span>;

  if (agrupado.length === 1) {
    return (
      <span style={{
        fontSize: 11.5, color: CORES.textoSub, whiteSpace: "nowrap",
        overflow: "hidden", textOverflow: "ellipsis", display: "block", maxWidth: 180,
      }} title={agrupado[0].nome}>
        {agrupado[0].nome}
      </span>
    );
  }

  return (
    <div style={{ position: "relative" }}>
      <span
        onClick={e => { e.stopPropagation(); setAberto(a => !a); }}
        style={{
          fontSize: 12, fontWeight: 700, color: CORES.accent,
          background: CORES.accentSoft, border: `1px solid ${CORES.accentBorder}`,
          borderRadius: 6, padding: "3px 9px", cursor: "pointer", whiteSpace: "nowrap",
          display: "inline-block",
        }}
      >
        {agrupado.length} Centros de Custo
      </span>

      {aberto && (
        <div
          onClick={e => e.stopPropagation()}
          style={{
            position: "absolute", top: "calc(100% + 6px)", left: 0, zIndex: 300,
            minWidth: 260, maxWidth: 360,
            background: CORES.bgCard, border: `1px solid ${CORES.bordaForte}`,
            borderRadius: 10, boxShadow: "0 12px 36px rgba(0,0,0,0.55)",
            padding: "10px 0 6px",
          }}
        >
          <div style={{ padding: "0 14px 8px", fontSize: 11, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${CORES.borda}` }}>
            Centros de Custo
          </div>
          {agrupado.map(({ nome, valor }) => (
            <div key={nome} style={{
              display: "flex", justifyContent: "space-between", alignItems: "center",
              padding: "7px 14px", gap: 12,
            }}>
              <span style={{ fontSize: 12.5, color: CORES.texto, flex: 1 }}>{nome}</span>
              <span style={{ fontSize: 12, fontFamily: "'IBM Plex Mono', monospace", color: CORES.textoSub, whiteSpace: "nowrap" }}>
                R$ {valor.toFixed(2)}
              </span>
            </div>
          ))}
          <div style={{ padding: "6px 14px 0", borderTop: `1px solid ${CORES.borda}`, marginTop: 4 }}>
            <button
              onClick={e => { e.stopPropagation(); setAberto(false); }}
              style={{
                width: "100%", padding: "5px", border: "none", background: "transparent",
                color: CORES.textoFraco, fontSize: 12, cursor: "pointer",
              }}
            >Fechar</button>
          </div>
        </div>
      )}
    </div>
  );
}

function LinhaTabela({ r, onEditar }) {
  const [hover, setHover] = useState(false);
  const tipoEfetivo = statusEfetivo(r);
  const c = STATUS_CONFIG[tipoEfetivo] || STATUS_CONFIG.ok;
  const relacionados = r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
  const custoDespesa = r.custoDespesaManual || r.custoDespesaPlanilha || Array.from(new Set(relacionados.map(l => l.custoDespesa).filter(Boolean))).join(" / ");
  const foiEditado = !!(r.statusManual || r.custoDespesaManual || r.centroCustoManual || r.observacao);

  return (
    <tr
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onClick={() => onEditar(r.id)}
      style={{ borderBottom: `1px solid ${CORES.borda}`, background: hover ? CORES.bgCardHover : "transparent", transition: "background 0.12s", cursor: "pointer" }}
    >
      <td style={{ padding: "12px 18px", borderLeft: `3px solid ${c.color}` }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          <Badge tipo={tipoEfetivo} />
          {isParcelaEfetiva(r) && !r.statusManual && <InfoTooltip texto={r.detalhes} />}
          {foiEditado && (
            <span title="Editado manualmente" style={{ display: "inline-flex", marginLeft: 6 }}>
              <Pencil size={12} color={CORES.textoSub} />
            </span>
          )}
        </div>
      </td>
      <td style={{ padding: "12px 18px", color: CORES.textoSub, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, whiteSpace: "nowrap" }}>
        {r.nf.dataEmissao ? String(r.nf.dataEmissao).slice(0, 10) : "—"}
      </td>
      <td style={{ padding: "12px 18px", color: CORES.textoSub }}>{r.nf.aba}</td>
      <td style={{ padding: "12px 18px", fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace" }}>NF {r.nf.numNF}</td>
      <td style={{ padding: "12px 18px", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.nf.fornecedor}>{r.nf.fornecedor}</td>
      <td style={{ padding: "12px 18px", fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace" }}>R$ {r.nf.valor.toFixed(2)}</td>
      <td style={{ padding: "12px 18px", fontFamily: "'IBM Plex Mono', monospace", color: tipoEfetivo === "divergente" ? CORES.alerta : CORES.textoSub }}>
        {relacionados.length > 1
          ? relacionados.map(l => `R$${l.valor.toFixed(2)}`).join(" + ")
          : relacionados.length === 1 ? `R$ ${relacionados[0].valor.toFixed(2)}` : "—"}
      </td>
      <td style={{ padding: "12px 18px" }} onClick={e => e.stopPropagation()}>
        <TagCentroCusto relacionados={relacionados} override={r.centroCustoManual} />
      </td>
      <td style={{ padding: "12px 18px" }}><TagCustoDespesa valor={custoDespesa} divergente={!r.custoDespesaManual && r.custoDespesaDivergente} /></td>
      <td style={{ padding: "12px 18px", fontSize: 12, color: tipoEfetivo === "ok" ? (isParcelaEfetiva(r) ? CORES.textoSub : CORES.ok) : c.color }}>
        {r.detalhes}
        <Pencil size={11} color={CORES.textoFraco} style={{ marginLeft: 7, opacity: hover ? 1 : 0, transition: "opacity 0.12s", verticalAlign: "middle" }} />
      </td>
      <td style={{ padding: "12px 18px", fontSize: 12, color: CORES.textoSub, fontStyle: r.observacao ? "normal" : "italic" }}>
        {r.observacao || "—"}
      </td>
    </tr>
  );
}

/* ================================ App ================================ */

// Chave única por nota no storage persistente: inclui o tipo de planilha, a aba e o
// número da nota para garantir que edições de sessões anteriores sejam reaplicadas
// mesmo que a posição da nota na tabela tenha mudado.
function chaveNF(tipoPlanilha, aba, numNF) {
  return ("nf:" + tipoPlanilha + ":" + aba + ":" + numNF).replace(/[^a-zA-Z0-9:._-]/g, "_");
}

async function carregarEdicoesSalvas(tipoPlanilha, resultados) {
  const editadasComStorage = [];
  for (const r of resultados) {
    try {
      const chave = chaveNF(tipoPlanilha, r.nf.aba, r.nf.numNF);
      const saved = await window.storage.get(chave);
      if (saved) {
        const ed = JSON.parse(saved.value);
        editadasComStorage.push({
          ...r,
          statusManual: ed.statusManual || undefined,
          custoDespesaManual: ed.custoDespesaManual || undefined,
          centroCustoManual: ed.centroCustoManual || undefined,
          observacao: ed.observacao || undefined,
        });
      } else {
        editadasComStorage.push(r);
      }
    } catch {
      editadasComStorage.push(r);
    }
  }
  return editadasComStorage;
}

async function persistirEdicao(tipoPlanilha, r, edicao) {
  const chave = chaveNF(tipoPlanilha, r.nf.aba, r.nf.numNF);
  const temDado = edicao.statusManual || edicao.custoDespesaManual || edicao.centroCustoManual || edicao.observacao;
  try {
    if (temDado) {
      await window.storage.set(chave, JSON.stringify(edicao));
    } else {
      await window.storage.delete(chave);
    }
  } catch {
    // storage indisponível: continua sem persistência (não quebra o fluxo)
  }
}

const MAX_SNAPSHOTS = 10;
const PREFIXO_SNAP = "snap:";

async function salvarSnapshot(tipoPlanilha, resultados, nome) {
  try {
    const ts = Date.now();
    const chave = PREFIXO_SNAP + tipoPlanilha + ":" + ts;
    const editados = resultados.filter(r => r.statusManual || r.custoDespesaManual || r.observacao);
    const payload = {
      ts,
      nome: nome || "",
      tipoPlanilha,
      total: resultados.length,
      editados: editados.length,
      itens: resultados.map(r => ({
        aba: r.nf.aba, numNF: r.nf.numNF,
        statusManual: r.statusManual,
        custoDespesaManual: r.custoDespesaManual,
        centroCustoManual: r.centroCustoManual,
        observacao: r.observacao,
      })),
    };
    await window.storage.set(chave, JSON.stringify(payload));
    // Limitar a MAX_SNAPSHOTS versões — apagar as mais antigas se exceder
    const lista = await listarSnapshots(tipoPlanilha);
    if (lista.length > MAX_SNAPSHOTS) {
      const paraApagar = lista.slice(MAX_SNAPSHOTS);
      for (const s of paraApagar) {
        await window.storage.delete(PREFIXO_SNAP + tipoPlanilha + ":" + s.ts);
      }
    }
    return ts;
  } catch { return null; }
}

async function listarSnapshots(tipoPlanilha) {
  try {
    const prefix = PREFIXO_SNAP + tipoPlanilha + ":";
    const keys = await window.storage.list(prefix);
    const snaps = [];
    for (const key of (keys.keys || [])) {
      try {
        const raw = await window.storage.get(key);
        if (raw) snaps.push(JSON.parse(raw.value));
      } catch { /* ignora entradas corrompidas */ }
    }
    return snaps.sort((a, b) => b.ts - a.ts); // mais recente primeiro
  } catch { return []; }
}

async function deletarSnapshot(tipoPlanilha, ts) {
  try {
    await window.storage.delete(PREFIXO_SNAP + tipoPlanilha + ":" + ts);
  } catch { /* ignora */ }
}

// Aplica as edições de um snapshot sobre os resultados atuais.
// Só sobrescreve campos de edição manual — não altera a classificação automática.
function aplicarSnapshot(snap, resultados) {
  const mapa = new Map(snap.itens.map(i => [i.aba + "|" + i.numNF, i]));
  return resultados.map(r => {
    const ed = mapa.get(r.nf.aba + "|" + r.nf.numNF);
    if (!ed) return r;
    return {
      ...r,
      statusManual: ed.statusManual || undefined,
      custoDespesaManual: ed.custoDespesaManual || undefined,
      centroCustoManual: ed.centroCustoManual || undefined,
      observacao: ed.observacao || undefined,
    };
  });
}

export default function App() {
  const [tipoPlanilha, setTipoPlanilha] = useState("nfe"); // "nfe" (SIEG) | "nfse"
  const [siegFile, setSiegFile] = useState(null);
  const [caFile, setCaFile] = useState(null);
  const [resultados, setResultados] = useState(null);
  const [processando, setProcessando] = useState(false);
  const [filtro, setFiltro] = useState("todos");
  const [erro, setErro] = useState("");
  const [editandoId, setEditandoId] = useState(null);
  const [wbNotas, setWbNotas] = useState(null); // workbook original da planilha de notas, para exportar com preenchimento
  const [wbNotasBuffer, setWbNotasBuffer] = useState(null); // ArrayBuffer original para preservar formatação
  const [painelVersions, setPainelVersions] = useState(false);
  const [snapshots, setSnapshots] = useState([]);
  const [salvandoSnap, setSalvandoSnap] = useState(false);
  const [nomeSnap, setNomeSnap] = useState("");
  const [msgSnap, setMsgSnap] = useState("");
  const [ordemCol, setOrdemCol] = useState(null);
  const [ordemDir, setOrdemDir] = useState("asc");
  const FILTRO_COL_VAZIO = { status: null, data: null, filial: null, numNF: null, fornecedor: null, valorSIEG: null, valorCA: null, centroCusto: null, custoDespesa: null };
  const [filtroCol, setFiltroCol] = useState(FILTRO_COL_VAZIO);
  const [dropdownAberto, setDropdownAberto] = useState(null);

  const abrirVersions = async () => {
    const lista = await listarSnapshots(tipoPlanilha);
    setSnapshots(lista);
    setPainelVersions(true);
  };

  const handleSalvarSnap = async () => {
    if (!resultados) return;
    setSalvandoSnap(true);
    const ts = await salvarSnapshot(tipoPlanilha, resultados, nomeSnap.trim());
    if (ts) {
      setMsgSnap("Versão salva!");
      setNomeSnap("");
      const lista = await listarSnapshots(tipoPlanilha);
      setSnapshots(lista);
    } else {
      setMsgSnap("Erro ao salvar.");
    }
    setSalvandoSnap(false);
    setTimeout(() => setMsgSnap(""), 2500);
  };

  const handleRestaurarSnap = async (snap) => {
    if (!resultados) return;
    const novo = aplicarSnapshot(snap, resultados);
    setResultados(novo);
    setPainelVersions(false);
  };

  const handleDeletarSnap = async (snap) => {
    await deletarSnapshot(tipoPlanilha, snap.ts);
    const lista = await listarSnapshots(tipoPlanilha);
    setSnapshots(lista);
  };

  const alternarOrdem = (col) => {
    if (ordemCol === col) {
      setOrdemDir(d => d === "asc" ? "desc" : "asc");
    } else {
      setOrdemCol(col);
      setOrdemDir("asc");
    }
  };

  const salvarEdicao = (id, edicao) => {
    setResultados(prev => {
      const next = prev.map(r => r.id !== id ? r : {
        ...r,
        statusManual: edicao.statusManual || undefined,
        custoDespesaManual: edicao.custoDespesaManual || undefined,
        centroCustoManual: edicao.centroCustoManual || undefined,
        observacao: edicao.observacao || undefined,
      });
      const alvo = next.find(r => r.id === id);
      if (alvo) persistirEdicao(tipoPlanilha, alvo, edicao);
      return next;
    });
    setEditandoId(null);
  };

  const restaurarAutomatico = (id) => {
    setResultados(prev => {
      const alvo = prev.find(r => r.id === id);
      if (alvo) persistirEdicao(tipoPlanilha, alvo, {});
      return prev.map(r => r.id !== id ? r : {
        ...r,
        statusManual: undefined,
        custoDespesaManual: undefined,
        centroCustoManual: undefined,
      });
    });
    setEditandoId(null);
  };

  const conferir = async () => {
    if (!siegFile || !caFile) { setErro("Anexe os dois arquivos antes de conferir."); return; }
    setProcessando(true); setErro("");
    try {
      const lerWB = (file) => new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = e => res(XLSX.read(e.target.result, { type: "array" }));
        reader.onerror = rej;
        reader.readAsArrayBuffer(file);
      });
      const lerBuffer = (file) => new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = e => res(e.target.result);
        reader.onerror = rej;
        reader.readAsArrayBuffer(file);
      });
      const [[wbNotasParsed, bufferNotas], wbCA] = await Promise.all([
        Promise.all([lerWB(siegFile), lerBuffer(siegFile)]),
        lerWB(caFile),
      ]);
      setWbNotas(wbNotasParsed);
      setWbNotasBuffer(bufferNotas);
      const notas = tipoPlanilha === "nfse" ? lerPlanilhaNFSE(wbNotasParsed) : lerPlanilhaSIEG(wbNotasParsed);
      const lancamentos = lerContaAzul(wbCA);
      const rotulo = tipoPlanilha === "nfse" ? "NFS-e" : "SIEG";
      if (notas.length === 0) { setErro(`Não foi possível ler as notas da planilha ${rotulo}. Verifique o formato.`); setProcessando(false); return; }
      if (lancamentos.length === 0) { setErro("Não foi possível ler os lançamentos do Conta Azul. Verifique o formato."); setProcessando(false); return; }
      const comparados = compararDados(notas, lancamentos);
      // Reaplicar edições manuais salvas de sessões anteriores automaticamente
      const comEdicoes = await carregarEdicoesSalvas(tipoPlanilha, comparados);
      setResultados(comEdicoes);
      setFiltroCol(FILTRO_COL_VAZIO);
      setOrdemCol(null);
    } catch (e) {
      setErro("Erro ao processar arquivos: " + e.message);
    }
    setProcessando(false);
  };

  const getCustoDespesaR = (r) => {
    const rel = r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
    return r.custoDespesaManual || r.custoDespesaPlanilha || Array.from(new Set(rel.map(l => l.custoDespesa).filter(Boolean))).join("/") || "—";
  };

  const getCentroR = (r) => {
    const rel = r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
    const centros = [...new Set(rel.map(l => l.centroCusto).filter(Boolean))];
    return centros.length === 0 ? "—" : centros.length === 1 ? centros[0] : centros.length + " Centros";
  };

  const getValorCAR = (r) => {
    const rel = r.lancamentosRelacionados || (r.lancamento ? [r.lancamento] : []);
    return rel.length === 0 ? "—" : rel.map(l => `R$${l.valor.toFixed(2)}`).join("+");
  };

  const filtrados = (() => {
    const base = !resultados ? [] :
      filtro === "todos" ? resultados :
      filtro === "divergencias" ? resultados.filter(r => ["nao_encontrada", "divergente", "duplicada", "devolucao"].includes(statusEfetivo(r))) :
      filtro === "parcela" ? resultados.filter(isParcelaEfetiva) :
      resultados.filter(r => statusEfetivo(r) === filtro);

    let f = base;
    if (filtroCol.status)      f = f.filter(r => filtroCol.status.has(STATUS_CONFIG[statusEfetivo(r)]?.label || statusEfetivo(r)));
    if (filtroCol.data)        f = f.filter(r => filtroCol.data.has(r.nf.dataEmissao ? String(r.nf.dataEmissao).slice(0, 10) : "—"));
    if (filtroCol.filial)      f = f.filter(r => filtroCol.filial.has(r.nf.aba));
    if (filtroCol.numNF)       f = f.filter(r => filtroCol.numNF.has(r.nf.numNF));
    if (filtroCol.fornecedor)  f = f.filter(r => filtroCol.fornecedor.has(r.nf.fornecedor));
    if (filtroCol.valorSIEG)   f = f.filter(r => filtroCol.valorSIEG.has(`R$${r.nf.valor.toFixed(2)}`));
    if (filtroCol.valorCA)     f = f.filter(r => filtroCol.valorCA.has(getValorCAR(r)));
    if (filtroCol.centroCusto) f = f.filter(r => filtroCol.centroCusto.has(getCentroR(r)));
    if (filtroCol.custoDespesa)f = f.filter(r => filtroCol.custoDespesa.has(getCustoDespesaR(r)));

    if (!ordemCol) return f;
    return [...f].sort((a, b) => {
      let va, vb;
      if (ordemCol === "numNF") {
        va = parseInt(a.nf.numNF, 10); vb = parseInt(b.nf.numNF, 10);
        if (isNaN(va)) va = 0; if (isNaN(vb)) vb = 0;
        return ordemDir === "asc" ? va - vb : vb - va;
      } else {
        va = (a.nf.fornecedor || "").toLowerCase();
        vb = (b.nf.fornecedor || "").toLowerCase();
        return ordemDir === "asc" ? va.localeCompare(vb, "pt-BR") : vb.localeCompare(va, "pt-BR");
      }
    });
  })();

  // StatCards reagem aos filtros ativos — contam a partir dos itens visíveis na tabela
  const nfsPor = (predicate) => filtrados.filter(predicate).map(r => `NF ${r.nf.numNF}`);
  const formatarNFs = (nfs) => {
    if (nfs.length === 0) return "Nenhuma";
    const MAX = 10;
    const linhas = nfs.slice(0, MAX).join(", ");
    return nfs.length > MAX ? linhas + ` e mais ${nfs.length - MAX}…` : linhas;
  };
  const resumo = resultados ? {
    total: filtrados.length,
    okPuro: filtrados.filter(r => statusEfetivo(r) === "ok" && !isParcelaEfetiva(r) && statusEfetivo(r) !== "devolucao").length,
    parcela: filtrados.filter(isParcelaEfetiva).length,
    devolucao: filtrados.filter(r => statusEfetivo(r) === "devolucao").length,
    get ok() { return this.okPuro + this.parcela + this.devolucao; },
    naoEncontrada: filtrados.filter(r => statusEfetivo(r) === "nao_encontrada").length,
    divergente: filtrados.filter(r => statusEfetivo(r) === "divergente").length,
    duplicada: filtrados.filter(r => statusEfetivo(r) === "duplicada").length,
    tooltips: {
      naoEncontrada: formatarNFs(nfsPor(r => statusEfetivo(r) === "nao_encontrada")),
      divergente: formatarNFs(nfsPor(r => statusEfetivo(r) === "divergente")),
      parcela: formatarNFs(nfsPor(isParcelaEfetiva)),
      duplicada: formatarNFs(nfsPor(r => statusEfetivo(r) === "duplicada")),
      devolucao: formatarNFs(nfsPor(r => statusEfetivo(r) === "devolucao")),
    },
  } : null;

  const itemEditando = resultados && editandoId !== null ? resultados.find(r => r.id === editandoId) : null;

  return (
    <div style={{ position: "relative", minHeight: "100vh", background: CORES.bgPage, fontFamily: "'Manrope', sans-serif", color: CORES.texto }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        ::selection { background: ${CORES.accentSoft}; }
      `}</style>

      <div style={{
        position: "absolute", top: -180, right: -160, width: 480, height: 480,
        background: `radial-gradient(circle, ${CORES.accent}26, transparent 70%)`,
        pointerEvents: "none",
      }} />

      <div style={{ position: "relative", maxWidth: 980, margin: "0 auto", padding: "44px 20px 70px" }}>

        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{
            width: 46, height: 46, borderRadius: 13,
            background: `linear-gradient(135deg, ${CORES.accent}, ${CORES.accentDark})`,
            display: "flex", alignItems: "center", justifyContent: "center",
            boxShadow: `0 6px 20px ${CORES.accentSoft}`,
          }}>
            <Search size={22} color="#fff" />
          </div>
          <div>
            <div style={{ fontWeight: 800, fontSize: 21, letterSpacing: "-0.01em" }}>Conferência de Notas Fiscais</div>
            <div style={{ fontSize: 13, color: CORES.textoSub }}>{tipoPlanilha === "nfse" ? "NFS-e" : "SIEG"} × Conta Azul — Contas a Pagar</div>
          </div>
        </div>
        <div style={{ height: 1, background: `linear-gradient(90deg, ${CORES.bordaForte}, transparent)`, margin: "24px 0 28px" }} />

        {/* Seletor de tipo de planilha */}
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 8 }}>
            Tipo de planilha de notas
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {[
              { val: "nfe", label: "NF-e (SIEG)" },
              { val: "nfse", label: "NFS-e (serviços)" },
            ].map(op => {
              const ativo = tipoPlanilha === op.val;
              return (
                <button
                  key={op.val}
                  onClick={() => { setTipoPlanilha(op.val); setSiegFile(null); setResultados(null); setErro(""); setWbNotas(null); setWbNotasBuffer(null); }}
                  style={{
                    flex: 1, padding: "10px 14px", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer",
                    border: `1px solid ${ativo ? CORES.accentBorder : CORES.borda}`,
                    background: ativo ? CORES.accentSoft : CORES.bgCard,
                    color: ativo ? CORES.accent : CORES.textoSub,
                  }}
                >{op.label}</button>
              );
            })}
          </div>
        </div>

        {/* Upload */}
        <div style={{ display: "flex", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
          <DropZone
            label={tipoPlanilha === "nfse" ? "Planilha NFS-e" : "Planilha SIEG"}
            sublabel={tipoPlanilha === "nfse" ? "Relatório de notas de serviço (.xlsx)" : "Relatório Cofre SIEG (.xlsx)"}
            onFile={f => { setSiegFile(f); setResultados(null); setErro(""); }}
            fileName={siegFile?.name}
          />
          <DropZone
            label="Export Conta Azul"
            sublabel="Contas a Pagar exportado (.xlsx)"
            onFile={f => { setCaFile(f); setResultados(null); setErro(""); }}
            fileName={caFile?.name}
          />
        </div>

        <div style={{
          background: CORES.bgCard, border: `1px solid ${CORES.borda}`, borderRadius: 10,
          padding: "11px 14px", fontSize: 12.5, color: CORES.textoSub, marginBottom: 16,
          display: "flex", gap: 8, alignItems: "flex-start",
        }}>
          <Info size={15} color={CORES.parcela} style={{ marginTop: 1, flexShrink: 0 }} />
          <span><strong style={{ color: CORES.texto }}>Como exportar do Conta Azul:</strong> Contas a Pagar → Exportar / Relatório → salvar como Excel (.xlsx) e anexar aqui.</span>
        </div>

        {erro && (
          <div style={{ background: CORES.erroSoft, border: `1px solid ${CORES.erro}55`, borderRadius: 10, padding: "11px 14px", fontSize: 13, color: CORES.erro, marginBottom: 16 }}>
            {erro}
          </div>
        )}

        <button
          onClick={conferir}
          disabled={!siegFile || !caFile || processando}
          style={{
            width: "100%", padding: "15px", borderRadius: 12, border: "none",
            background: (!siegFile || !caFile || processando) ? CORES.bgCardHover : `linear-gradient(135deg, ${CORES.accent}, ${CORES.accentDark})`,
            color: (!siegFile || !caFile || processando) ? CORES.textoFraco : "#fff",
            fontWeight: 700, fontSize: 15.5,
            cursor: (!siegFile || !caFile || processando) ? "not-allowed" : "pointer",
            marginBottom: 30, transition: "all 0.18s",
            boxShadow: (!siegFile || !caFile || processando) ? "none" : `0 8px 24px ${CORES.accentSoft}`,
          }}
        >
          {processando ? "Processando…" : "Conferir agora"}
        </button>

      </div>

      {/* Cards de resumo + filtros + tabela: todos full-width */}
      {resultados && (
        <div style={{ padding: "0 20px 70px" }}>
          <div style={{ display: "flex", gap: 12, marginBottom: 18, flexWrap: "nowrap", overflowX: "auto" }}>
            <StatCard label={`Total ${tipoPlanilha === "nfse" ? "NFS-e" : "SIEG"}`} valor={resumo.total} color={CORES.texto} Icon={FileSpreadsheet} />
            <StatCard
              label="OK"
              valor={resumo.ok}
              color={CORES.ok}
              Icon={CheckCircle2}
              tooltip={`${resumo.ok} OK totais\n(${resumo.okPuro} OK, ${resumo.parcela} Parcelas, ${resumo.devolucao} Devoluções)`}
            />
            <StatCard label="Não lançadas" valor={resumo.naoEncontrada} color={CORES.erro} Icon={XCircle}
              tooltip={resumo.naoEncontrada > 0 ? resumo.tooltips.naoEncontrada : undefined} />
            <StatCard label="Divergências" valor={resumo.divergente} color={CORES.alerta} Icon={AlertTriangle}
              tooltip={resumo.divergente > 0 ? resumo.tooltips.divergente : undefined} />
            <StatCard label="Parcelas" valor={resumo.parcela} color={CORES.parcela} Icon={Receipt}
              tooltip={resumo.parcela > 0 ? resumo.tooltips.parcela : undefined} />
            <StatCard label="Duplicadas" valor={resumo.duplicada} color={CORES.duplicada} Icon={Copy}
              tooltip={resumo.duplicada > 0 ? resumo.tooltips.duplicada : undefined} />
            <StatCard label="Devolução" valor={resumo.devolucao} color={CORES.devolucao} Icon={CornerDownLeft}
              tooltip={resumo.devolucao > 0 ? resumo.tooltips.devolucao : undefined} />
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
            <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
              {[
                { val: "todos", label: "Todos" },
                { val: "divergencias", label: "Só problemas" },
                { val: "nao_encontrada", label: "Não lançadas" },
                { val: "divergente", label: "Divergências" },
                { val: "parcela", label: "Parcelas" },
                { val: "duplicada", label: "Duplicadas" },
                { val: "devolucao", label: "Devolução" },
                { val: "ok", label: "OK" },
              ].map(f => (
                <button key={f.val} onClick={() => setFiltro(f.val)} style={{
                  padding: "6px 14px", borderRadius: 999, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                  border: `1px solid ${filtro === f.val ? CORES.accentBorder : CORES.borda}`,
                  background: filtro === f.val ? CORES.accentSoft : "transparent",
                  color: filtro === f.val ? CORES.accent : CORES.textoSub,
                }}>{f.label}</button>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <button onClick={handleSalvarSnap} style={{
                display: "flex", alignItems: "center", gap: 7,
                padding: "9px 16px", borderRadius: 10, border: `1px solid ${CORES.borda}`,
                background: "transparent", color: CORES.textoSub, fontWeight: 700, fontSize: 13, cursor: "pointer",
              }}>
                <Save size={15} /> Salvar versão
              </button>
              <button onClick={abrirVersions} style={{
                display: "flex", alignItems: "center", gap: 7,
                padding: "9px 16px", borderRadius: 10, border: `1px solid ${CORES.borda}`,
                background: "transparent", color: CORES.textoSub, fontWeight: 700, fontSize: 13, cursor: "pointer",
              }}>
                <History size={15} /> Versões
              </button>
              <button
                onClick={() => wbNotasBuffer && gerarPlanilhaPreenchida(wbNotasBuffer, resultados, tipoPlanilha, siegFile?.name)}
                disabled={!wbNotasBuffer}
                title="Retorna a planilha original com Custo, Despesa e Aceite/Recusa preenchidos"
                style={{
                  display: "flex", alignItems: "center", gap: 7,
                  padding: "9px 16px", borderRadius: 10, border: `1px solid ${CORES.bordaForte}`,
                  background: "transparent", color: CORES.textoSub, fontWeight: 700, fontSize: 13,
                  cursor: wbNotasBuffer ? "pointer" : "not-allowed", opacity: wbNotasBuffer ? 1 : 0.4,
                }}
              >
                <FileSpreadsheet size={15} /> Exportar planilha preenchida
              </button>
              <button onClick={() => gerarExcel(resultados, tipoPlanilha)} style={{
                display: "flex", alignItems: "center", gap: 7,
                padding: "9px 16px", borderRadius: 10, border: `1px solid ${CORES.accentBorder}`,
                background: "transparent", color: CORES.accent, fontWeight: 700, fontSize: 13, cursor: "pointer",
              }}>
                <Download size={15} /> Baixar relatório Excel
              </button>
            </div>
          </div>

          <div style={{ background: CORES.bgCard, border: `1px solid ${CORES.borda}`, borderRadius: 14, overflowX: "auto" }}
            onClick={() => dropdownAberto && setDropdownAberto(null)}
          >
            <table style={{ width: "100%", minWidth: 1220, borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: CORES.bgCardAlt }}>
                  {[
                    { label: "Status",    col: "status",      getVal: r => STATUS_CONFIG[statusEfetivo(r)]?.label || statusEfetivo(r) },
                    { label: "Data",      col: "data",         getVal: r => r.nf.dataEmissao ? String(r.nf.dataEmissao).slice(0, 10) : "—" },
                    { label: "Filial",    col: "filial",       getVal: r => r.nf.aba },
                    { label: "Num NF",    col: "numNF",        getVal: r => r.nf.numNF },
                    { label: `Fornecedor ${tipoPlanilha === "nfse" ? "NFS-e" : "SIEG"}`, col: "fornecedor", getVal: r => r.nf.fornecedor },
                    { label: `Valor ${tipoPlanilha === "nfse" ? "NFS-e" : "SIEG"}`, col: "valorSIEG", getVal: r => `R$${r.nf.valor.toFixed(2)}` },
                    { label: "Valor(es) CA",    col: "valorCA",      getVal: r => getValorCAR(r) },
                    { label: "Centro de Custo", col: "centroCusto",  getVal: r => getCentroR(r) },
                    { label: "Custo/Despesa",   col: "custoDespesa", getVal: r => getCustoDespesaR(r) },
                  ].map(({ label, col, getVal }) => {
                    const todosSet = new Set((resultados || []).map(getVal).filter(v => v && v !== "—"));
                    const sortavel = col === "numNF" || col === "fornecedor";
                    return (
                      <ColunaFiltro
                        key={col}
                        label={label}
                        col={col}
                        todos={todosSet}
                        selecionados={filtroCol[col]}
                        onToggle={(c, val, modo) => {
                          if (modo === "todos") {
                            setFiltroCol(prev => ({ ...prev, [c]: null }));
                          } else {
                            setFiltroCol(prev => {
                              const atual = prev[c] ? new Set(prev[c]) : new Set(todosSet);
                              if (atual.has(val)) atual.delete(val); else atual.add(val);
                              return { ...prev, [c]: atual.size === todosSet.size ? null : atual };
                            });
                          }
                        }}
                        onLimpar={c => setFiltroCol(prev => ({ ...prev, [c]: null }))}
                        ordemCol={sortavel ? ordemCol : null}
                        ordemDir={ordemDir}
                        onOrdem={sortavel ? alternarOrdem : null}
                        dropdownAberto={dropdownAberto}
                        onAbrirDropdown={setDropdownAberto}
                      />
                    );
                  })}
                  {["Detalhes", "Observação"].map(h => (
                    <th key={h} style={{
                      padding: "12px 18px", textAlign: "left", fontWeight: 600, fontSize: 11.5,
                      color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em",
                      whiteSpace: "nowrap", borderBottom: `1px solid ${CORES.borda}`,
                    }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtrados.length === 0 ? (
                  <tr><td colSpan={11} style={{ padding: 30, textAlign: "center", color: CORES.textoFraco }}>Nenhum item nesta categoria.</td></tr>
                ) : filtrados.map((r) => <LinhaTabela key={r.id} r={r} onEditar={setEditandoId} />)}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {painelVersions && (
        <div onClick={() => setPainelVersions(false)} style={{
          position: "fixed", inset: 0, background: "rgba(2,5,8,0.7)",
          display: "flex", alignItems: "center", justifyContent: "center",
          zIndex: 100, padding: 18,
        }}>
          <div onClick={e => e.stopPropagation()} style={{
            width: "100%", maxWidth: 520, maxHeight: "85vh", overflow: "auto",
            background: CORES.bgCard, border: `1px solid ${CORES.bordaForte}`,
            borderRadius: 16, padding: 22, boxShadow: "0 24px 60px rgba(0,0,0,0.6)",
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
              <div style={{ fontWeight: 800, fontSize: 16 }}>Versões salvas</div>
              <button onClick={() => setPainelVersions(false)} style={{
                border: "none", background: CORES.bgCardAlt, borderRadius: 8,
                padding: 6, cursor: "pointer", color: CORES.textoSub, display: "flex",
              }}><X size={16} /></button>
            </div>

            {/* Salvar nova versão */}
            <div style={{
              background: CORES.bgCardAlt, borderRadius: 10, padding: "12px 14px", marginBottom: 20,
              border: `1px solid ${CORES.borda}`,
            }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: CORES.textoSub, textTransform: "uppercase", letterSpacing: "0.03em", marginBottom: 8 }}>
                Salvar versão atual
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  value={nomeSnap}
                  onChange={e => setNomeSnap(e.target.value)}
                  placeholder="Nome da versão (opcional)…"
                  onKeyDown={e => e.key === "Enter" && handleSalvarSnap()}
                  style={{
                    flex: 1, borderRadius: 8, padding: "8px 12px",
                    background: CORES.bgCard, border: `1px solid ${CORES.borda}`,
                    color: CORES.texto, fontSize: 13, fontFamily: "inherit", outline: "none",
                  }}
                />
                <button
                  onClick={handleSalvarSnap}
                  disabled={salvandoSnap}
                  style={{
                    display: "flex", alignItems: "center", gap: 6,
                    padding: "8px 14px", borderRadius: 8, border: "none",
                    background: `linear-gradient(135deg, ${CORES.accent}, ${CORES.accentDark})`,
                    color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer",
                    opacity: salvandoSnap ? 0.6 : 1,
                  }}
                ><Save size={14} /> Salvar</button>
              </div>
              {msgSnap && <div style={{ marginTop: 8, fontSize: 12, color: CORES.ok }}>{msgSnap}</div>}
            </div>

            {/* Lista de versões */}
            {snapshots.length === 0 ? (
              <div style={{ textAlign: "center", padding: "24px 0", color: CORES.textoFraco, fontSize: 13 }}>
                Nenhuma versão salva ainda.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {snapshots.map((snap, idx) => {
                  const data = new Date(snap.ts);
                  const fmt = data.toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
                  return (
                    <div key={snap.ts} style={{
                      background: CORES.bgCardAlt, borderRadius: 10, padding: "12px 14px",
                      border: `1px solid ${CORES.borda}`,
                    }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 3 }}>
                            {snap.nome || `Versão ${snapshots.length - idx}`}
                            {idx === 0 && <span style={{ marginLeft: 8, fontSize: 10, fontWeight: 700, color: CORES.ok, background: CORES.okSoft, borderRadius: 4, padding: "2px 6px" }}>MAIS RECENTE</span>}
                          </div>
                          <div style={{ fontSize: 11.5, color: CORES.textoSub, display: "flex", alignItems: "center", gap: 5 }}>
                            <Clock size={11} /> {fmt} · {snap.total} notas · {snap.editados} editadas
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                          <button
                            onClick={() => handleRestaurarSnap(snap)}
                            style={{
                              padding: "6px 11px", borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: "pointer",
                              border: `1px solid ${CORES.accentBorder}`, background: CORES.accentSoft, color: CORES.accent,
                            }}
                          >Restaurar</button>
                          <button
                            onClick={() => handleDeletarSnap(snap)}
                            style={{
                              padding: "6px 11px", borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: "pointer",
                              border: `1px solid ${CORES.borda}`, background: "transparent", color: CORES.textoFraco,
                            }}
                          >Apagar</button>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      )}

      {itemEditando && (
        <PainelEdicao
          r={itemEditando}
          onSalvar={salvarEdicao}
          onRestaurar={restaurarAutomatico}
          onFechar={() => setEditandoId(null)}
        />
      )}
    </div>
  );
}
