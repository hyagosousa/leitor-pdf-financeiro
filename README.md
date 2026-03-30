<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>BMD - Sistema Contábil</title>

<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<style>
body { 
  font-family: Arial; 
  background: #111; 
  color: white; 
  text-align: center; 
  padding: 20px; 
}

/* 🔥 LOGO CONTÁBIL */
.logo-contabil {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 15px;
  margin-bottom: 20px;
}

.logo-icon {
  font-size: 35px;
  background: #00aa88;
  padding: 12px;
  border-radius: 10px;
  box-shadow: 0 0 10px rgba(0,170,136,0.5);
}

.logo-texto {
  text-align: left;
}

.logo-titulo {
  font-size: 32px;
  font-weight: bold;
  letter-spacing: 3px;
  color: #00ffcc;
}

.logo-sub {
  font-size: 12px;
  color: #aaa;
  letter-spacing: 2px;
}

h1 { color: #00ffcc; }

input, button { margin: 10px; padding: 10px; }

table { width: 95%; margin: 20px auto; border-collapse: collapse; }

th, td { border: 1px solid #fff; padding: 8px; text-align: right; }

th { background: #00aa88; }

td:first-child, td:nth-child(2) { text-align: left; }

td { background: #063; }

.maior { background: #0044ff; }
.menor { background: #880000; }

button { 
  background: #00aa88; 
  color: white; 
  border: none; 
  cursor: pointer; 
}

button:hover { background: #008866; }

</style>
</head>

<body>

<!-- 🔥 LOGO PROFISSIONAL -->
<div class="logo-contabil">
  <div class="logo-icon">📊</div>
  <div class="logo-texto">
    <div class="logo-titulo">BMD</div>
    <div class="logo-sub">SISTEMA CONTÁBIL</div>
  </div>
</div>

<h1>Resumo de PDFs Positivos</h1>

<input type="file" id="pdfInput" multiple accept="application/pdf">
<br>
<button onclick="exportarExcel()">Exportar para Excel</button>

<table id="tabela">
<thead>
<tr>
<th>Arquivo</th>
<th>Empresa</th>
<th>Resultado (2600)</th>
<th>Produtos (2603)</th>
<th>Mercadoria (2652)</th>
<th>Serviços (2700)</th>
<th>Simples (2831)</th>
<th>Serviços + Simples</th>
<th>Comparação</th>
<th>Prod + Merc + Simples</th>
<th>Comparação</th>
</tr>
</thead>
<tbody id="tabelaResumo"></tbody>
</table>

<script>

pdfjsLib.GlobalWorkerOptions.workerSrc =
"https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";

const input = document.getElementById("pdfInput");

input.addEventListener("change", async (event) => {
document.getElementById("tabelaResumo").innerHTML = "";
const arquivos = event.target.files;

for (let file of arquivos) {
const texto = await lerPDF(file);
extrairInformacoes(texto, file.name);
}
});

// 🔥 LEITURA ORGANIZADA
async function lerPDF(file) {
const reader = new FileReader();

return new Promise((resolve) => {
reader.onload = async function () {
const typedarray = new Uint8Array(this.result);
const pdf = await pdfjsLib.getDocument(typedarray).promise;

let linhas = [];

for (let i = 1; i <= pdf.numPages; i++) {
const pagina = await pdf.getPage(i);
const conteudo = await pagina.getTextContent();

const items = conteudo.items.sort((a, b) => b.transform[5] - a.transform[5]);

let linhaAtual = "";
let yAnterior = null;

items.forEach(item => {
const y = item.transform[5];

if (yAnterior !== null && Math.abs(y - yAnterior) > 5) {
linhas.push(linhaAtual);
linhaAtual = "";
}

linhaAtual += item.str + " ";
yAnterior = y;
});

if (linhaAtual) linhas.push(linhaAtual);
}

resolve(linhas.join("\n").toLowerCase());
};

reader.readAsArrayBuffer(file);
});
}

// 🏢 NOME EMPRESA
function pegarNomeEmpresa(texto) {
const linhas = texto.split("\n");

for (let linha of linhas) {
linha = linha.trim();

if (
linha.length > 5 &&
!linha.includes("página") &&
!linha.includes("balancete") &&
!linha.includes("contábil")
) {
return linha.toUpperCase();
}
}

return "NÃO IDENTIFICADO";
}

// 🔢 CONVERSÃO
function converterParaNumero(valor) {
if (!valor || valor === "-") return 0;

return parseFloat(
valor
.replace(/\./g, "")
.replace(",", ".")
.replace("(", "-")
.replace(")", "")
);
}

function limpar(valor) {
if (!valor || valor === "-") return "-";
return valor.replace(/[()]/g, "");
}

function buscarLinha(texto, codigo) {
const linhas = texto.split("\n");

for (let linha of linhas) {
if (linha.trim().startsWith(codigo + " ")) {
return linha;
}
}
return "";
}

function pegarValor(linha) {
if (!linha) return "-";

const numeros = linha.match(/\(?\d{1,3}(?:\.\d{3})*,\d{2}\)?/g);
if (!numeros) return "-";

return numeros[numeros.length - 1];
}

// 🔥 PROCESSAMENTO
function extrairInformacoes(texto, nomeArquivo) {

const nomeEmpresa = pegarNomeEmpresa(texto);

const resultado = pegarValor(buscarLinha(texto, "2600"));
const produtos = pegarValor(buscarLinha(texto, "2603"));
const mercadoria = pegarValor(buscarLinha(texto, "2652"));
const servicos = pegarValor(buscarLinha(texto, "2700"));
const simples = pegarValor(buscarLinha(texto, "2831"));

const vResultado = converterParaNumero(resultado);
const vProdutos = converterParaNumero(produtos);
const vMercadoria = converterParaNumero(mercadoria);
const vServicos = converterParaNumero(servicos);
const vSimples = converterParaNumero(simples);

// cálculos
const calcProdutos = vProdutos * 0.08;
const calcMercadoria = vMercadoria * 0.08;
const calcServicos = vServicos * 0.32;
const calcSimples = vSimples * 0.05;

const totalServicos = calcServicos + calcSimples;
const totalGeral = calcProdutos + calcMercadoria + calcSimples;

// comparação
const comparacao1 = totalServicos > vResultado ? "MAIOR" : "MENOR";
const comparacao2 = totalGeral > vResultado ? "MAIOR" : "MENOR";

const classe1 = comparacao1 === "MAIOR" ? "maior" : "menor";
const classe2 = comparacao2 === "MAIOR" ? "maior" : "menor";

// tabela
const tbody = document.getElementById("tabelaResumo");
const tr = document.createElement("tr");

tr.innerHTML = `
<td>${nomeArquivo}</td>
<td>${nomeEmpresa}</td>
<td>${limpar(resultado)}</td>
<td>${limpar(produtos)}</td>
<td>${limpar(mercadoria)}</td>
<td>${limpar(servicos)}</td>
<td>${limpar(simples)}</td>
<td>${totalServicos.toFixed(2)}</td>
<td class="${classe1}">${comparacao1}</td>
<td>${totalGeral.toFixed(2)}</td>
<td class="${classe2}">${comparacao2}</td>
`;

tbody.appendChild(tr);
}

// 📊 EXPORTAR EXCEL
function exportarExcel() {
const tabela = document.getElementById("tabela");
const wb = XLSX.utils.table_to_book(tabela, { sheet: "Resumo" });
XLSX.writeFile(wb, "Resumo_PDFs.xlsx");
}

</script>

</body>
</html>

