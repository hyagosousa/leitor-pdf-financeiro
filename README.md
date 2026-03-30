<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Analisador Completo</title>

<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<style>
body { font-family: Arial; background: #111; color: white; text-align: center; padding: 20px; }
h1 { color: #00ffcc; }

input, button { margin: 10px; padding: 10px; }

button {
  background: #00aa88;
  color: white;
  border: none;
  cursor: pointer;
}

.box {
  display: flex;
  justify-content: space-around;
  margin-top: 20px;
}

.coluna {
  width: 45%;
  padding: 15px;
  border-radius: 10px;
}

.positivo { background: #063; }
.negativo { background: #600; }

li { margin: 8px 0; text-align: left; }

/* tabela */
table { width: 95%; margin: 20px auto; border-collapse: collapse; }
th, td { border: 1px solid #fff; padding: 8px; text-align: right; }
th { background: #00aa88; }
td:first-child, td:nth-child(2) { text-align: left; }
td { background: #063; }
.maior { background: #0044ff; }
.menor { background: #880000; }

</style>
</head>

<body>

<h1>📄 Analisador Completo de Balancete</h1>

<input type="file" id="pdfInput" multiple accept="application/pdf">

<br>
<button onclick="baixarNegativosZIP()">⬇️ Negativos (ZIP)</button>
<button onclick="baixarPositivosZIP()">⬇️ Positivos (ZIP)</button>
<button onclick="exportarExcel()">📊 Exportar Excel</button>

<!-- LISTAS -->
<div class="box">
  <div class="coluna positivo">
    <h2>✅ Positivos</h2>
    <ul id="positivos"></ul>
  </div>

  <div class="coluna negativo">
    <h2>❌ Negativos</h2>
    <ul id="negativos"></ul>
  </div>
</div>

<!-- TABELA -->
<table id="tabela">
<thead>
<tr>
<th>Arquivo</th>
<th>Empresa</th>
<th>Resultado</th>
<th>Produtos</th>
<th>Mercadoria</th>
<th>Serviços</th>
<th>Simples</th>
<th>Serv + Simples</th>
<th>Comp</th>
<th>Total</th>
<th>Comp</th>
</tr>
</thead>
<tbody id="tabelaResumo"></tbody>
</table>

<script>

pdfjsLib.GlobalWorkerOptions.workerSrc =
"https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";

const input = document.getElementById("pdfInput");

let arquivosNegativos = [];
let arquivosPositivos = [];

input.addEventListener("change", async (event) => {

  document.getElementById("positivos").innerHTML = "";
  document.getElementById("negativos").innerHTML = "";
  document.getElementById("tabelaResumo").innerHTML = "";

  arquivosNegativos = [];
  arquivosPositivos = [];

  const arquivos = event.target.files;

  for (let file of arquivos) {

    const textoSimples = await lerPDFSimples(file);
    analisarTexto(textoSimples, file.name, file);

    const textoTabela = await lerPDFTabela(file);
    extrairInformacoes(textoTabela, file.name);
  }
});

/* ================= PDF SIMPLES (NEG/POS) ================= */
async function lerPDFSimples(file) {
  const reader = new FileReader();

  return new Promise((resolve) => {
    reader.onload = async function () {
      const typedarray = new Uint8Array(this.result);
      const pdf = await pdfjsLib.getDocument(typedarray).promise;

      let texto = "";

      for (let i = 1; i <= pdf.numPages; i++) {
        const pagina = await pdf.getPage(i);
        const conteudo = await pagina.getTextContent();

        conteudo.items.forEach(item => {
          texto += item.str + " ";
        });

        texto += "\n";
      }

      resolve(texto.toLowerCase());
    };

    reader.readAsArrayBuffer(file);
  });
}

function analisarTexto(texto, nomeArquivo, fileAtual) {

  texto = texto.replace(/\s+/g, " ");

  const index = texto.indexOf("resultado do período");

  if (index !== -1) {

    const depois = texto.substring(index, index + 300);
    const valores = depois.match(/\(?\s*\d{1,3}(?:\.\d{3})*,\d{2}\s*\)?/g);

    if (valores && valores.length >= 4) {

      const saldoBruto = valores[3];
      const saldo = saldoBruto.replace(/\s+/g, "");

      const posSaldo = depois.indexOf(saldoBruto);
      const contexto = depois.substring(
        Math.max(0, posSaldo - 15),
        posSaldo + saldoBruto.length + 15
      );

      const ehNegativo =
        saldo.startsWith("(") ||
        saldo.endsWith(")") ||
        contexto.includes("(");

      if (ehNegativo) {
        adicionarLista("negativos", nomeArquivo, saldo);
        arquivosNegativos.push(fileAtual);
      } else {
        adicionarLista("positivos", nomeArquivo, saldo);
        arquivosPositivos.push(fileAtual);
      }
    }
  }
}

function adicionarLista(tipo, nome, saldo) {
  const ul = document.getElementById(tipo);
  const li = document.createElement("li");
  li.textContent = `${nome} → Saldo: ${saldo}`;
  ul.appendChild(li);
}

/* ================= PDF TABELA ================= */

async function lerPDFTabela(file) {
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

  return "Não identificado";
}

function converterParaNumero(valor) {
  if (!valor || valor === "-") return 0;

  return parseFloat(
    valor.replace(/\./g, "").replace(",", ".").replace("(", "-").replace(")", "")
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

  const calcProdutos = vProdutos * 0.08;
  const calcMercadoria = vMercadoria * 0.08;
  const calcServicos = vServicos * 0.32;
  const calcSimples = vSimples * 0.05;

  const totalServicos = calcServicos + calcSimples;
  const totalGeral = calcProdutos + calcMercadoria + calcSimples;

  const comparacao1 = totalServicos > vResultado ? "MAIOR" : "MENOR";
  const comparacao2 = totalGeral > vResultado ? "MAIOR" : "MENOR";

  const classe1 = comparacao1 === "MAIOR" ? "maior" : "menor";
  const classe2 = comparacao2 === "MAIOR" ? "maior" : "menor";

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

/* ================= DOWNLOAD ================= */

async function baixarNegativosZIP() {
  if (arquivosNegativos.length === 0) {
    alert("Nenhum PDF negativo encontrado.");
    return;
  }

  const zip = new JSZip();
  arquivosNegativos.forEach(file => zip.file(file.name, file));

  const conteudo = await zip.generateAsync({ type: "blob" });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(conteudo);
  link.download = "pdfs_negativos.zip";
  link.click();
}

async function baixarPositivosZIP() {
  if (arquivosPositivos.length === 0) {
    alert("Nenhum PDF positivo encontrado.");
    return;
  }

  const zip = new JSZip();
  arquivosPositivos.forEach(file => zip.file(file.name, file));

  const conteudo = await zip.generateAsync({ type: "blob" });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(conteudo);
  link.download = "pdfs_positivos.zip";
  link.click();
}

function exportarExcel() {
  const tabela = document.getElementById("tabela");
  const wb = XLSX.utils.table_to_book(tabela, { sheet: "Resumo" });
  XLSX.writeFile(wb, "Resumo_PDFs.xlsx");
}

</script>

</body>
</html>

</script>

</body>
</html>>
