// Constantes para os cabeçalhos das colunas
const HEADER_FIELDS = [
    "PEDIDO", "ITEM", "CÓDIGO", "DESCRITIVO", "QUANTIDADE", "UN", 
    "LOTE", "DATA LAN", "NOTA FISCAL", "VALOR", "CC", "ORDEM", 
    "PEP", "CRIADO POR", "ACOMPANHAMENTO" 
];

const headerMapping = {
    "PEDIDO": "pedido", "ITEM": "item", "CÓDIGO": "código", 
    "DESCRITIVO": "descritivo", "QUANTIDADE": "quantidade", "UN": "un", 
    "LOTE": "lote", "DATA LAN": "data_lan", "NOTA FISCAL": "nota_fiscal", 
    "VALOR": "valor", "CC": "cc", "ORDEM": "ordem", 
    "PEP": "pep", "CRIADO POR": "criado_por", "ACOMPANHAMENTO": "acompanhamento" 
};

const columnKeys = HEADER_FIELDS.map(h => h.replace(/\s/g, '_').toLowerCase());

// =========================================================================
// FUNÇÕES GERAIS E DE UTILIDADE
// =========================================================================

/**
 * Converte serial date do Excel para formato dd/mm/aaaa
 */
function excelSerialDateToJSDate(serial) {
    // 25569 é o número de dias entre 1 Jan 1900 (Excel) e 1 Jan 1970 (Unix epoch)
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400000; // Multiplica por 1000 para ms
    
    // Cria uma data UTC (sem ajuste de fuso)
    const date_info = new Date(utc_value);

    // Ajuste: A hora do fuso horário pode bagunçar o dia.
    // Usaremos as funções UTC para garantir que o dia seja mantido.
    const day = date_info.getUTCDate().toString().padStart(2, '0');
    const month = (date_info.getUTCMonth() + 1).toString().padStart(2, '0');
    const year = date_info.getUTCFullYear();

    return `${day}/${month}/${year}`;
}

/**
 * Adiciona o item selecionado ao array de itens no sessionStorage e navega.
 */
function enviarItem(item) {
    let itensSelecionados = JSON.parse(sessionStorage.getItem('itensDetalhe')) || [];
    
    // Cria uma chave única para o item (necessário para a exclusão)
    const uniqueKey = `${item.pedido}-${item.item}-${item.código}`;
    
    if (!itensSelecionados.some(i => i.__uniqueKey === uniqueKey)) {
        const itemComChave = {...item, __uniqueKey: uniqueKey};
        itensSelecionados.push(itemComChave);
        alert(`Item ${item.código} - Pedido ${item.pedido} adicionado à visualização de Detalhe.`);
    } else {
        alert(`O item ${item.código} - Pedido ${item.pedido} já está na lista de Detalhes.`);
    }

    sessionStorage.setItem('itensDetalhe', JSON.stringify(itensSelecionados));
    window.location.href = 'detalhe.html';
}

/**
 * Função global para exportar dados para Excel.
 */
function exportarParaExcel(dataArray, fileName) {
    if (dataArray.length === 0) {
        alert("Não há dados para exportar.");
        return;
    }
    
    const worksheet = XLSX.utils.json_to_sheet(dataArray, {header: columnKeys});
    const workbook = XLSX.utils.book_new();
    
    XLSX.utils.sheet_add_aoa(worksheet, [HEADER_FIELDS], { origin: "A1" });

    XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");
    
    XLSX.writeFile(workbook, fileName + ".xlsx");
}

/**
 * FUNÇÃO DE EXCLUSÃO PERMANENTE (usada na Consulta).
 */
function excluirLinha(itemParaExcluir, rowElement = null) {
    if (!confirm('Tem certeza que deseja EXCLUIR PERMANENTEMENTE esta linha?')) {
        return false;
    }
    
    dadosCarregados = JSON.parse(localStorage.getItem('dadosExcel')) || [];

    const indexOriginal = dadosCarregados.findIndex(item => 
        item.pedido === itemParaExcluir.pedido && 
        item.item === itemParaExcluir.item && 
        item.código === itemParaExcluir.código
    );
    
    if (indexOriginal !== -1) {
        dadosCarregados.splice(indexOriginal, 1);
        
        localStorage.setItem('dadosExcel', JSON.stringify(dadosCarregados));

        if (rowElement) {
             rowElement.remove();
             aplicarFiltros(); 
        }

        // Tenta remover também do Detalhe se estiver lá
        let itensDetalhe = JSON.parse(sessionStorage.getItem('itensDetalhe')) || [];
        const uniqueKey = `${itemParaExcluir.pedido}-${itemParaExcluir.item}-${itemParaExcluir.código}`;
        itensDetalhe = itensDetalhe.filter(item => item.__uniqueKey !== uniqueKey);
        sessionStorage.setItem('itensDetalhe', JSON.stringify(itensDetalhe));

        return true; 
    } else {
        alert("Erro: Item não encontrado no banco de dados para exclusão.");
        return false;
    }
}


// =========================================================================
// LÓGICA DA PÁGINA DE IMPORTAÇÃO (index.html)
// =========================================================================

/**
 * Função principal para importar os dados do arquivo Excel.
 */
function importarDados() {
    const file = document.getElementById('arquivoExcel').files[0];
    const mensagem = document.getElementById('mensagem');
    mensagem.textContent = '';
    
    if (!file) {
        mensagem.textContent = 'Por favor, selecione um arquivo Excel.';
        mensagem.style.color = 'red';
        return;
    }

    // Ponto de Início: Mostra que a leitura começou
    mensagem.textContent = 'Lendo arquivo...';
    mensagem.style.color = 'blue';

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            // Configuração para forçar a leitura de datas como números (seriais)
            const workbook = XLSX.read(data, {type: 'array', cellDates: false, raw: true}); 
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Lê para JSON com o cabeçalho como Array, para fazer o mapeamento manualmente
            const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            if (json.length < 2) {
                mensagem.textContent = 'A planilha está vazia ou não tem cabeçalhos.';
                mensagem.style.color = 'red';
                return;
            }

            const rawHeaders = json[0] || [];
            // Normaliza os cabeçalhos do Excel
            const normalizedHeaders = rawHeaders.map(h => (h ? String(h).toUpperCase().trim() : ''));

            const formattedData = [];

            for (let i = 1; i < json.length; i++) {
                const row = {};
                let isValidRow = false;
                
                HEADER_FIELDS.forEach((expectedHeader, index) => {
                    const colIndex = normalizedHeaders.indexOf(expectedHeader);
                    const key = columnKeys[index];
                    let value = '';

                    if (colIndex !== -1 && json[i][colIndex] !== undefined) {
                        value = json[i][colIndex];
                        
                        if (expectedHeader === 'DATA LAN') {
                            // Converte serial number para data formatada
                            if (typeof value === 'number') {
                                value = excelSerialDateToJSDate(value);
                            } else {
                                value = String(value || '').trim();
                            }
                        } else {
                            value = String(value || '').trim();
                        }
                    } 
                    
                    row[key] = value; 

                    // Verifica se a linha tem pelo menos um valor preenchido (não vazio)
                    if (value !== '') {
                        isValidRow = true;
                    }
                });

                if (isValidRow) {
                    formattedData.push(row);
                }
            }
            
            if (formattedData.length === 0) {
                 mensagem.textContent = 'Não foi possível encontrar dados válidos na planilha.';
                 mensagem.style.color = 'red';
                 return;
            }

            localStorage.setItem('dadosExcel', JSON.stringify(formattedData));
            mensagem.textContent = `Importação concluída! ${formattedData.length} registros carregados. Redirecionando...`;
            mensagem.style.color = 'green';

            setTimeout(() => {
                window.location.href = 'consulta.html';
            }, 1500);
            
        } catch (error) {
            console.error("Erro ao processar o arquivo Excel:", error);
            mensagem.textContent = `Erro ao processar o arquivo: ${error.message}. Verifique o formato e o cabeçalho.`;
            mensagem.style.color = 'red';
        }
    };

    reader.onerror = function() {
        mensagem.textContent = 'Erro de leitura de arquivo (FileReader).';
        mensagem.style.color = 'red';
    };

    reader.readAsArrayBuffer(file);
}


// =========================================================================
// LÓGICA DA PÁGINA DE CONSULTA (consulta.html)
// =========================================================================

let dadosCarregados = [];
let dadosAtuais = []; 

function inicializarConsulta() {
    dadosCarregados = JSON.parse(localStorage.getItem('dadosExcel')) || [];
    dadosAtuais = [...dadosCarregados];

    const tabelaBody = document.querySelector('#tabelaDados tbody');
    const statusElement = document.getElementById('status');
    
    if (dadosCarregados.length === 0) {
        statusElement.textContent = 'Nenhum dado encontrado. Retorne para a página de Importação.';
        tabelaBody.innerHTML = '<tr><td colspan="16">Nenhum dado importado. <a href="index.html">Voltar para Importação</a></td></tr>';
    } else {
        renderizarTabela(dadosAtuais);
    }
}

function renderizarTabela(dados) {
    const tabelaBody = document.querySelector('#tabelaDados tbody');
    const statusElement = document.getElementById('status');
    tabelaBody.innerHTML = '';
    
    if (dados.length === 0) {
        tabelaBody.innerHTML = '<tr><td colspan="16">Nenhum registro corresponde aos filtros aplicados.</td></tr>';
        statusElement.textContent = `0 registros exibidos (total: ${dadosCarregados.length}).`;
        return;
    }

    dados.forEach((item) => {
        const row = tabelaBody.insertRow();
        
        columnKeys.forEach(key => {
            const cell = row.insertCell();
            cell.textContent = item[key] || '';
        });

        const actionCell = row.insertCell();
        
        // Botão para ENVIAR (para Detalhe)
        const btnEnviar = document.createElement('button');
        btnEnviar.textContent = 'Enviar';
        btnEnviar.className = 'btn-enviar';
        btnEnviar.onclick = () => enviarItem(item); 
        actionCell.appendChild(btnEnviar);
        
        // Botão para EXCLUIR PERMANENTEMENTE
        const btnExcluir = document.createElement('button');
        btnExcluir.textContent = 'Excluir';
        btnExcluir.className = 'btn-excluir';
        btnExcluir.onclick = () => excluirLinha(item, row); 
        actionCell.appendChild(btnExcluir);
    });
    
    statusElement.textContent = `${dados.length} registros exibidos (total: ${dadosCarregados.length}).`;
}

function aplicarFiltros() {
    const filtroCodigo = document.getElementById('filtroCodigo').value.toUpperCase().trim();
    const filtroPedido = document.getElementById('filtroPedido').value.toUpperCase().trim();
    const filtroNota = document.getElementById('filtroNota').value.toUpperCase().trim();
    const filtroData = document.getElementById('filtroData').value.toUpperCase().trim();

    dadosAtuais = dadosCarregados.filter(item => {
        const codigo = (item.código || '').toUpperCase();
        const pedido = (item.pedido || '').toUpperCase();
        const notaFiscal = (item.nota_fiscal || '').toUpperCase();
        const dataLan = (item.data_lan || '').toUpperCase();

        let match = true;
        
        if (filtroCodigo && !codigo.includes(filtroCodigo)) {
            match = false;
        }
        if (filtroPedido && !pedido.includes(filtroPedido)) {
            match = false;
        }
        if (filtroNota && !notaFiscal.includes(filtroNota)) {
            match = false;
        }
        if (filtroData && !dataLan.includes(filtroData)) {
            match = false;
        }

        return match;
    });

    renderizarTabela(dadosAtuais);
}

function limparFiltros() {
    document.getElementById('filtroCodigo').value = '';
    document.getElementById('filtroPedido').value = '';
    document.getElementById('filtroNota').value = '';
    document.getElementById('filtroData').value = '';
    
    dadosAtuais = [...dadosCarregados];
    renderizarTabela(dadosAtuais);
}

function irParaDetalhe() {
    if (sessionStorage.getItem('itensDetalhe')) {
        window.location.href = 'detalhe.html';
    } else {
        alert("Nenhum item foi enviado. Por favor, use o botão 'Enviar' na tabela para selecionar um item.");
    }
}


// =========================================================================
// LÓGICA DA PÁGINA DE DETALHES (detalhe.html)
// =========================================================================

let itensAtuaisDetalhe = [];

function inicializarDetalhe() {
    const itensJson = sessionStorage.getItem('itensDetalhe');
    itensAtuaisDetalhe = JSON.parse(itensJson) || [];

    const tabelaDetalhesContainer = document.getElementById('tabelaDetalhesContainer');
    const detalhesMensagem = document.getElementById('detalhesMensagem');

    if (itensAtuaisDetalhe.length === 0) {
        tabelaDetalhesContainer.innerHTML = '<h2>Nenhum item selecionado.</h2>';
        detalhesMensagem.textContent = 'Por favor, volte para a página de Consulta e selecione itens para enviar.';
        
        document.getElementById('btnExportarItem').disabled = true;
        document.getElementById('btnLimparVisualizacao').disabled = true;
        return;
    }
    
    document.getElementById('btnExportarItem').disabled = false;
    document.getElementById('btnLimparVisualizacao').disabled = false;
    detalhesMensagem.textContent = `${itensAtuaisDetalhe.length} item(s) em visualização.`;
    
    renderizarDetalheComoTabela(itensAtuaisDetalhe);
}

/**
 * Renderiza o array de itens detalhados em formato de tabela.
 */
function renderizarDetalheComoTabela(itens) {
    const tabelaBody = document.querySelector('#tabelaDetalhes tbody');
    tabelaBody.innerHTML = ''; 
    
    itens.forEach((item) => {
        const row = tabelaBody.insertRow();
        
        columnKeys.forEach(key => {
            const cell = row.insertCell();
            cell.textContent = item[key] || '';
        });

        const actionCell = row.insertCell();
        const btnApagar = document.createElement('button');
        btnApagar.textContent = 'Apagar';
        btnApagar.className = 'btn-apagar-visualizacao';
        btnApagar.onclick = () => apagarLinhaDetalhe(item.__uniqueKey); 
        actionCell.appendChild(btnApagar);
    });
}


function exportarItemExcel() {
    if (itensAtuaisDetalhe.length === 0) {
        alert("Nenhum item para exportar.");
        return;
    }
    const fileName = `Detalhe_${new Date().toLocaleDateString('pt-BR').replace(/\//g, '-')}`;
    
    const dadosParaExportar = itensAtuaisDetalhe.map(item => {
        const { __uniqueKey, ...rest } = item;
        return rest;
    });

    exportarParaExcel(dadosParaExportar, fileName);
}


/**
 * NOVO: Remove apenas a linha específica do array de visualização (sessionStorage).
 */
function apagarLinhaDetalhe(uniqueKey) {
    if (!confirm('Deseja remover esta linha da visualização de Detalhe? Ela CONTINUARÁ na página de Consulta.')) {
        return;
    }
    
    const indexParaRemover = itensAtuaisDetalhe.findIndex(item => item.__uniqueKey === uniqueKey);
    
    if (indexParaRemover !== -1) {
        itensAtuaisDetalhe.splice(indexParaRemover, 1);
        
        sessionStorage.setItem('itensDetalhe', JSON.stringify(itensAtuaisDetalhe));
        
        inicializarDetalhe(); // Re-renderiza a página
    } else {
        alert("Erro: Item não encontrado na lista de visualização.");
    }
}

/**
 * FUNÇÃO DE LIMPEZA GERAL: Remove TUDO da visualização de Detalhe (sessionStorage).
 */
function limparItemDetalheVisualizacao() {
    if (itensAtuaisDetalhe.length === 0) {
        alert("Nenhum item selecionado para limpar.");
        return;
    }

    if (!confirm('Deseja limpar TODOS os itens da visualização de Detalhe? Eles CONTINUARÃO na página de Consulta.')) {
        return;
    }
    
    sessionStorage.removeItem('itensDetalhe');
    itensAtuaisDetalhe = []; 

    // Limpa a tela e desativa os botões
    document.getElementById('tabelaDetalhesContainer').innerHTML = '<h2>Visualização limpa.</h2>';
    document.getElementById('detalhesMensagem').textContent = 'Selecione novos itens na Consulta.';
    document.getElementById('btnExportarItem').disabled = true;
    document.getElementById('btnLimparVisualizacao').disabled = true;
}