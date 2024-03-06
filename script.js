document.addEventListener('DOMContentLoaded', function () {
    var gerarPDFBtn = document.getElementById('gerar_pdf_btn');
    gerarPDFBtn.addEventListener('click', function () {
        var arquivoInput = document.getElementById('arquivo_excel');
        var arquivo = arquivoInput.files[0];
        var leitor = new FileReader();

        leitor.onload = function(e) {
            var workbook = XLSX.read(e.target.result, { type: 'binary' });
            var sheet_name_list = workbook.SheetNames;
            var dados = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], { raw: false, defval: '' });

            // Extrai o nome do arquivo sem a extensão
            var nomeArquivo = arquivo.name.replace(/_/g, ' ').replace(/\(.*?\)/g, '').replace(/\.[^.]*$/, '').trim();

            var docDefinition = {
                content: [],
                styles: {
                    header: {
                        fontSize: 16,
                        bold: true,
                        margin: [0, 0, 0, 8],
                        alignment: 'center'
                    },
                    subheader: {
                        fontSize: 12,
                        bold: true,
                        margin: [0, 8, 0, 4]
                    },
                    defaultStyle: {
                        fontSize: 10,
                        margin: [0, 0, 0, 8]
                    },
                    basicStyle: {
                        fontSize: 10,
                        margin: [0, 0, 0, 0]
                    }
                }
            };

            // Agrupa os dados por ID
            var dadosPorID = {};
            dados.forEach(function(row) {
                var id = row.ID;
                if (!dadosPorID[id]) {
                    dadosPorID[id] = [];
                }
                dadosPorID[id].push(row);
            });

            // Cria uma página separada para cada ID
            Object.keys(dadosPorID).forEach(function(id) {
                // Adiciona os itens adicionais
                var rowData = [];
                rowData.push({ text: nomeArquivo, style: 'header' });
                dadosPorID[id].forEach(function(row) {
                    for (var key in row) {
                        if (row.hasOwnProperty(key) && (key === 'Placa' || key === 'Data' || key === 'Quilometragem' || key === 'Preventiva ou Corretiva')) {
                            rowData.push({ text: `${key}: ${row[key]}`, style: 'subheader', margin: [0, 0, 0, 0] });
                        }
                    }
                });
                docDefinition.content.push(rowData);

                // Adiciona uma linha em branco entre as seções
                docDefinition.content.push({ text: '', margin: [0, 0, 0, 8] });

                // Cria uma tabela para os itens de revisão do ID atual
                var tableBody = [];
                tableBody.push([{ text: 'Item de Revisão', style: 'subheader' }, { text: 'Status', style: 'subheader' }]);
                dadosPorID[id].forEach(function(row) {
                    for (var key in row) {
                        if (row.hasOwnProperty(key) && key !== 'ID' && key !== 'Hora de início' && key !== 'Hora de conclusão' && key !== 'E-mail' && key !== 'Nome' && key !== 'Placa' && key !== 'Data' && key !== 'Quilometragem' && key !== 'Preventiva ou Corretiva') {
                            tableBody.push([key, row[key]]);
                        }
                    }
                });

                // Adiciona a tabela de itens de revisão aos dados do PDF
                docDefinition.content.push({ table: { body: tableBody }, style: 'defaultStyle' });

                // Adiciona os itens adicionais abaixo da tabela
                var finalData = [];
                finalData.push({ text: `ID: ${id}`, style: 'basicStyle' });
                dadosPorID[id].forEach(function(row) {
                        for (var key in row) {
                            if (row.hasOwnProperty(key) && (key === 'Hora de início' || key === 'Hora de conclusão')) {
                                finalData.push({ text: `${key}: ${row[key]}`, style: 'basicStyle' });
                            }
                        }
                    });
                docDefinition.content.push(finalData);
                
                // Adiciona uma quebra de página para o próximo ID
                docDefinition.content.push({ text: '', pageBreak: 'after' });
            });

            pdfMake.createPdf(docDefinition).download('Relatório Checklist.pdf');
        };

        leitor.readAsBinaryString(arquivo);
    });
});