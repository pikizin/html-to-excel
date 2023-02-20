const express = require('express');
const ExcelJS = require('exceljs');
const app = express();

// Configuração da planilha do Excel
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Form Data');

// Configuração do endpoint para receber os dados do formulário
app.post('/form-data', (req, res) => {
  const formData = req.body;

  // Escreve os dados na planilha
  worksheet.columns = [
    { header: 'Nome', key: 'nome' },
    { header: 'Idade', key: 'idade' },
    { header: 'Email', key: 'email' },
    { header: 'Sexo', key: 'sexo' },
    { header: 'Mensagem', key: 'mensagem' }
  ];

  worksheet.addRow({
    nome: formData.nome,
    idade: formData.idade,
    email: formData.email,
    sexo: formData.sexo,
    mensagem: formData.mensagem
  });

  // Grava o arquivo do Excel
  workbook.xlsx.writeFile('dados.xlsx').then(() => {
    console.log('Arquivo salvo com sucesso');
  }).catch((error) => {
    console.log('Erro ao salvar arquivo:', error);
  });

  // Envia uma resposta de sucesso ao cliente
  res.send('Dados do formulário recebidos com sucesso');
});

// Inicia o servidor
app.listen(3000, () => {
  console.log('Servidor iniciado na porta 3000');
});
