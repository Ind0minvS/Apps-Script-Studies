// Programa - Monitor de Prazos de Tarefas (Gerenciamento de Projetos)

function verificarPrazosTabela() {
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  var dados = planilha.getDataRange().getValues();
  
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  // O loop começa na linha 3 que é onde contem os dados que serão analizados da tabela
  for (var i = 2; i < dados.length; i++) {
    var coluna = dados[i];

    //As colunas começam a ter dados a partir da segunda
    var tarefa = coluna[1];      
    var responsavel = coluna[2];   
    var dataVencimento = new Date(coluna[3]); 
    var estado = coluna[4];      
    
    dataVencimento.setHours(0, 0, 0, 0);

    if (estado !== "Concluído") {       
      // te9154281@gmail.com é um email teste, em caso de testes, só trocar o email

      // Verifica se vence hoje
      if (dataVencimento.getTime() === hoje.getTime()) {
    
        MailApp.sendEmail({ to: "te9154281@gmail.com", subject: "Atenção: Tarefa vence hoje!", 
          body: `Olá, a tarefa (${tarefa}) de ${responsavel} vence hoje.`});
        }
        
        // Atrasado (Data menor que hoje)
        else if (dataVencimento < hoje) {

          MailApp.sendEmail({ to: "te9154281@gmail.com", subject: "URGENTE: Tarefa em atraso", 
          body: `Olá, a tarefa (${tarefa}) de ${responsavel} estava prevista para ${dataVencimento.toLocaleDateString()} e já passou do prazo.`});
        }

         Logger.log("Email enviado ");
      }
  }
}

function enviarEmail(destinatario, assunto, mensagem) {
    MailApp.sendEmail({
      to: destinatario,
      subject: assunto,
      body: mensagem
    });
    Logger.log("Email enviado para: " + destinatario);
}
