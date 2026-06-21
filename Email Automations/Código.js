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

    // Definição do novo destinatário do teste
    var destinatarioTeste = "testgmail939@gmail.com";
    // Definição do seu alias configurado
    var aliasRemetente = "testgmail939+automaticalert@gmail.com";

    if (estado !== "Concluído") {       

      // Verifica se vence hoje
      if (dataVencimento.getTime() === hoje.getTime()) {
        
        // Alterado para GmailApp com parâmetro "from" avançado
        GmailApp.sendEmail(destinatarioTeste, "Atenção: Tarefa vence hoje!", `Olá, a tarefa (${tarefa}) de ${responsavel} vence hoje.`, {
          from: aliasRemetente
        });
        Logger.log("Email enviado: Vence Hoje");
      }
        
      // Atrasado (Data menor que hoje)
      else if (dataVencimento < hoje) {

        // Alterado para GmailApp com parâmetro "from" avançado
        GmailApp.sendEmail(destinatarioTeste, "URGENTE: Tarefa em atraso", `Olá, a tarefa (${tarefa}) de ${responsavel} estava prevista para ${dataVencimento.toLocaleDateString()} e já passou do prazo.`, {
          from: aliasRemetente
        });
        Logger.log("Email enviado: Atrasado");
      }
    }
  }
}

// Função auxiliar atualizada para suportar o alias via GmailApp
function enviarEmail(destinatario, assunto, mensagem) {
    var aliasRemetente = "testgmail939+automaticalert@gmail.com";
    
    GmailApp.sendEmail(destinatario, assunto, mensagem, {
      from: aliasRemetente
    });
    
    Logger.log("Email enviado para: " + destinatario);
}
