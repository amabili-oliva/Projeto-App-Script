function onFormSubmit() {
  try {
    // Acesso às planilhas
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respostas');
    var aderenciaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aderencia');
    var analiseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Analise');
    
    var lastRow = sheet.getLastRow();
    var emailEnviadoCol = sheet.getLastColumn();
    
    // Verifica a partir da última linha até encontrar uma linha sem "Email Enviado"
    for (var i = lastRow; i > 0; i--) {
      var statusEmail = sheet.getRange(i, emailEnviadoCol).getValue();
      
      // Se a última coluna está vazia, envia o email
      if (!statusEmail) {
        // Obter dados da linha
        var rowData = sheet.getRange(i, 1, 1, emailEnviadoCol - 1).getValues()[0];
        
        // Pega o email e nome e armazena na variável
        var email = rowData[1] ? rowData[1] : 'Não informado';
        var nome = rowData[2] ? rowData[2] : 'Não informado';
        
        // mostra o email e nome
        Logger.log('Email: ' + email);
        Logger.log('Nome: ' + nome);
        
        // Divide as perguntas em  trilha 1 e trilha 2
        var respostasTrilha1 = rowData.slice(3, 9); // Primeiras 6 perguntas
        var respostasTrilha2 = rowData.slice(9, 19); // Últimas 10 perguntas
        
        // somar a pontuação de aderencia das trilhas
        var somaTrilha1 = calcularSomaTrilha(respostasTrilha1, aderenciaSheet, 0, 3); // 4 primeiros perfis
        var somaTrilha2 = calcularSomaTrilha(respostasTrilha2, aderenciaSheet, 4, 16); // 13 últimos perfis
        
        //calcular o percentual de aderencia das trilhas
        var percentuaisTrilha1 = calcularPercentuais(somaTrilha1, respostasTrilha1.length);
        var percentuaisTrilha2 = calcularPercentuais(somaTrilha2, respostasTrilha2.length);
        
        var emailBody = construirEmail(percentuaisTrilha1, percentuaisTrilha2, nome);
        
        // Envia o email
        MailApp.sendEmail(email, "TRILHA Resultados do Questionário de Carreira", emailBody);
        
        // Marca a linha como email enviado
        sheet.getRange(i, emailEnviadoCol).setValue("Email Enviado");

        // Salva os resultados na tabela "Analise"
        salvarRespostasAnalise(analiseSheet, nome, email, percentuaisTrilha1, percentuaisTrilha2);
      }
    }
  } catch (error) {
    Logger.log('Erro ao processar o formulário: ' + error.message);
  }
}

function calcularSomaTrilha(respostas, aderenciaSheet, startProfileIndex, endProfileIndex) {
  var perfis = [];
  
  //  armazenar a soma de cada perfil
  for (var i = startProfileIndex; i <= endProfileIndex; i++) {
    perfis.push({ perfil: aderenciaSheet.getRange(1, i + 3).getValue(), soma: 0 });
  }

  var numPerguntas = respostas.length;
  
  // Percorre todas as respostas e compara com a tabela de aderência
  for (var j = 0; j < numPerguntas; j++) {
    var pergunta = j + 1;
    var resposta = respostas[j];
    
    // Obter as linhas  da tabela de aderência para a pergunta atual
    var range = aderenciaSheet.getRange(2, 1, aderenciaSheet.getLastRow() - 1, aderenciaSheet.getLastColumn());
    var data = range.getValues();
    
    for (var row = 0; row < data.length; row++) {
      if (data[row][0] == pergunta && data[row][1] == resposta) {
        for (var col = startProfileIndex + 2; col <= endProfileIndex + 2; col++) {
          perfis[col - startProfileIndex - 2].soma += data[row][col];
        }
      }
    }
  }
  
  // Ordenar perfis por soma em ordem decrescente
  perfis.sort(function(a, b) {
    return b.soma - a.soma;
  });

  return perfis;
}

function calcularPercentuais(perfis, numPerguntas) {
  return perfis.map(function(perfil) {
    return {
      perfil: perfil.perfil,
      percentual: (perfil.soma / numPerguntas * 100).toFixed(0) + '%'
    };
  });
}

function construirEmail(percentuaisTrilha1, percentuaisTrilha2, nome) {
  var emailBody = 'Olá ' + nome + ',\n\n';
  emailBody += 'Aqui estão os resultados do seu questionário de carreira:\n\n';

  emailBody += 'Relacionamos o seu percentual de aderência para cada trilha mapeada.\n';
  emailBody += '1ª Trilha de Perfil:\n';
  percentuaisTrilha1.forEach(function(perfil) {
    emailBody += '[' + perfil.percentual + ']' + perfil.perfil + '\n';
  });

  emailBody += '\n2ª Trilha de Profissões:\n';
  emailBody += "Aqui está a lista de profissões que mais se encaixa no seu perfil"
  percentuaisTrilha2.forEach(function(perfil) {
    emailBody += '[' + perfil.percentual + ']' + perfil.perfil + '\n';
  });

  emailBody += '\nGostaríamos de saber a sua opinião sobre a análise recebida. Por favor, clique no link abaixo para avaliar os resultados e nos ajudar a melhorar nossos serviços:\n';
  emailBody += '>> [Avalie a Análise de Carreira](https://forms.gle/xJRvnHH72uqm6Xi56)\n\n'

  return emailBody;
}

function salvarRespostasAnalise(analiseSheet, nome, email, percentuaisTrilha1, percentuaisTrilha2) {
  var lastRow = analiseSheet.getLastRow();
  var newRow = lastRow + 1;

  // Adiciona os dados básicos
  analiseSheet.getRange(newRow, 1).setValue(nome);
  analiseSheet.getRange(newRow, 2).setValue(email);

  // Adiciona os percentuais da Trilha 1
  for (var i = 0; i < percentuaisTrilha1.length; i++) {
    analiseSheet.getRange(newRow, 3 + i).setValue(percentuaisTrilha1[i].perfil + ': ' + percentuaisTrilha1[i].percentual);
  }

  // Adiciona os percentuais da Trilha 2
  for (var j = 0; j < percentuaisTrilha2.length; j++) {
    analiseSheet.getRange(newRow, 3 + percentuaisTrilha1.length + j).setValue(percentuaisTrilha2[j].perfil + ': ' + percentuaisTrilha2[j].percentual);
  }
}

