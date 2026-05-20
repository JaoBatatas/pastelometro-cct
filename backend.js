function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("Dados");
  var sheetBloqueados = ss.getSheetByName("Bloqueados");
  
  // Cria automaticamente a aba de Bloqueados se o usuário esquecer de criar
  if (!sheetBloqueados) {
    sheetBloqueados = ss.insertSheet("Bloqueados");
    sheetBloqueados.appendRow(["IP"]);
  }
  
  // Captura o IP do usuário enviado pelo frontend
  var userIp = e.parameter.ip || "0.0.0.0";
  
  // Verifica se o IP está na lista de bloqueados (evita erro caso a tabela esteja vazia)
  var bloqueado = false;
  var lastRowBloqueados = sheetBloqueados.getLastRow();
  if (lastRowBloqueados > 1) {
    var listaBloqueados = sheetBloqueados.getRange(2, 1, lastRowBloqueados - 1, 1).getValues();
    bloqueado = listaBloqueados.some(function(row) { 
      return String(row[0]).trim() === String(userIp).trim(); 
    });
  }
  
  // Se o usuário estiver bloqueado, recusa a requisição imediatamente
  if (bloqueado) {
    return ContentService.createTextOutput(JSON.stringify({ status: "bloqueado", mensagem: "Acesso negado." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = e.parameter ? (e.parameter.action || "consultar") : "consultar";  
  
  var agora = new Date();
  var dataHoje = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy");
  var horaFormatada = Utilities.formatDate(agora, "America/Sao_Paulo", "HH:mm");
  
  var ultimaLinha = sheetDados.getLastRow();
  // Captura até a coluna F (índice 6) para ler todas as informações
  var dadosUltimaLinha = ultimaLinha > 1 ? sheetDados.getRange(ultimaLinha, 1, 1, 6).getValues()[0] : ["", "", "", "", "", ""];
  
  function formatarDataPlanilha(d) {
    if (!d) return "";
    if (d instanceof Date || Object.prototype.toString.call(d) === '[object Date]') {
      return Utilities.formatDate(d, "America/Sao_Paulo", "dd/MM/yyyy");
    }
    return String(d).trim().split(" ")[0];
  }

  function formatarHoraPlanilha(h) {
    if (!h) return "";
    if (h instanceof Date || Object.prototype.toString.call(h) === '[object Date]') {
      return Utilities.formatDate(h, "America/Sao_Paulo", "HH:mm");
    }
    return String(h).trim();
  }

  var dataUltimoRegistro = formatarDataPlanilha(dadosUltimaLinha[0]);
  var horaAcabouUltimo = formatarHoraPlanilha(dadosUltimaLinha[3]); 

  // 1. AÇÃO: Registrar que o pastel ficou PRONTO
  if (action === "registrar_pronto") {
    if (dataUltimoRegistro !== dataHoje) {
      var horas = agora.getHours();
      var minutos = agora.getMinutes();
      var minutosDoDia = (horas * 60) + minutos;
      
      // Salva: Data, Hora, Minutos, Hora_Acabou (vazio), IP_Pronto (userIp), IP_Acabou (vazio)
      sheetDados.appendRow([dataHoje, horaFormatada, minutosDoDia, "", userIp, ""]);
      
      ultimaLinha = sheetDados.getLastRow();
      dataUltimoRegistro = dataHoje;
      horaAcabouUltimo = "";
    }
  }
  
  // 2. AÇÃO: Registrar que o pastel ACABOU
  if (action === "registrar_acabou") {
    if (dataUltimoRegistro === dataHoje && horaAcabouUltimo === "") {
      // Registra a hora que acabou na coluna D (4)
      sheetDados.getRange(ultimaLinha, 4).setValue(horaFormatada);
      // Registra o IP de quem marcou na coluna F (6)
      sheetDados.getRange(ultimaLinha, 6).setValue(userIp);
      SpreadsheetApp.flush(); 
      horaAcabouUltimo = horaFormatada; 
    }
  }
  
  // 3. CONSULTA: Gera estatísticas e monta as curvas de Gauss
  var dados = sheetDados.getDataRange().getValues();
  var listaMinutosPronto = [];
  var listaMinutosAcabou = [];
  var statusHoje = "aguardando";
  var horaProntoHoje = "";

  for (var i = 1; i < dados.length; i++) {
    // Coleta coluna C (Minutos Pronto)
    if (dados[i][2] !== "") {
      listaMinutosPronto.push(Number(dados[i][2]));
    }
    // Coleta coluna D (Hora Acabou) e converte para minutos
    var horaAcabouStr = formatarHoraPlanilha(dados[i][3]);
    if (horaAcabouStr !== "") {
       var partes = horaAcabouStr.split(':');
       if (partes.length === 2) {
          var minAcabou = parseInt(partes[0], 10) * 60 + parseInt(partes[1], 10);
          listaMinutosAcabou.push(minAcabou);
       }
    }
  }
  
  if (dados.length > 1) {
    var ultima = dados[dados.length - 1];
    var dataUltimaCalculada = formatarDataPlanilha(ultima[0]);
    
    if (dataUltimaCalculada === dataHoje) {
      horaProntoHoje = formatarHoraPlanilha(ultima[1]); 
      
      if (formatarHoraPlanilha(ultima[3]) === "") {
        statusHoje = "disponivel"; 
      } else {
        statusHoje = "esgotado"; 
      }
    }
  }
  
  var resposta = {
    historico: listaMinutosPronto,
    historicoAcabou: listaMinutosAcabou,
    statusHoje: statusHoje,
    horaPronto: horaProntoHoje
  };
  
  return ContentService.createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}