/**
 * Função que serve a página HTML quando alguém acessa o Web App.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Corretor de Redações ENEM - Web App')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, user-scalable=no');
}

/**
 * Busca a lista de alunos e emails da aba "DADOS".
 * Coluna A = Nomes, Coluna B = Emails.
 */
function getStudentsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DADOS');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return data.map(row => ({
    name: String(row[0]).trim(),
    email: String(row[1]).trim()
  })).filter(student => student.name !== "");
}

/**
 * Busca o histórico de notas de um aluno na aba "HISTORICO".
 * Estrutura da aba HISTORICO:
 * A = Nome do Aluno | B = Nº Aula | C = Tema | D = Data |
 * E = C1 | F = C2 | G = C3 | H = C4 | I = C5 | J = Nota Final
 */
function getStudentHistory(studentName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('HISTORICO');

  if (!sheet) {
    sheet = ss.insertSheet('HISTORICO');
    sheet.appendRow(['Aluno', 'Nº Aula', 'Tema', 'Data', 'C1', 'C2', 'C3', 'C4', 'C5', 'Nota Final']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  return data
    .filter(row => String(row[0]).trim().toLowerCase() === studentName.trim().toLowerCase())
    .map(row => ({
      name: String(row[0]).trim(),
      lessonNumber: Number(row[1]),
      theme: String(row[2]).trim(),
      date: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      c1: Number(row[4]),
      c2: Number(row[5]),
      c3: Number(row[6]),
      c4: Number(row[7]),
      c5: Number(row[8]),
      total: Number(row[9])
    }))
    .sort((a, b) => a.lessonNumber - b.lessonNumber);
}

/**
 * Salva a correção no histórico (aba HISTORICO).
 */
function saveToHistory(studentName, lessonNumber, theme, c1, c2, c3, c4, c5, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('HISTORICO');

  if (!sheet) {
    sheet = ss.insertSheet('HISTORICO');
    sheet.appendRow(['Aluno', 'Nº Aula', 'Tema', 'Data', 'C1', 'C2', 'C3', 'C4', 'C5', 'Nota Final']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }

  // Verifica se já existe registro para este aluno nesta aula (evita duplicatas)
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === studentName.trim().toLowerCase()
          && Number(data[i][1]) === Number(lessonNumber)) {
        // Atualiza registro existente
        const row = i + 2;
        sheet.getRange(row, 1, 1, 10).setValues([[
          studentName, lessonNumber, theme, new Date(), c1, c2, c3, c4, c5, total
        ]]);
        return { updated: true, message: 'Registro atualizado na aula ' + lessonNumber };
      }
    }
  }

  // Novo registro
  sheet.appendRow([studentName, lessonNumber, theme, new Date(), c1, c2, c3, c4, c5, total]);
  return { updated: false, message: 'Novo registro salvo para aula ' + lessonNumber };
}

/**
 * Lista as pastas disponíveis no Drive para a professora escolher.
 * Retorna pastas dentro de "Correções Redação".
 */
function getDriveFolders() {
  const rootFolder = DriveApp.getRootFolder();
  const mainFolderName = "Correções Redação";
  let mainFolder;
  const folders = rootFolder.getFoldersByName(mainFolderName);

  if (folders.hasNext()) {
    mainFolder = folders.next();
  } else {
    mainFolder = rootFolder.createFolder(mainFolderName);
  }

  const result = [{ id: mainFolder.getId(), name: mainFolderName + ' (Raiz)', path: mainFolderName }];

  // Listar subpastas (1 nível)
  const subFolders = mainFolder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    result.push({ id: sub.getId(), name: sub.getName(), path: mainFolderName + '/' + sub.getName() });
  }

  return result.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Cria uma nova subpasta dentro de "Correções Redação".
 */
function createNewFolder(folderName) {
  const rootFolder = DriveApp.getRootFolder();
  const mainFolderName = "Correções Redação";
  let mainFolder;
  const folders = rootFolder.getFoldersByName(mainFolderName);

  if (folders.hasNext()) {
    mainFolder = folders.next();
  } else {
    mainFolder = rootFolder.createFolder(mainFolderName);
  }

  const newFolder = mainFolder.createFolder(folderName);
  return { id: newFolder.getId(), name: newFolder.getName(), path: mainFolderName + '/' + newFolder.getName() };
}

/**
 * Recebe o PDF da correção + dados, salva no histórico, gera relatório,
 * salva no Drive na pasta escolhida e envia e-mail com os anexos.
 */
function processCorrection(pdfBase64, studentName, studentEmail, lessonNumber, theme, scores, folderId) {
  try {
    const { c1, c2, c3, c4, c5, total } = scores;

    // 1. Salvar no histórico
    const histResult = saveToHistory(studentName, lessonNumber, theme, c1, c2, c3, c4, c5, total);

    // 2. Buscar histórico completo para o relatório
    const history = getStudentHistory(studentName);

    // 3. Converter o PDF da correção
    const correctionBlob = Utilities.newBlob(
      Utilities.base64Decode(pdfBase64.split(',')[1]),
      'application/pdf',
      'Correcao-Aula' + lessonNumber + '-' + sanitizeName(studentName) + '.pdf'
    );

    // 4. Gerar PDF do relatório de histórico
    const reportHtml = generateHistoryReportHtml(studentName, history);
    const reportBlob = HtmlService.createHtmlOutput(reportHtml)
      .getBlob()
      .setName('Historico-' + sanitizeName(studentName) + '.pdf');

    // 5. Salvar na pasta escolhida
    let targetFolder;
    if (folderId) {
      targetFolder = DriveApp.getFolderById(folderId);
    } else {
      targetFolder = getOrCreateStudentFolder(studentName);
    }

    const correctionFile = targetFolder.createFile(correctionBlob);
    const reportFile = targetFolder.createFile(reportBlob);
    const fileUrl = correctionFile.getUrl();

    // 6. Enviar e-mail com ambos os anexos
    let emailStatus = "E-mail não enviado (endereço não encontrado).";
    if (studentEmail && studentEmail.includes('@')) {
      const scoreTable = history.map(h =>
        '  Aula ' + h.lessonNumber + ' (' + h.date + '): ' + h.total + '/1000'
      ).join('\n');

      MailApp.sendEmail({
        to: studentEmail,
        subject: 'Correção da Redação - Aula ' + lessonNumber + ': ' + theme,
        htmlBody: generateEmailHtml(studentName, lessonNumber, theme, scores, history),
        attachments: [correctionBlob, reportBlob]
      });
      emailStatus = "E-mail enviado com sucesso!";
    }

    return {
      success: true,
      url: fileUrl,
      emailStatus: emailStatus,
      historyMessage: histResult.message,
      historyCount: history.length
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Auxiliar: remove acentos e caracteres especiais do nome para uso em arquivos.
 */
function sanitizeName(name) {
  return name.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '-');
}

/**
 * Auxiliar: obtém ou cria a pasta do aluno dentro de "Correções Redação".
 */
function getOrCreateStudentFolder(studentName) {
  const rootFolder = DriveApp.getRootFolder();
  const mainFolderName = "Correções Redação";
  let mainFolder;
  const folders = rootFolder.getFoldersByName(mainFolderName);

  if (folders.hasNext()) {
    mainFolder = folders.next();
  } else {
    mainFolder = rootFolder.createFolder(mainFolderName);
  }

  let studentFolder;
  const studentFolders = mainFolder.getFoldersByName(studentName);
  if (studentFolders.hasNext()) {
    studentFolder = studentFolders.next();
  } else {
    studentFolder = mainFolder.createFolder(studentName);
  }

  return studentFolder;
}

/**
 * Gera o HTML do relatório de histórico com tabela e gráfico embutido (SVG).
 */
function generateHistoryReportHtml(studentName, history) {
  // Gerar gráfico SVG embutido
  const svgChart = generateSvgChart(history);

  let tableRows = '';
  history.forEach(h => {
    tableRows += '<tr>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.lessonNumber + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;">' + h.theme + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.date + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.c1 + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.c2 + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.c3 + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.c4 + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;">' + h.c5 + '</td>' +
      '<td style="border:1px solid #ddd;padding:8px;text-align:center;font-weight:bold;">' + h.total + '</td>' +
      '</tr>';
  });

  // Calcular média
  let avgTotal = 0;
  if (history.length > 0) {
    avgTotal = Math.round(history.reduce((sum, h) => sum + h.total, 0) / history.length);
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:40px;color:#333;}' +
    'h1{color:#1e40af;border-bottom:3px solid #1e40af;padding-bottom:10px;}' +
    'h2{color:#374151;margin-top:30px;}' +
    'table{border-collapse:collapse;width:100%;margin-top:10px;}' +
    'th{background-color:#1e40af;color:white;padding:10px;border:1px solid #1e40af;text-align:center;}' +
    '.summary{background:#f0f9ff;border:1px solid #bfdbfe;border-radius:8px;padding:20px;margin-top:20px;}' +
    '.summary-grid{display:flex;justify-content:space-around;text-align:center;}' +
    '.summary-item .number{font-size:28px;font-weight:bold;color:#1e40af;}' +
    '.summary-item .label{font-size:12px;color:#6b7280;}' +
    '.chart-container{margin-top:20px;text-align:center;}' +
    '.footer{margin-top:40px;text-align:center;color:#9ca3af;font-size:11px;border-top:1px solid #e5e7eb;padding-top:10px;}' +
    '</style></head><body>' +
    '<h1>Relatório de Desempenho - Redação ENEM</h1>' +
    '<p><strong>Aluno(a):</strong> ' + studentName + '</p>' +
    '<p><strong>Data do Relatório:</strong> ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '</p>' +

    '<div class="summary"><div class="summary-grid">' +
    '<div class="summary-item"><div class="number">' + history.length + '</div><div class="label">Redações Corrigidas</div></div>' +
    '<div class="summary-item"><div class="number">' + avgTotal + '</div><div class="label">Média Geral</div></div>' +
    '<div class="summary-item"><div class="number">' + (history.length > 0 ? history[history.length - 1].total : 0) + '</div><div class="label">Última Nota</div></div>' +
    '</div></div>' +

    '<h2>Evolução das Notas</h2>' +
    '<div class="chart-container">' + svgChart + '</div>' +

    '<h2>Detalhamento por Aula</h2>' +
    '<table><thead><tr>' +
    '<th>Aula</th><th>Tema</th><th>Data</th><th>C1</th><th>C2</th><th>C3</th><th>C4</th><th>C5</th><th>Total</th>' +
    '</tr></thead><tbody>' + tableRows + '</tbody></table>' +

    '<div class="footer">Professora Jaqueline Cardoso — Corretor de Redações ENEM — Projeto Johnny v2.0</div>' +
    '</body></html>';
}

/**
 * Gera um gráfico SVG de linhas mostrando a evolução das notas.
 */
function generateSvgChart(history) {
  if (history.length === 0) {
    return '<p style="color:#999;">Nenhum dado de histórico disponível.</p>';
  }

  const width = 600, height = 300;
  const padding = { top: 30, right: 30, bottom: 60, left: 50 };
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;

  const maxScore = 1000;
  const numPoints = history.length;

  // Calcular posições dos pontos
  const points = history.map((h, i) => ({
    x: padding.left + (numPoints === 1 ? chartWidth / 2 : (i / (numPoints - 1)) * chartWidth),
    y: padding.top + chartHeight - (h.total / maxScore) * chartHeight,
    total: h.total,
    lesson: h.lessonNumber,
    date: h.date
  }));

  // Linha do gráfico
  let pathD = 'M ' + points.map(p => p.x + ' ' + p.y).join(' L ');

  // Área preenchida
  let areaD = pathD + ' L ' + points[points.length - 1].x + ' ' + (padding.top + chartHeight) +
              ' L ' + points[0].x + ' ' + (padding.top + chartHeight) + ' Z';

  // Linhas de grade horizontais
  let gridLines = '';
  for (let score = 0; score <= 1000; score += 200) {
    const y = padding.top + chartHeight - (score / maxScore) * chartHeight;
    gridLines += '<line x1="' + padding.left + '" y1="' + y + '" x2="' + (width - padding.right) + '" y2="' + y + '" stroke="#e5e7eb" stroke-width="1"/>';
    gridLines += '<text x="' + (padding.left - 8) + '" y="' + (y + 4) + '" text-anchor="end" fill="#6b7280" font-size="11">' + score + '</text>';
  }

  // Pontos e labels
  let pointsSvg = '';
  points.forEach(p => {
    pointsSvg += '<circle cx="' + p.x + '" cy="' + p.y + '" r="5" fill="#1e40af" stroke="white" stroke-width="2"/>';
    pointsSvg += '<text x="' + p.x + '" y="' + (p.y - 12) + '" text-anchor="middle" fill="#1e40af" font-size="12" font-weight="bold">' + p.total + '</text>';
    pointsSvg += '<text x="' + p.x + '" y="' + (padding.top + chartHeight + 20) + '" text-anchor="middle" fill="#374151" font-size="10">Aula ' + p.lesson + '</text>';
    pointsSvg += '<text x="' + p.x + '" y="' + (padding.top + chartHeight + 35) + '" text-anchor="middle" fill="#9ca3af" font-size="9">' + p.date + '</text>';
  });

  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + width + ' ' + height + '" width="100%" style="max-width:600px;">' +
    '<rect width="' + width + '" height="' + height + '" fill="white" rx="8"/>' +
    gridLines +
    '<path d="' + areaD + '" fill="rgba(30,64,175,0.1)"/>' +
    '<path d="' + pathD + '" fill="none" stroke="#1e40af" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>' +
    pointsSvg +
    '</svg>';
}

/**
 * Gera o HTML do e-mail enviado ao aluno.
 */
function generateEmailHtml(studentName, lessonNumber, theme, scores, history) {
  let historyRows = '';
  history.forEach(h => {
    const isCurrent = h.lessonNumber === Number(lessonNumber);
    const style = isCurrent ? 'background-color:#eff6ff;font-weight:bold;' : '';
    historyRows += '<tr style="' + style + '">' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;">Aula ' + h.lessonNumber + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;">' + h.date + '</td>' +
      '<td style="border:1px solid #ddd;padding:6px;text-align:center;font-weight:bold;">' + h.total + '/1000</td>' +
      '</tr>';
  });

  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#1e40af;color:white;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="margin:0;">Correção da Redação</h2>' +
    '<p style="margin:5px 0 0;">Aula ' + lessonNumber + ' — ' + theme + '</p></div>' +

    '<div style="padding:20px;background:#f9fafb;border:1px solid #e5e7eb;">' +
    '<p>Olá, <strong>' + studentName + '</strong>!</p>' +
    '<p>Sua correção referente à <strong>Aula ' + lessonNumber + '</strong> está pronta.</p>' +

    '<div style="background:white;border:1px solid #e5e7eb;border-radius:8px;padding:15px;margin:15px 0;">' +
    '<h3 style="margin:0 0 10px;color:#374151;">Notas por Competência</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:4px 8px;">C1 - Norma culta</td><td style="text-align:right;font-weight:bold;">' + scores.c1 + '</td></tr>' +
    '<tr><td style="padding:4px 8px;">C2 - Compreensão da proposta</td><td style="text-align:right;font-weight:bold;">' + scores.c2 + '</td></tr>' +
    '<tr><td style="padding:4px 8px;">C3 - Organizar informações</td><td style="text-align:right;font-weight:bold;">' + scores.c3 + '</td></tr>' +
    '<tr><td style="padding:4px 8px;">C4 - Mecanismos linguísticos</td><td style="text-align:right;font-weight:bold;">' + scores.c4 + '</td></tr>' +
    '<tr><td style="padding:4px 8px;">C5 - Proposta de intervenção</td><td style="text-align:right;font-weight:bold;">' + scores.c5 + '</td></tr>' +
    '<tr style="border-top:2px solid #1e40af;"><td style="padding:8px;font-weight:bold;font-size:16px;">TOTAL</td>' +
    '<td style="text-align:right;font-weight:bold;font-size:16px;color:#1e40af;">' + scores.total + '/1000</td></tr>' +
    '</table></div>' +

    (history.length > 1 ?
      '<h3 style="color:#374151;">Seu Histórico</h3>' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<tr style="background:#1e40af;color:white;"><th style="padding:8px;">Aula</th><th style="padding:8px;">Data</th><th style="padding:8px;">Nota</th></tr>' +
      historyRows + '</table>' +
      '<p style="font-size:12px;color:#6b7280;margin-top:10px;">Veja o relatório completo com gráfico de evolução no PDF anexo.</p>'
      : '') +

    '<p>Os arquivos em anexo contêm:<br>' +
    '1. PDF da correção com anotações<br>' +
    '2. Relatório de desempenho com gráfico de evolução</p>' +
    '</div>' +

    '<div style="background:#f3f4f6;padding:15px;text-align:center;border-radius:0 0 8px 8px;border:1px solid #e5e7eb;border-top:none;">' +
    '<p style="margin:0;color:#6b7280;font-size:12px;">Professora Jaqueline Cardoso — Corretor de Redações ENEM</p>' +
    '</div></div>';
}
