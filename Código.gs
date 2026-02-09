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
 * Recebe o PDF completo (correcao + notas + historico já embutido),
 * salva no histórico da planilha, salva no Drive e envia por e-mail.
 * O PDF já vem pronto do frontend com todas as páginas.
 */
function processCorrection(pdfBase64, studentName, studentEmail, lessonNumber, theme, scores, folderId) {
  try {
    const { c1, c2, c3, c4, c5, total } = scores;

    // 1. Salvar no histórico da planilha
    const histResult = saveToHistory(studentName, lessonNumber, theme, c1, c2, c3, c4, c5, total);

    // 2. Converter o PDF (já contém correcao + notas + historico + grafico)
    const pdfBlob = Utilities.newBlob(
      Utilities.base64Decode(pdfBase64.split(',')[1]),
      'application/pdf',
      'Correcao-Aula' + lessonNumber + '-' + sanitizeName(studentName) + '.pdf'
    );

    // 3. Salvar na pasta escolhida
    let targetFolder;
    if (folderId) {
      targetFolder = DriveApp.getFolderById(folderId);
    } else {
      targetFolder = getOrCreateStudentFolder(studentName);
    }

    const file = targetFolder.createFile(pdfBlob);
    const fileUrl = file.getUrl();

    // 4. Enviar e-mail com o PDF único
    let emailStatus = "E-mail não enviado (endereço não encontrado).";
    if (studentEmail && studentEmail.includes('@')) {
      const history = getStudentHistory(studentName);
      MailApp.sendEmail({
        to: studentEmail,
        subject: 'Correção da Redação - Aula ' + lessonNumber + ': ' + theme,
        htmlBody: generateEmailHtml(studentName, lessonNumber, theme, scores, history),
        attachments: [pdfBlob]
      });
      emailStatus = "E-mail enviado com sucesso!";
    }

    return {
      success: true,
      url: fileUrl,
      emailStatus: emailStatus,
      historyMessage: histResult.message,
      historyCount: getStudentHistory(studentName).length
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
 * Gera o HTML do e-mail enviado ao aluno.
 * O PDF anexo já contém o histórico completo com gráfico (gerado no frontend).
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
      historyRows + '</table>'
      : '') +

    '<p style="margin-top:15px;">O PDF em anexo contém a correção completa com anotações, notas e histórico de desempenho.</p>' +
    '</div>' +

    '<div style="background:#f3f4f6;padding:15px;text-align:center;border-radius:0 0 8px 8px;border:1px solid #e5e7eb;border-top:none;">' +
    '<p style="margin:0;color:#6b7280;font-size:12px;">Professora Jaqueline Cardoso — Corretor de Redações ENEM</p>' +
    '</div></div>';
}
