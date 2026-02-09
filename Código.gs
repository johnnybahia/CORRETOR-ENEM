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
 * Espera-se: Coluna A = Nomes, Coluna B = Emails.
 */
function getStudentsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DADOS');
  if (!sheet) return [];
  
  // Pega todos os dados da linha 2 até a última (ignorando cabeçalho)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  
  // Retorna um array de objetos {name, email}
  return data.map(row => ({
    name: String(row[0]).trim(),
    email: String(row[1]).trim()
  })).filter(student => student.name !== ""); // Remove linhas vazias
}

/**
 * Recebe o PDF em base64, salva no Drive e envia por e-mail.
 */
function processCorrection(pdfBase64, studentName, studentEmail, lessonName) {
  try {
    // 1. Converter o base64 de volta para arquivo binário (Blob)
    const blob = Utilities.newBlob(
      Utilities.base64Decode(pdfBase64.split(',')[1]), 
      'application/pdf', 
      `Correcao-${lessonName}-${studentName}.pdf`
    );

    // 2. Criar ou obter a pasta principal "Correções Redação"
    const rootFolder = DriveApp.getRootFolder();
    const folderName = "Correções Redação";
    let folder;
    const folders = rootFolder.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = rootFolder.createFolder(folderName);
    }

    // 3. (Opcional) Criar subpasta para o Aluno
    let studentFolder;
    const studentFolders = folder.getFoldersByName(studentName);
    if (studentFolders.hasNext()) {
      studentFolder = studentFolders.next();
    } else {
      studentFolder = folder.createFolder(studentName);
    }

    // 4. Salvar o arquivo na pasta do aluno
    const file = studentFolder.createFile(blob);
    const fileUrl = file.getUrl();

    // 5. Enviar E-mail (se houver e-mail cadastrado)
    let emailStatus = "E-mail não enviado (endereço não encontrado).";
    if (studentEmail && studentEmail.includes('@')) {
      MailApp.sendEmail({
        to: studentEmail,
        subject: `Correção da Redação: ${lessonName}`,
        htmlBody: `
          <p>Olá, <strong>${studentName}</strong>.</p>
          <p>Sua correção referente à aula <strong>${lessonName}</strong> está pronta.</p>
          <p>O arquivo PDF com as anotações e sua nota detalhada está em anexo.</p>
          <p>Atenciosamente,<br>Professora Jaqueline Cardoso</p>
        `,
        attachments: [blob]
      });
      emailStatus = "E-mail enviado com sucesso!";
    }

    return {
      success: true,
      url: fileUrl,
      emailStatus: emailStatus
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}
