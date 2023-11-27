function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Contrato Automatizado');
  menu.addItem('Gerar contrato', 'createNewGoogleDocs');
  menu.addToUi();
}


function createNewGoogleDocs() {
  // This value should be the id of your document template that we created in the last step
  const googleDocTemplate = DriveApp.getFileById('1wXRAysLnydppKhEHVtVwOi2KQmKF44a5BXpzN1VmlVU');


  // This value should be the id of the folder where you want your completed documents stored
  const destinationFolderId = '1o0zD26abD4Lx6zHzeu84CiqEwKpO6y2s';
 
  // Here we store the sheet as a variable
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dados do cliente');
 
  // Get the destination folder
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);


  // Create a single document to store all entries
  const combinedDocument = DocumentApp.create('Combined Contract Document');
  const combinedBody = combinedDocument.getBody();
 
  // Now we get all of the values as a 2D array
  const rows = sheet.getDataRange().getValues();
 
  // Start processing each spreadsheet row
  rows.forEach(function(row, index) {
    // Here we check if this row is the headers or if a document has already been generated, if so we skip it
    if (index === 0 || row[1]) return;


    // Using the row data in a template literal, we make a copy of our template document in our destinationFolder
    const copy = googleDocTemplate.makeCopy(`Contrato ${row[4]} ${row[5]}`, destinationFolder);
    // Once we have the copy, we then open it using the DocumentApp
    const doc = DocumentApp.openById(copy.getId());
    // All of the content lives in the body, so we get that for editing
    const body = doc.getBody();
    // In this line we do some friendly date formatting, that may or may not work for your locale
    const data1 = new Date(row[22]).toLocaleDateString();
    const data2 = new Date(row[25]).toLocaleDateString();
    const data3 = new Date(row[28]).toLocaleDateString();


        let numeroEntrada = row[21];
    let numeroEntradaFormatado = numeroEntrada.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    console.log(numeroEntradaFormatado);
        let numeroParcela1 = row[24];
    let numeroParcela1Formatado = numeroParcela1.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    console.log(numeroParcela1Formatado);
        let numeroParcela2 = row[27];
    let numeroParcela2Formatado = numeroParcela2.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    console.log(numeroParcela2Formatado);
        let numeroParcela3 = row[30];
    let numeroParcela3Formatado = numeroParcela3.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    console.log(numeroParcela3Formatado);
        let numeroTotal = row[31];
    let numeroTotalFormatado = numeroTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    console.log(numeroTotalFormatado);


   
    // In these lines, we replace our replacement tokens with values from our spreadsheet row
    body.replaceText('{{Nome completo do cliente para contrato}}', row[4]);
    body.replaceText('{{RG}}', row[6]);
    body.replaceText('{{CPF}}', row[5]);
    body.replaceText('{{Endereço}}', row[9]);
    body.replaceText('{{CEP}}', row[10]);
    body.replaceText('{{Telefone}}', row[7]);
    body.replaceText('{{PRAZO PARA MEDIÇÃO FINAL DOS AMBIENTES}}', row[15]);
    body.replaceText('{{PRAZO PARA CADERNO EXECUTIVO}}', row[16]);
    body.replaceText('{{PRAZO PARA ENTREGA DOS MÓVEIS}}', row[17]);
    body.replaceText('{{PRAZO PARA MONTAGEM considerar apenas horário comercial}}', row[18]);
    body.replaceText('{{FORMA DE PAGAMENTO DA ENTRADA}}', row[20]);


    body.replaceText('{{VALOR DA ENTRADA}}', numeroEntradaFormatado);


    body.replaceText('{{FORMA DE PAGAMENTO DA PARCELA 1}}', row[23]);
    body.replaceText('{{1ª PARCELA  DATA}} ',  data1);


    body.replaceText('{{VALOR DE PAGAMENTO DA PARCELA 1}}', numeroParcela1Formatado);


    body.replaceText('{{FORMA DE PAGAMENTO DA PARCELA 2}}', row[26]);
    body.replaceText('{{2ª PARCELA  DATA}} ',  data2);


    body.replaceText('{{VALOR DE PAGAMENTO DA PARCELA 2}}', numeroParcela2Formatado);


    body.replaceText('{{FORMA DE PAGAMENTO DA PARCELA 3}}', row[29]);
    body.replaceText('{{3ª PARCELA  DATA}} ',  data3);


    body.replaceText('{{VALOR DE PAGAMENTO DA PARCELA 3}}', numeroParcela3Formatado);


    body.replaceText('{{TOTAL}}', numeroTotalFormatado);


    body.replaceText('{{Endereço de entrega}}', row[11]);
    body.replaceText('{{CEP do endereço de entrega}}', row[12]);
    body.replaceText('{{NOME DO CONDOMÍNIO}}', row[13]);
    body.replaceText('{{BAIRRO}}', row[14]);
    body.replaceText('{{DIA DO MÊS DE ASSINATURA DO CONTRATO}}', row[32]);
    body.replaceText('{{MÊS DE ASSINATURA DO CONTRATO}}', row[33]);
    body.replaceText('{{ANO DE ASSINATURA DO CONTRATO}}', row[34]);
   
      // Append the contents of the current document to the combined document
    combinedBody.appendParagraph(body.getText());


    // Save and close the current document
    doc.saveAndClose();


    // Create a PDF copy of the document
    const pdfCopy = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
    const pdfFileName = `Contrato ${row[4]} ${row[5]}.pdf`;
    const pdfFile = destinationFolder.createFile(pdfCopy);
    pdfFile.setName(pdfFileName);


        // Send the PDF as an email attachment
    const recipientEmail = row[8];
    const emailsubject = `Contrato ${row[4]}`;
    const emailbody = `Caro ${row[4]},\n\nSegue anexa a sua via do contrato. Favor assiná-lo e responder a esta mensagem com o anexo de contrato assinado.\n\nAtenciosamente,\nRaphael Lima`;
    const emailadvancedOpts = { attachments: [pdfFile.getAs(MimeType.PDF)] };


    MailApp.sendEmail(recipientEmail, emailsubject, emailbody, emailadvancedOpts);


    // Write the URL to the 'Document Link' column in the spreadsheet
    sheet.getRange(index + 1, 2).setValue(doc.getUrl());


    // Save and close the current document
    doc.saveAndClose();
  });


  // Save and close the combined document
  combinedDocument.saveAndClose();
}
