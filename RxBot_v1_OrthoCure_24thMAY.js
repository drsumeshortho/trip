
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('OrthoCure Tools')
    .addItem('Send Rx (HTML-PDF)', 'sendDynamicRxHTMLPDF')
    .addToUi();
}

function sendDynamicRxHTMLPDF() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RxBot');
  const row = sheet.getActiveRange().getRow();

  const name = sheet.getRange(row, 1).getValue();
  const age = sheet.getRange(row, 2).getValue();
  const gender = sheet.getRange(row, 3).getValue();
  const email = sheet.getRange(row, 5).getValue();
  const diagnosis = sheet.getRange(row, 7).getValue();
  const advice = sheet.getRange(row, 8).getValue();
  const medications = sheet.getRange(row, 9).getValue();
  const followup = sheet.getRange(row, 10).getValue();

  const html = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 30px; color: #000; }
        h1 { font-size: 22px; font-weight: bold; color: #00324E; border-bottom: 1px solid #ccc; padding-bottom: 8px; }
        h2 { font-size: 16px; margin-top: 20px; }
        .section { margin-bottom: 15px; }
        .signature { margin-top: 40px; }
        .footer { font-size: 10px; margin-top: 30px; color: #555; }
      </style>
    </head>
    <body>
      <h1>
        OrthoCure Bone & Joint Speciality Clinic<br/>
        <span style="font-size:16px; font-weight:normal; color:#008080;">
          Online Consultation Prescription
        </span>
      </h1>
      <div class="section">Date: ${new Date().toLocaleDateString()}</div>
      <div class="section"><strong>Patient Name:</strong> ${name}</div>
      <div class="section">Age: ${age} &nbsp;&nbsp;&nbsp;&nbsp; Gender: ${gender}</div>

      <h2>Diagnosis</h2>
      <div class="section">${diagnosis}</div>

      <h2>Advice</h2>
      <div class="section">${advice.replace(/\n/g, '<br/>')}</div>

      <h2>Medications</h2>
      <div class="section">${medications.replace(/\n/g, '<br/>')}</div>

      <h2>Follow-up Date</h2>
      <div class="section">${followup}</div>

      <div class="signature">
        <img src="https://raw.githubusercontent.com/drsumeshortho/trip/main/sig.jpg" width="180"/><br/>
        <strong>Dr. Sumesh Subramanian</strong><br/>
        MBBS, MS Ortho (TNMC 119811)<br/>
        OrthoCure Bone & Joint Speciality Clinic
      </div>

      <div class="footer">
        Note: This prescription is issued after an online consultation and is based on information provided by the patient.
        It does not replace a physical examination. For emergencies, visit your nearest hospital.
      </div>
    </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, "text/html", "Prescription.html");
  const pdfFile = DriveApp.createFile(blob.getAs("application/pdf"));
  const subject = "Your OrthoCure Prescription";
  const body = "Dear " + name + ",\n\n" +
    "Thank you for your consultation.\n\n" +
    "Attached is your prescription from Dr. Sumesh Subramanian.\n\n" +
    "Please note: This prescription is based on the information provided during online consultation and is not a substitute for physical examination.\n\n" +
    "For follow-up or queries, feel free to reply to this email.\n\n" +
    "– OrthoCure Bone & Joint Speciality Clinic";

  GmailApp.sendEmail(email, subject, body, {
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: "OrthoCure Clinic"
  });

  sheet.getRange(row, 13).setValue("Rx Sent: " + new Date().toLocaleString());
  SpreadsheetApp.getUi().alert("Prescription sent to " + email);
}
