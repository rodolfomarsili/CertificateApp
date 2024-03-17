# Setup
```javascript
const params = {
  placeholder: '',
  spreadsheetID: '',
  sheetName: '',
  basePresentationID: '',
  destinationFolderID: '', 
  senderName: ''
}

CertificateApp.init(
  params.basePresentationID,
  params.destinationFolderID,
  params.placeholder
)

let batchProcessor = CertificateApp.newBatchProcessor()
                                   .setSpreadsheetID(params.spreadsheetID)
                                   .setSheetName(params.sheetName)
                                   .setNamesColumnHeader("Name")
                                   .setEmailsColumnHeader("Email")
                                   .setCertificatesIDColumnHeader("Certificate ID")
                                   .deletePresentations(true);
```

## Examples

### Create Certificates
```javascript

  batchProcessor.getUsersFromSheet();
  batchProcessor.createAllCertificates();
  batchProcessor.saveUsersToSheet();

```

### Send Certificates by Email
```javascript

  batchProcessor.getUsersFromSheet();
  const emailData = {
    subject: '',
    htmlBody: '',
    senderName: ''
  };
  batchProcessor.sendEmails(emailData)

```
