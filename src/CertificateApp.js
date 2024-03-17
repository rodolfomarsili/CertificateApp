/**
 * CertificateApp module for managing user certificates and batch processing.
 * @namespace CertificateApp
 */
const CertificateApp = (function() {
  
  let BASE_PRESENTATION_ID = ''; // ID of the base presentation template
  let DESTINATION_FOLDER_ID = ''; // ID of the folder where certificates will be stored
  let PLACEHOLDER = ''; // Placeholder text to be replaced in the presentation
  
  /**
   * Represents a user.
   * @class
   */
  class User {
    constructor() {
      let _name = ''; // User's name
      let _email = ''; // User's email
      let _certificateID = ''; // ID of the user's certificate
      
      /**
       * Set the name of the user.
       * @method
       * @param {string} name - The name of the user.
       * @returns {User} The User instance.
       */
      this.setName = function(name) {
        _name = name;
        return this;
      };
      
      /**
       * Get the name of the user.
       * @method
       * @returns {string} The name of the user.
       */
      this.getName = function() {
        return _name;
      };
      
      /**
       * Set the email of the user.
       * @method
       * @param {string} email - The email of the user.
       * @returns {User} The User instance.
       */
      this.setEmail = function(email) {
        _email = email;
        return this;
      };
      
      /**
       * Get the email of the user.
       * @method
       * @returns {string} The email of the user.
       */
      this.getEmail = function() {
        return _email;
      };
      
      /**
       * Set the ID of the user's certificate.
       * @method
       * @param {string} certificateID - The ID of the user's certificate.
       * @returns {User} The User instance.
       */
      this.setCertificateID = function(certificateID) {
        _certificateID = certificateID;
        return this;
      };
      
      /**
       * Get the ID of the user's certificate.
       * @method
       * @returns {string} The ID of the user's certificate.
       */
      this.getCertificateID = function() {
        return _certificateID;
      };
      
      /**
       * Convert the User instance to JSON format.
       * @method
       * @returns {Object} JSON representation of the User instance.
       */
      this.toJSON = function() {
        return {
          name: _name,
          email: _email,
          certificateID: _certificateID
        };
      };
      
      /**
       * Initialize the User instance from a JSON object.
       * @method
       * @param {Object} jsonObject - JSON object representing the User.
       */
      this.fromJSON = function(jsonObject) {
        _name = jsonObject?.name;
        _email = jsonObject?.email;
        _certificateID = jsonObject?.certificateID || '';
      };
      
      /**
       * Send an email to the user with their certificate attached.
       * @method
       * @param {Object} params - Parameters for sending the email.
       */
      this.sendEmail = function(params) {
        if(!_certificateID) return;
        const subject = 'Certificado de Participação no Treinamento';
        const htmlBody = `<div dir="ltr">Oi ${_name}!<br><br>Agradecemos muito a sua participação no nosso treinamento.<br><br>Esperamos que tenha gostado dele!<br><br>Aqui está o seu certificado de participação.</div>\r\n`;
        
        MailApp.sendEmail({
          name: params?.senderName || "",
          to: _email,
          subject: params?.subject || subject,
          htmlBody: params?.htmlBody || htmlBody,
          attachments: [DriveApp.getFileById(_certificateID)]
        });
        showFeedback(`Email to ${_name} successfully sent`);
      };
    }
  }
  
  /**
   * Represents a certificate.
   * @class
   */
  class Certificate {
    constructor() {
      let _placeholderReplacement = ''; // Placeholder replacement for the certificate
      let _certificateName = ''; // Name of the certificate
      let _deletePresentation = false; // Whether to delete the presentation after creating the certificate

      let _certificateID = ''; // ID of the certificate
      
      /**
       * Set the placeholder replacement for the certificate.
       * @method
       * @param {string} placeholderReplacement - Placeholder replacement for the certificate.
       * @returns {Certificate} The Certificate instance.
       */
      this.setPlaceholderReplacement = function(placeholderReplacement) {
        _placeholderReplacement = placeholderReplacement;
        return this;
      };
      
      /**
       * Set the name of the certificate.
       * @method
       * @param {string} certificateName - Name of the certificate.
       * @returns {Certificate} The Certificate instance.
       */
      this.setCertificateName = function(certificateName) {
        _certificateName = certificateName;
        return this;
      };
      
      /**
       * Set whether to delete the presentation after creating the certificate.
       * @method
       * @param {boolean} deletePresentation - Whether to delete the presentation.
       * @returns {Certificate} The Certificate instance.
       */
      this.deletePresentation = function(deletePresentation) {
        _deletePresentation = deletePresentation;
        return this;
      };

      /**
       * Get the ID of the certificate.
       * @method
       * @returns {string} The ID of the certificate.
       */
      this.getCertificateID = function() {
        return _certificateID;
      };
      
      /**
       * Create the certificate based on the template.
       * @method
       * @returns {Certificate} The Certificate instance.
       */
      this.createCertificate = function() {
        let destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
        let presentationBlob = DriveApp.getFileById(BASE_PRESENTATION_ID)
                                          .makeCopy()
                                          .moveTo(destinationFolder)
                                          .setName(_certificateName);
        let presentationBlobID = presentationBlob.getId();
        let request = {
          replaceAllText: {
            replaceText: _placeholderReplacement,
            containsText: {
              text: PLACEHOLDER,
              matchCase: false
            }
          }
        };
        Slides.Presentations.batchUpdate({ requests: [request] }, presentationBlobID);
        let pdfBlob = presentationBlob.getAs('application/pdf');
        let certificate = destinationFolder.createFile(pdfBlob);
        _certificateID = certificate.getId();
        presentationBlob.setTrashed(_deletePresentation);
        showFeedback(`${_certificateName} successfully created`);
        return this;
      };
    }
  }
  
  /**
   * Represents a batch processor for managing multiple users and certificates.
   * @class
   */
  class BatchProcessor {
    constructor() {
      let _users = []; // Array of User instances
      let _spreadsheetID = ''; // ID of the Google Sheets spreadsheet
      let _sheetName = ''; // Name of the sheet containing user data
      let _namesColumnHeader = ''; // Header name for user names column
      let _emailsColumnHeader = ''; // Header name for user emails column
      let _certificatesIDColumnHeader = ''; // Header name for certificate IDs column
      let _deletePresentations = false; // Whether to delete presentations after creating certificates
      
      /**
       * Set the ID of the Google Sheets spreadsheet.
       * @method
       * @param {string} spreadsheetID - ID of the spreadsheet.
       * @returns {BatchProcessor} The BatchProcessor instance.
       */
      this.setSpreadsheetID = function(spreadsheetID) {
        _spreadsheetID = spreadsheetID;
        return this;
      };
      
      /**
       * Set the name of the sheet containing user data.
       * @method
       * @param {string} sheetName - Name of the sheet.
       * @returns {BatchProcessor} The BatchProcessor instance.
       */
      this.setSheetName = function(sheetName) {
        _sheetName = sheetName;
        return this;
      };
      
      /**
       * Set the header name for user names column.
       * @method
       * @param {string} namesColumnHeader - Header name.
       * @returns {BatchProcessor} The BatchProcessor instance.
       */
      this.setNamesColumnHeader = function(namesColumnHeader) {
        _namesColumnHeader = namesColumnHeader;
        return this;
      };
      
      /**
       * Set the header name for user emails column.
       * @method
       * @param {string} emailsColumnHeader - Header name.
       * @returns {BatchProcessor} The BatchProcessor instance.
       */
      this.setEmailsColumnHeader = function(emailsColumnHeader) {
        _emailsColumnHeader = emailsColumnHeader;
        return this;
      };

      /**
       * Set the header name for certificate IDs column.
       * @method
       * @param {string} certificatesIDColumnHeader - Header name.
       * @returns {BatchProcessor} The BatchProcessor instance.
       */
      this.setCertificatesIDColumnHeader = function(certificatesIDColumnHeader) {
        _certificatesIDColumnHeader = certificatesIDColumnHeader;
        return this;
      };

      /**
       * Set whether to delete presentations after creating certificates.
       * @method
       * @param {boolean} deletePresentations - Whether to delete presentations.
       * @returns {BatchProcessor} The BatchProcessor instance.
       */
      this.deletePresentations = function(deletePresentations) {
        _deletePresentations = deletePresentations;
        return this;
      };
      
      /**
       * Retrieve users from the specified sheet and populate the _users array.
       * @method
       */
      this.getUsersFromSheet = function() {
        const sheet = SpreadsheetApp.openById(_spreadsheetID)
                                    .getSheetByName(_sheetName);
        const [headers, ...data] = sheet.getDataRange().getValues();
        const namesColumn = headers.indexOf(_namesColumnHeader);
        const emailsColumn = headers.indexOf(_emailsColumnHeader);
        const certificatesIDColumn = headers.indexOf(_certificatesIDColumnHeader);
        
        data.forEach(row => {
          if (!row[namesColumn] || !row[emailsColumn]) return;
          const user = new User().setName(row[namesColumn].toString().trim())
                                 .setEmail(row[emailsColumn].toString().trim())
                                 .setCertificateID(row[certificatesIDColumn].toString().trim());
          _users.push(user);
        });
      };

      /**
       * Save users' data back to the sheet.
       * @method
       */
      this.saveUsersToSheet = function() {
        const sheet = SpreadsheetApp.openById(_spreadsheetID)
                                    .getSheetByName(_sheetName);
        const headers = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
        const namesColumn = headers.indexOf(_namesColumnHeader);
        const emailsColumn = headers.indexOf(_emailsColumnHeader);
        const certificatesIDColumn = headers.indexOf(_certificatesIDColumnHeader);
        let rows = [];
        _users.forEach(user => { 
          let row = [];
          row[namesColumn] = user.getName();
          row[emailsColumn] = user.getEmail();
          row[certificatesIDColumn] = user.getCertificateID();
          rows.push(row);
        });
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      };
      
      /**
       * Create certificates for all users in the _users array.
       * @method
       */
      this.createAllCertificates = function() {
        _users.forEach(user => { 
          let certificate = CertificateApp.newCertificate().setCertificateName(`Certificate: ${user.getName()}`)
                                                           .setPlaceholderReplacement(user.getName())
                                                           .deletePresentation(_deletePresentations)
                                                           .createCertificate();
          user.setCertificateID(certificate.getCertificateID());
        });
      };

      /**
       * Send emails to all users with their certificates attached.
       * @method
       * @param {Object} params - Parameters for sending the email.
       */
      this.sendEmails = function(params) {
        _users.forEach(user => { 
          user.sendEmail(params)
        });
      }
    }
  }
  
  return {
    /**
     * Create a new User instance.
     * @method
     * @returns {User} A new User instance.
     */
    newUser: function() {
      return new User();
    },
    /**
     * Create a new Certificate instance.
     * @method
     * @returns {Certificate} A new Certificate instance.
     */
    newCertificate: function() {
      return new Certificate();
    },
    /**
     * Create a new BatchProcessor instance.
     * @method
     * @returns {BatchProcessor} A new BatchProcessor instance.
     */
    newBatchProcessor: function() {
      return new BatchProcessor();
    },
    /**
     * Initialize the CertificateApp module with base presentation ID, destination folder ID, and placeholder text.
     * @method
     * @param {string} basePresentationID - ID of the base presentation template.
     * @param {string} destinationFolderID - ID of the folder where certificates will be stored.
     * @param {string} placeholder - Placeholder text to be replaced in the presentation.
     */
    init: function(basePresentationID, destinationFolderID, placeholder) {
      BASE_PRESENTATION_ID = basePresentationID;
      DESTINATION_FOLDER_ID = destinationFolderID;
      PLACEHOLDER = placeholder;
    }
  };
})();

/**
 * Retrieve the CertificateApp module.
 * @function
 * @returns {Object} The CertificateApp module.
 */
function getApp() {
  return CertificateApp;
}

/**
 * Show feedback message.
 * @function
 * @param {string} feedback - Feedback message to be displayed.
 */
function showFeedback(feedback) {
  console.log(feedback);
  SpreadsheetApp.getActive().toast(feedback);
}
