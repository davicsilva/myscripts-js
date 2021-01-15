/**
 * Function sendEmailPromocao(): this function reads a specific number of lines from a 
 *                       spreadsheet/sheet - in this case the sheet "Promoções"- and,
 *                       an email is sent to: 
 *                                   - dextran@'s manager and 
 *                                   - dextran@'s mentor (if there is) and
 *                                   - dextran@'s tribe leader
 *                       with information about the dextran@ who is been recognized. 
 *                       
 * authors: 
 *      - edite.martins@dextra-sw.com
 *      - davi.silva@dextra-sw.com
 *
 * versions:
 * v4, November, 2020
 * v3, July 7th, 2020
 * v2, October, 9th, 2019
 */
function sendEmailPromocao() {
    // Indentify the sheet to process
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Promocoes");

    // First row and number of rows from this sheet
    var startRow = 2; // First row of data to process
    var numRows = 20; // Number of rows to process

    // Fetch the range of cells 
    var dataRange = sheet.getRange(startRow, 1, numRows, 10);

    // Fetch values for each row in the Range.
    var data = dataRange.getValues();

    // Email content and subject (initial values)
    var message = "";
    var subject = "";
    var emailAddress = "";

    // Incluir Edite na cópia
    var peopleAddress = "edite.martins@dextra-sw.com"

    var ui = SpreadsheetApp.getUi();

    var emailIntro = "";        // 1st column   "Car@s,oa promoção de..."
    var nameDextrano = "";        // 2nd column   "NOME "  
    var emailResumo1 = "";        // 3rd column   "foi aprovada. A promoção será refletida..."
    var emailResumo2 = "";        // 4rd column   "foi aprovada. A promoção será refletida..."  
    var currentPosition = "";        // 5th column   "Analista de Software Jr/Pl/Sr/..."
    var newPosition = "";        // 6th column   "Analista de Software Jr/Pl/Sr/..."
    var valueAdd = "";        // 7th column    Additional benefits ("basket") 
    //---
    var emailManager = "";        // 8th column
    var emailMentor = "";        // 9th column
    var emailTribe = "";        // 10th column 
    //---
    var detailPromotion = "";        // Current Role, 

    for (var i in data) {
        // Get the line to process
        var row = data[i];

        // Email content (text)
        emailIntro = row[0];        // 1st column   "Car@s,oa promoção de..."
        nameDextrano = row[1];        // 2nd column   "NOME "  
        emailResumo1 = row[2];        // 3rd column   "foi aprovada. A promoção será refletida..."
        emailResumo2 = row[3];        // 4th column   "A pessoa deve procurar a Eniandra..."
        currentPosition = row[4];        // 5th column   "Analista de Software Jr/Pl/Sr/..."
        newPosition = row[5];        // 6th column   "Analista de Software Jr/Pl/Sr/..."
        valueAdd = row[8];        // 9th column    Additional benefits ("basket") 

        // Get values (manager, mentor, Dextrano's name etc.)
        emailManager = row[6];        // 6th column
        emailMentor = row[7];        // 7th column
        emailTribe = row[9];        // 10th column 

        // Current and new role, additional benefits
        detailPromotion = 'Cargo Atual...: ' + currentPosition + '\n' + 'Novo Cargo....: ' + newPosition + '\n' + 'Novo valor da cesta (se aplicável) R$: ' + valueAdd + '\n'

        // Email content (full text)
        message = '\n' + emailIntro + ' ' + nameDextrano + '\n' + emailResumo1 + '\n\n' + detailPromotion + '\n' + emailResumo2 + '\n'

        // Email subject and recipients
        subject = 'Comitê de Mérito 2020 - P3';
        emailAddress = "edite.martins@dextra-sw.com, " + " people-matters@dextra-sw.com" + ", " + emailManager + ", " + emailMentor + ", " + emailTribe;

        // Use this alerts to test
        // ui.alert(message);

        // Send the email
        MailApp.sendEmail(emailAddress, subject, message);
    }
}