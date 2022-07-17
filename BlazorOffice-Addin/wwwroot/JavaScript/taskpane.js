/** @type {OfficeCore.Range} */
/** @type {Office.Range} */
/** @type {OfficeExtension.Range} */

//Reads the body of the mail in text format
function readMail() {
    try {
        let body = Office.context.mailbox.item.body;
        body.getAsync(Office.CoercionType.Text, function (result) {
            document.getElementById("readLabel").innerHTML = result.value;
        });
    }
    catch (err) {
        document.getElementById("readLabel").innerHTML =
            "An error has occured, please make sure you are running this in an Outlook Client. If you are, please check console to read the error message";
        console.log(err.message);
    }
}

//Replaces the text of the body of the mail with a new text
function writeMail() {
    try {
        let body = Office.context.mailbox.item.body;
        body.setAsync(
            document.getElementById("writeBox").value,
            { coercionType: Office.CoercionType.Text },
            function (result) {
            }
        );
    }
    catch(err) {
        document.getElementById("writeError").innerHTML =
            "An error has occured, please make sure you are composing an email when attempting to use the write mail functionality";
        console.log(err.message);
    }
}
