/** @type {OfficeCore.Range} */
/** @type {Office.Range} */
/** @type {OfficeExtension.Range} */

//Reads the body of the mail in text format
function readMail() {
    let body = Office.context.mailbox.item.body;

    body.getAsync(Office.CoercionType.Text, function (result) {
        document.getElementById("readLabel").innerHTML = result.value;
    });
}

//Replaces the text of the body of the mail with a new text
function writeMail() {
    let body = Office.context.mailbox.item.body;
    body.setAsync(
        document.getElementById("writeBox").value,
        { coercionType: Office.CoercionType.Text },
        function (result) {
            return;
        }
    );
}
