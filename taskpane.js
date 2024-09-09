Office.onReady(function() {
    // Office is ready
    console.log("Office is ready");
});

function sendEmail() {
    // Get the selected name and comments
    let name = document.getElementById('nameSelect').value;
    let comments = document.getElementById('comments').value;

    // Create the email content
    let emailBody = `
        <p>Dear ${name},</p>
        <p>${comments}</p>
        <p>Best regards,</p>
        <p>Your Company</p>
    `;

    // Prepare the email using Office.js API
    Office.context.mailbox.item.body.setAsync(
        emailBody, 
        { coercionType: Office.CoercionType.Html },
        function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                Office.context.mailbox.item.to.setAsync(
                    [{ displayName: name, emailAddress: "example@example.com" }],
                    function(result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Email prepared successfully");
                        } else {
                            console.error("Error setting recipient");
                        }
                    }
                );
            } else {
                console.error("Error setting body content");
            }
        }
    );
}
