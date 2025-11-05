// Outlook Add-in Commands - Attachment validation before email sending

// Office initialization (empty function in this case)
Office.onReady(function() {
    // Empty initialization - add-in is ready
});

// Associate an action named "action" to intercept email sending
Office.actions.associate("action", function(eventArgs) {
    
    // Get the current email item
    const mailboxItem = Office.context.mailbox.item;
    
    // Check if the item exists
    if (!mailboxItem) {
        console.error("❌ No item found.");
        eventArgs.completed({ allowEvent: true });
        return;
    }
    
    // Get email subject asynchronously
    mailboxItem.subject.getAsync(function(subjectResult) {
        
        // Check if operation was successful
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("❌ Failed to get email subject:", subjectResult.error);
            eventArgs.completed({ allowEvent: true });
            return;
        }
        
        // Extract the subject
        const emailSubject = subjectResult.value || "";
        console.log("📌 Email Subject:", emailSubject);

        // Search for a project number between parentheses in the subject
        const projectMatch = emailSubject.match(/\(([^)]+)\)/);
        const projectNumber = projectMatch ? projectMatch[1] : null;

        // If no project number is found
        if (!projectNumber) {
            console.warn("⚠ No project number found in subject.");
            eventArgs.completed({ allowEvent: true });
            return;
        }

        console.log("📌 Project Number from subject:", projectNumber);

        // Get email body to check for Zoom links
        mailboxItem.body.getAsync(Office.CoercionType.Text, function(bodyResult) {

            // Check if operation was successful
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("❌ Failed to get email body:", bodyResult.error);
                eventArgs.completed({ allowEvent: true });
                return;
            }

            const emailBody = bodyResult.value || "";

            // Check if email contains a Zoom recording link
            const hasZoomLink = /zoom\.us/i.test(emailBody);

            if (hasZoomLink) {
                console.log("🎥 Zoom recording link detected. Checking if code from subject is in body...");

                // Check if the project number from subject appears in the body
                if (!emailBody.includes(projectNumber)) {
                    console.warn("⚠ Code from subject line not found in email body with Zoom link.");

                    // Show error dialog - code not found in body
                    Office.context.ui.displayDialogAsync(
                        //"http:s//colift.de/mailchecker/commands.html",
                        "http://localhost:3000/commands.html",
                        {
                            height: 40,
                            width: 30,
                            displayInIframe: true
                        },
                        function(dialogResult) {
                            if (dialogResult.status !== Office.AsyncResultStatus.Succeeded) {
                                console.error("❌ Failed to open dialog:", dialogResult.error);
                                eventArgs.completed({ allowEvent: false });
                                return;
                            }

                            const dialog = dialogResult.value;

                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(messageArgs) {
                                console.log("📩 Dialog Response:", messageArgs.message);
                                dialog.close();

                                if (messageArgs.message === "confirmed") {
                                    console.log("✅ User confirmed to send anyway.");
                                    // Continue with attachment validation
                                    checkAttachments();
                                } else {
                                    console.log("❌ User cancelled. Blocking email.");
                                    eventArgs.completed({ allowEvent: false });
                                }
                            });
                        }
                    );
                    return;
                }

                console.log("✅ Code from subject line found in body.");
            }            
        });        
    });
});

/*
ADD-IN LOGIC:

1. **Intercepts email sending** via Office.actions.associate
2. **Extracts project number** from subject (format: "Title (PROJECT123)")
3. **Checks for Zoom recording links** (zoom.us in email body)
   - If Zoom link found: Verifies project number from subject exists in email body
   - Shows error if code missing from body
4. **Checks Excel and PDF attachments** (.xlsx/.xls/.pdf)
5. **Controls matching** between project number and file names
6. **Shows confirmation** if no match is found
7. **Allows or blocks sending** based on user response

WORKFLOW:
- Subject: "Monthly Report (PROJ001)"
- Body with zoom.us link containing "PROJ001" ✅ → Continue validation
- Body with zoom.us link without "PROJ001" ⚠️ → Confirmation required

*/