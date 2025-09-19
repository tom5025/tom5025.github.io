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
        
        // Get attachments
        mailboxItem.getAttachmentsAsync(function(attachmentsResult) {
            
            // Check if operation was successful
            if (attachmentsResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("❌ Failed to fetch attachments:", attachmentsResult.error);
                eventArgs.completed({ allowEvent: true });
                return;
            }
            
            const attachments = attachmentsResult.value;
            
            // If no attachments are found
            if (attachments.length === 0) {
                console.log("📂 No attachments found. Proceeding with sending email.");
                eventArgs.completed({ allowEvent: true });
                return;
            }
            
            // Filter Excel (.xlsx, .xls) and PDF (.pdf) attachments
            const relevantAttachments = attachments.filter(function(attachment) {
                return attachment.name.endsWith(".xlsx") || 
                       attachment.name.endsWith(".xls") || 
                       attachment.name.endsWith(".pdf");
            });
            
            // If no Excel or PDF attachments are found
            if (relevantAttachments.length === 0) {
                console.log("📂 No Excel or PDF attachments found. Proceeding with sending email.");
                eventArgs.completed({ allowEvent: true });
                return;
            }
            
            // Check if any Excel or PDF attachment contains the project number
            const hasMatchingAttachment = relevantAttachments.some(function(attachment) {
                return attachment.name.includes(projectNumber);
            });
            
            if (hasMatchingAttachment) {
                // Matching attachment found - allow sending
                console.log("✅ Matching Excel or PDF attachment found. Proceeding with sending email.");
                eventArgs.completed({ allowEvent: true });
            } else {
                // No matching attachment - show confirmation dialog
                console.warn("⚠ No matching Excel or PDF attachment found.");
                
                Office.context.ui.displayDialogAsync(
                    "https://colift.de/mailchecker/commands.html",
                    {
                        height: 40,
                        width: 30,
                        displayInIframe: true
                    },
                    function(dialogResult) {
                        
                        // Check if dialog opened successfully
                        if (dialogResult.status !== Office.AsyncResultStatus.Succeeded) {
                            console.error("❌ Failed to open dialog:", dialogResult.error);
                            eventArgs.completed({ allowEvent: true });
                            return;
                        }
                        
                        const dialog = dialogResult.value;
                        
                        // Handle dialog messages
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(messageArgs) {
                            console.log("📩 Dialog Response:", messageArgs.message);
                            dialog.close();
                            
                            if (messageArgs.message === "confirmed") {
                                // User clicked OK - allow sending
                                console.log("✅ User clicked OK. Proceeding with sending email.");
                                eventArgs.completed({ allowEvent: true });
                            } else {
                                // User clicked Cancel - block sending
                                console.log("❌ User clicked Cancel. Blocking email.");
                                eventArgs.completed({ allowEvent: false });
                            }
                        });
                    }
                );
            }
        });
    });
});

/* 
ADD-IN LOGIC:

1. **Intercepts email sending** via Office.actions.associate
2. **Extracts project number** from subject (format: "Title (PROJECT123)")
3. **Checks Excel and PDF attachments** (.xlsx/.xls/.pdf)
4. **Controls matching** between project number and file names
5. **Shows confirmation** if no match is found
6. **Allows or blocks sending** based on user response

WORKFLOW:
- Subject: "Monthly Report (PROJ001)"
- Attachment: "PROJ001_data.xlsx" ✅ → Send allowed
- Attachment: "PROJ001_report.pdf" ✅ → Send allowed
- Attachment: "other_file.xlsx" ⚠️ → Confirmation required

*/