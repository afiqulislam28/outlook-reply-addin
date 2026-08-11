/* global Office */

Office.onReady(() => {
    // If needed, Office is ready
});

/**
 * Helper function to retrieve the attachment content.
 * @param {string} attachmentId
 * @returns {Promise<any>}
 */
function getAttachmentContentAsync(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}

/**
 * Main logic to handle replying with attachments.
 * @param {Office.AddinCommands.Event} event
 * @param {boolean} isReplyAll
 */
async function replyWithAttachments(event, isReplyAll) {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;
    const notificationId = "replyProgress";

    // Show progress notification
    item.notificationMessages.replaceAsync(notificationId, {
        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
        message: "Fetching attachments and preparing reply..."
    });

    try {
        if (!attachments || attachments.length === 0) {
            item.notificationMessages.replaceAsync(notificationId, {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: "No attachments found.",
                icon: "iconid"
            });
            event.completed();
            return;
        }

        item.notificationMessages.replaceAsync(notificationId, {
            type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
            message: `Step 1: Found ${attachments.length} attachments, downloading...`
        });

        // Fetch content for all attachments (excluding inline images like signatures)
        const attachPromises = attachments.map(async (att) => {
            if (att.isInline) {
                return null; // Skip embedded images
            }
            try {
                const content = await getAttachmentContentAsync(att.id);
                if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64 || content.format === "base64") {
                    return {
                        type: "base64",
                        name: att.name,
                        base64file: content.content
                    };
                }
            } catch (err) {
                // ignore
            }
            return null;
        });

        const rawAttachments = await Promise.all(attachPromises);
        const replyAttachments = rawAttachments.filter(a => a !== null);

        item.notificationMessages.replaceAsync(notificationId, {
            type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
            message: `Step 2: Downloaded ${replyAttachments.length} attachments. Opening form...`
        });

        const formData = {
            htmlBody: ""
        };
        
        // Only include attachments property if we actually have some
        if (replyAttachments.length > 0) {
            formData.attachments = replyAttachments;
        }

        const handleDisplayReplyResult = (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                item.notificationMessages.replaceAsync(notificationId, {
                    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                    message: "Failed to open reply form: " + result.error.message
                });
            } else {
                item.notificationMessages.replaceAsync(notificationId, {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: "Reply opened successfully!",
                    icon: "iconid"
                });
            }
            event.completed();
        };

        // Fallback to ensure event is eventually completed
        setTimeout(() => { 
            item.notificationMessages.replaceAsync(notificationId, {
                type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                message: "Process timed out after 15 seconds."
            });
            event.completed(); 
        }, 15000);

        if (isReplyAll) {
            item.displayReplyAllFormAsync(formData, handleDisplayReplyResult);
        } else {
            item.displayReplyFormAsync(formData, handleDisplayReplyResult);
        }

    } catch (error) {
        console.error(error);
        item.notificationMessages.replaceAsync(notificationId, {
            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            message: "An error occurred while fetching attachments."
        });
        event.completed();
    }
}

function replyWithAttachment(event) {
    replyWithAttachments(event, false);
}

function replyAllWithAttachment(event) {
    replyWithAttachments(event, true);
}

// Register the functions so Outlook can call them from the manifest definitions
Office.actions.associate("replyWithAttachment", replyWithAttachment);
Office.actions.associate("replyAllWithAttachment", replyAllWithAttachment);
