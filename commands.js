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
                message: "No attachments found on this email.",
                icon: "iconid"
            });
            event.completed();
            return;
        }

        // Fetch content for all attachments
        const attachPromises = attachments.map(async (att) => {
            try {
                const content = await getAttachmentContentAsync(att.id);
                // In Mailbox 1.15, Base64 is supported in displayReplyFormAsync
                if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                    return {
                        type: Office.MailboxEnums.AttachmentType.Base64,
                        name: att.name,
                        base64file: content.content
                    };
                }
            } catch (err) {
                console.warn(`Could not fetch attachment ${att.name}`, err);
            }
            return null; // Ignore unsupported/failed attachments
        });

        const rawAttachments = await Promise.all(attachPromises);
        const replyAttachments = rawAttachments.filter(a => a !== null);

        const formData = {
            htmlBody: "", // This will be prepended to the original email body
            attachments: replyAttachments
        };

        const handleDisplayReplyResult = (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                item.notificationMessages.replaceAsync(notificationId, {
                    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                    message: "Failed to open reply form: " + result.error.message
                });
            } else {
                item.notificationMessages.removeAsync(notificationId);
            }
            event.completed();
        };

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
