/* global Office */

Office.onReady(() => {
    // Office is ready
});

/**
 * Safely encodes a UTF-8 string to Base64 (handles Unicode characters and large text).
 * @param {string} str
 * @returns {string}
 */
function utf8ToBase64(str) {
    if (typeof TextEncoder !== "undefined") {
        const bytes = new TextEncoder().encode(str);
        let binary = "";
        const len = bytes.length;
        const chunkSize = 8192;
        for (let i = 0; i < len; i += chunkSize) {
            binary += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(i + chunkSize, len)));
        }
        return btoa(binary);
    }
    return btoa(unescape(encodeURIComponent(str)));
}

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
 * Main logic to handle replying with attachments silently without confusing step banners.
 * @param {Office.AddinCommands.Event} event
 * @param {boolean} isReplyAll
 */
async function replyWithAttachments(event, isReplyAll) {
    const item = Office.context.mailbox.item;
    const notificationId = "replyProgress";

    // Clear any previous notification message from earlier actions
    try {
        item.notificationMessages.removeAsync(notificationId);
    } catch (e) {
        // ignore
    }

    let timeoutTimer = null;
    let completed = false;
    const completeEvent = () => {
        if (!completed) {
            completed = true;
            if (timeoutTimer) {
                clearTimeout(timeoutTimer);
                timeoutTimer = null;
            }
            try {
                event.completed();
            } catch (e) {
                // ignore
            }
        }
    };

    // Safety fallback to complete the event if host hangs
    timeoutTimer = setTimeout(() => {
        completeEvent();
    }, 15000);

    try {
        const attachments = item.attachments || [];
        let replyAttachments = [];

        if (attachments.length > 0) {
            // Process all attachments (excluding inline images like signatures)
            const attachPromises = attachments.map(async (att) => {
                if (att.isInline) {
                    return null; // Skip embedded images (native reply preserves them in quoted body)
                }

                // 1. Outlook Item Attachments (attached emails / .msg / appointments)
                // Attach directly as an Outlook Item by ID so it preserves the native Outlook item format and icon
                if (att.attachmentType === Office.MailboxEnums.AttachmentType.Item || att.attachmentType === "item") {
                    return {
                        type: Office.MailboxEnums.AttachmentType.Item,
                        name: att.name,
                        itemId: att.id,
                        id: att.id
                    };
                }

                // 2. Standard File Attachments (PDF, Excel, Word, Images, etc.)
                try {
                    const content = await getAttachmentContentAsync(att.id);

                    if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64 || content.format === "base64") {
                        return {
                            type: "base64",
                            name: att.name,
                            base64file: content.content
                        };
                    }

                    // Fallback for item attachments if not captured by attachmentType
                    if (content.format === Office.MailboxEnums.AttachmentContentFormat.Eml || content.format === "eml") {
                        return {
                            type: Office.MailboxEnums.AttachmentType.Item,
                            name: att.name,
                            itemId: att.id,
                            id: att.id
                        };
                    }

                    if (content.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar || content.format === "iCalendar") {
                        return {
                            type: "base64",
                            name: att.name,
                            base64file: utf8ToBase64(content.content)
                        };
                    }
                } catch (err) {
                    console.warn(`Could not fetch attachment content for "${att.name}":`, err);
                    // Fallback: If it's an item attachment and content retrieval failed, try attaching by itemId
                    if (att.attachmentType === Office.MailboxEnums.AttachmentType.Item || att.attachmentType === "item") {
                        return {
                            type: Office.MailboxEnums.AttachmentType.Item,
                            name: att.name,
                            itemId: att.id,
                            id: att.id
                        };
                    }
                }
                return null;
            });

            const rawAttachments = await Promise.all(attachPromises);
            replyAttachments = rawAttachments.filter(a => a !== null);
        }

        const formData = {
            htmlBody: ""
        };

        // Only attach if there are non-inline attachments
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
                // Ensure notification bar is clean and clear
                item.notificationMessages.removeAsync(notificationId);
            }
            completeEvent();
        };

        if (isReplyAll) {
            item.displayReplyAllFormAsync(formData, handleDisplayReplyResult);
        } else {
            item.displayReplyFormAsync(formData, handleDisplayReplyResult);
        }

        // Dismiss Outlook's native "Working on your request" message immediately
        completeEvent();

    } catch (error) {
        console.error("Error in replyWithAttachments:", error);
        item.notificationMessages.replaceAsync(notificationId, {
            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            message: "An error occurred while opening the reply."
        });
        completeEvent();
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
