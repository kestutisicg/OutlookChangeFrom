/* global Office */

function onMessageComposeHandler(event: Office.AddinCommands.Event) {
    const item = Office.context.mailbox.item;
    if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {
        item.from.setAsync("warehouse@icgfarma.eu", (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Nepavyko automatiškai pakeisti siuntėjo: " + asyncResult.error.message);
            }
            event.completed();
        });
    } else {
        event.completed();
    }
}

Office.onReady(() => {
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
});