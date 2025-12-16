/*
 * Text Formatter Add-in
 * commands.js - Handles ribbon commands that don't show UI
 */

// Initialize Office Add-in
Office.onReady(() => {
    // Office is ready
});

/**
 * Quick format function that can be triggered from the ribbon
 * without opening the task pane.
 * @param {Office.AddinCommands.Event} event - The event object
 */
function quickFormat(event) {
    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.font.name = "Times New Roman";
        selection.font.bold = true;
        await context.sync();
        event.completed();
    }).catch((error) => {
        console.error(error);
        event.completed();
    });
}

// Register the function with Office
Office.actions.associate("quickFormat", quickFormat);


