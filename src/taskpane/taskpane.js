/*
 * Text Formatter Add-in
 * taskpane.js - Main JavaScript file for the Task Pane
 */

// Initialize the Office Add-in when the Office.js library is ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Office is ready, initialize the add-in
        document.getElementById("format-button").onclick = run;
        console.log("Text Formatter Add-in is ready!");
    }
});

/**
 * Formats the currently selected text in the Word document.
 * - Changes the font to "Times New Roman"
 * - Makes the text Bold
 */
async function run() {
    try {
        await Word.run(async (context) => {
            // Get the current selection in the document
            const selection = context.document.getSelection();

            // Load the text property to check if something is selected
            selection.load("text");

            // Synchronize the document state
            await context.sync();

            // Check if there is any text selected
            if (!selection.text || selection.text.trim() === "") {
                showStatus("Please select some text first!", "error");
                return;
            }

            // Apply formatting to the selection
            // Set font name to Times New Roman
            selection.font.name = "Times New Roman";

            // Set font weight to Bold
            selection.font.bold = true;

            // Synchronize to apply the changes
            await context.sync();

            // Show success message
            showStatus("Text formatted successfully! ✓", "success");
            console.log("Formatting applied: Times New Roman, Bold");
        });
    } catch (error) {
        // Handle any errors that occur
        console.error("Error formatting text:", error);
        showStatus("Error: " + error.message, "error");
    }
}

/**
 * Displays a status message to the user
 * @param {string} message - The message to display
 * @param {string} type - The type of message: "success" or "error"
 */
function showStatus(message, type) {
    const statusElement = document.getElementById("status-message");
    
    // Remove any existing classes
    statusElement.classList.remove("hidden", "success", "error");
    
    // Add the appropriate class based on message type
    statusElement.classList.add(type);
    
    // Set the message text
    statusElement.textContent = message;
    
    // Auto-hide success messages after 3 seconds
    if (type === "success") {
        setTimeout(() => {
            statusElement.classList.add("hidden");
        }, 3000);
    }
}


