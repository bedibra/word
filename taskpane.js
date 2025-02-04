/* global Office, Word, diff_match_patch, DIFF_DELETE, DIFF_INSERT, DIFF_EQUAL */

Office.onReady((info) => {
  // Ensure we are in Word
  if (info.host === Office.HostType.Word) {
    console.log("Office is ready in Word.");

    // Get the button by its ID
    const compareBtn = document.getElementById("compareBtn");
    
    if (!compareBtn) {
      console.error("Compare button not found in the DOM!");
    } else {
      // Attach click event
      compareBtn.onclick = compareAndReplace;
      console.log("Compare button event attached.");
    }
  }
});

/**
 * Reads selected text in Word, compares it to the user input,
 * and replaces the selection with diff-based HTML.
 */
async function compareAndReplace() {
  console.log("compareAndReplace triggered.");
  try {
    await Word.run(async (context) => {
      // Get the currently selected text
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync(); // Loads the selection text

      const oldText = selection.text || "";
      const newText = document.getElementById("customText").value || "";

      console.log("Selected text:", oldText);
      console.log("New text:", newText);

      // Create a diff-match-patch instance
      const dmp = new diff_match_patch();
      let diffs = dmp.diff_main(oldText, newText);
      dmp.diff_cleanupSemantic(diffs);

      // Build HTML with <del> / <ins> tags using a standard for loop
      let diffHtml = "";
      for (let i = 0; i < diffs.length; i++) {
        const op = diffs[i][0];
        const data = diffs[i][1];
        // Escape HTML special characters
        const escaped = data
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        
        if (op === DIFF_EQUAL) {
          diffHtml += `<span>${escaped}</span>`;
        } else if (op === DIFF_DELETE) {
          diffHtml += `<del style="color:black;">${escaped}</del>`;
        } else if (op === DIFF_INSERT) {
          diffHtml += `<ins style="color:black;">${escaped}</ins>`;
        }
      }

      console.log("Final diffHtml:", diffHtml);

      // Replace the selection with the diff-based HTML
      selection.insertHtml(diffHtml, Word.InsertLocation.replace);
      await context.sync();

      console.log("Comparison inserted into the document!");
    });
  } catch (error) {
    console.error("Error in compareAndReplace:", error);
  }
}
