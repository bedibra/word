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
  try {
    await Word.run(async context => {
      // 1. Get the current selection and load its font properties
      const selection = context.document.getSelection();
      selection.load("text, font/name, font/size, font/color");
      await context.sync();

      const oldText = selection.text;
      const newText = document.getElementById("customText").value;

      // 2. Generate diff HTML (using diff-match-patch)
      const diffHtml = generateDiffHtml(oldText, newText);

      // 3. Insert the diff HTML, replacing the current selection
      selection.insertHtml(diffHtml, Word.InsertLocation.replace);
      await context.sync();

      // 4. Reapply the original formatting to the newly inserted content.
      //    Note: This applies the same font properties to all content,
      //    so if you have inline styles in your diff markup (like red/green colors)
      //    for insertions/deletions, you might need to adjust accordingly.
      const newRange = context.document.getSelection();
      newRange.font.name = selection.font.name;
      newRange.font.size = selection.font.size;
      newRange.font.color = selection.font.color;
      await context.sync();

      console.log("Comparison inserted with original formatting!");
    });
  } catch (error) {
    console.error("Error performing compare and replace:", error);
  }
}
