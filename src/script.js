Office.onReady(() => {
    Word.run(async (context) => {
    try {
        // Get the styles from the document
        let styles = context.document.getStyles();
        context.load(styles);
        await context.sync();

        // Function to populate a select element with styles
        const populateSelect = ($selectElement, defaultStyleName = '') => {
          if (!$selectElement.length) {
              console.error('Select element not found');
              return;
          }
          $selectElement.empty();
          $selectElement.append(
              $('<option>', {
              value: '',
              text: 'Select a style...'
              })
          );

          styles.items.forEach(style => {
              const $option = $('<option>', {
              value: style.nameLocal,
              text: style.nameLocal,
              selected: style.nameLocal === defaultStyleName
              });
              $selectElement.append($option);
          });
        };

        const $confidentialSelect = $('#confidential');
        const $blindSelect = $('#blind');

        if ($confidentialSelect.length) {
        populateSelect($confidentialSelect, 'NICE CIC');
        } else {
        console.warn('Confidential select element not found');
        }

        if ($blindSelect.length) {
        populateSelect($blindSelect, 'NICE blind');
        } else {
        console.warn('Blind select element not found');
        }

    } catch (error) {
        console.error('Error:', error);
    }
    });
});

$("#blind-btn").on("click", () => tryCatch(blind));

async function blind() {
  await Word.run(async (context) => {
    let old_style = $('#confidential').val();
    let new_style = $('#blind').val();

    console.log("Starting style update...");
    let foundCount = 0;

    // Get all paragraphs
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("text, style");
    await context.sync();

    // Process each paragraph
    for (let para of paragraphs.items) {
      if (para.style === old_style) {
        foundCount++;
        // Change paragraph style
        para.style = new_style;
        // Replace text with dashes
        const text = para.text;
        const dashes = "-".repeat(text.trim().length);
        para.insertText(dashes, Word.InsertLocation.replace);
      } else {
        // Search within paragraph for styled content
        const searchResults = para.search("*", { matchWildcards: true });
        searchResults.load(["text", "style", "font"]);
        await context.sync();
        
        for (let result of searchResults.items) {
          if (result.style === old_style || result.font.style === old_style) {
            foundCount++;
            result.style = new_style;
            const dashes = "-".repeat(result.text.trim().length);
            result.insertText(dashes, Word.InsertLocation.replace);
          }
        }
      }
    }

    await context.sync();
    if (foundCount === 0) {
      console.log("No instances of ", old_style, " style found");
    } else {
      console.log(`Updated ${foundCount} instances from ${old_style} to ${new_style} style`);
    }
    
    // Save the document as a new file with 'BLINDED' prefix.
    // Here we use a timestamp to ensure the filename is unique.
    const newFileName = "BLINDED_" + new Date().toISOString() + ".docx";
    Office.context.document.saveAsAsync(newFileName, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Document saved successfully as:", newFileName);
      } else {
        console.error("Failed to save document as new file:", asyncResult.error.message);
      }
    });
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error("Error:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", JSON.stringify(error.debugInfo));
    }
  }
}