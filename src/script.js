Office.onReady(() => {
  Word.run(async (context) => {
      try {
          // Get all styles from the document
          let allStyles = context.document.getStyles();
          context.load(allStyles, 'items/nameLocal, items/type, items/builtin');
          await context.sync();

          // Filter styles to match Word's UI display
          const visibleStyles = allStyles.items.filter(style => {
              // Only include styles that would typically appear in Word's UI
              return (
                  // Focus on paragraph styles (most common in UI)
                  (style.type === Word.StyleType.paragraph || 
                   style.type === Word.StyleType.character) &&
                  // Filter out styles that start with special characters
                  !style.nameLocal.startsWith('_') &&
                  // Include styles that match naming patterns shown in your UI
                  (style.nameLocal.startsWith('NICE') || 
                   style.nameLocal.includes('Heading') ||
                   style.nameLocal.includes('Title') ||
                   style.builtin) // Include built-in styles
              );
          });

          // Function to populate select elements
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

              // Sort styles to match UI order
              visibleStyles.sort((a, b) => a.nameLocal.localeCompare(b.nameLocal));
              
              visibleStyles.forEach(style => {
                  const $option = $('<option>', {
                      value: style.nameLocal,
                      text: style.nameLocal,
                      selected: style.nameLocal === defaultStyleName
                  });
                  $selectElement.append($option);
              });
          };

          // Populate your select elements
          const $confidentialSelect = $('#confidential');
          const $blindSelect = $('#blind');

          if ($confidentialSelect.length) {
              populateSelect($confidentialSelect, 'NICE CIC');
          }

          if ($blindSelect.length) {
              populateSelect($blindSelect, 'NICE blind');
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