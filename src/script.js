Office.onReady(() => {
    // Register the function to run when the document is loaded
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
  
          // Clear any existing options
          $selectElement.empty();
  
          // Add a default "Select a style..." option
          $selectElement.append(
            $('<option>', {
              value: '',
              text: 'Select a style...'
            })
          );
  
          // Add each style to the select element
          styles.items.forEach(style => {
            const $option = $('<option>', {
              value: style.nameLocal,
              text: style.nameLocal,
              selected: style.nameLocal === defaultStyleName
            });
            $selectElement.append($option);
          });
  
          // Add event listener for style changes
          $selectElement.on('change', async function() {
            const selectedStyle = $(this).val();
            if (selectedStyle) {
              await applySelectedStyle(selectedStyle);
            }
          });
  
          // If default style was set, trigger the change event
          if (defaultStyleName && $selectElement.val() === defaultStyleName) {
            $selectElement.trigger('change');
          }
        };
  
        // Get and populate both select elements using jQuery
        const $confidentialSelect = $('#confidential');
        const $blindSelect = $('#blind');
  
        // Try to populate each select if it exists with its default style
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