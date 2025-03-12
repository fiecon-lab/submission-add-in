/**
 * Combined Word document utility script
 * Provides functionality for:
 * 1. Style management and blinding confidential content
 * 2. Finding and formatting abbreviations
 */

$(document).ready(function() {
  // Initialize when Office is ready
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

              // Populate select elements
              const $confidentialSelect = $('#confidential');
              const $blindSelect = $('#blind');

              if ($confidentialSelect.length) {
                  populateSelect($confidentialSelect, 'NICE CIC');
              }

              if ($blindSelect.length) {
                  populateSelect($blindSelect, 'NICE blinded');
              }

          } catch (error) {
              console.error('Error:', error);
          }
      });
  });

  // Attach event handlers for buttons
  $("#blind-btn").on("click", () => tryCatch(blind));
  $("#abbrev").on("click", () => tryCatch(findAbbreviations));
  $("#test").on("click", async () => {
        console.log("start")
        const url = "https://cria-api.fiecon.com/api/generate";
        const apiKey = "0a2e6ef6-4a96-406f-888e-865a8c5a7209";
    
        const requestData = {
        model: "Mistral:7b",
        prompt: "Hello!",
        stream: false,
        };
        console.log("await response")
        const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            APIKey: apiKey,
        },
        body: JSON.stringify(requestData),
        });
    
        console.log(response);
        console.log("end...");
    });

  /**
   * Utility function to handle errors consistently
   */
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

  /**
   * Function to blind confidential content
   */
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
          
          // Save the document as a new file with 'BLINDED' prefix
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

  /**
   * Find definitions for identified abbreviations
   * @param {string} text - The full document text
   * @param {string[]} abbreviations - List of identified abbreviations
   * @return {Object} Map of abbreviations to their definitions
   */
  function findDefinitions(text, abbreviations) {
      const definitionMap = {};

      // Initialize all abbreviations with empty definitions
      abbreviations.forEach(abbr => {
          definitionMap[abbr] = "";
      });

      // Pattern 1: "Full Name (ABBR)" - check if first letters match
      abbreviations.forEach(abbr => {
          try {
              // Look for the pattern: anything followed by the abbreviation in brackets
              const pattern = new RegExp(`([^(]+)\\(${abbr}\\)`, 'gi'); // Case insensitive search
              const matches = [];
              let match;

              // Find all matches
              while ((match = pattern.exec(text)) !== null) {
                  matches.push(match);
              }

              // Check capitalization variations
              for (const m of matches) {
                  const beforeBrackets = m[1].trim();

                  // Extract words, filtering out common connecting words
                  const words = beforeBrackets.split(/[\s-]+/).filter(word =>
                      word.length > 0 &&
                      !['and', 'or', 'the', 'of', 'for', 'in', 'on', 'by', 'to', 'with', 'a', 'an'].includes(word.toLowerCase())
                  );

                  // Get first letter of each word (uppercase for comparison)
                  const firstLetters = words.map(word => word[0].toUpperCase()).join('');

                  // Check if abbreviation matches the first letters (case insensitive)
                  if (firstLetters === abbr.toUpperCase()) {
                      definitionMap[abbr] = beforeBrackets;
                      break;
                  }
              }

              // Find definitions where words don't match exactly the abbreviation order
              // If no definition found yet, try a more flexible approach
              if (!definitionMap[abbr] || definitionMap[abbr] === "") {
                  const parenthesesPattern = new RegExp(`([^(]{3,100})\\(${abbr}\\)`, 'gi');
                  let parenthesesMatch;

                  while ((parenthesesMatch = parenthesesPattern.exec(text)) !== null) {
                      const phraseBeforeBrackets = parenthesesMatch[1].trim();

                      // Match each abbreviation letter with a word in the phrase
                      const phraseWords = phraseBeforeBrackets.split(/[\s-]+/).filter(w => w.length > 0);
                      const abbrLetters = abbr.toUpperCase().split('');

                      // Try to find a "tight" match of consecutive words
                      let bestMatchStart = -1;
                      let bestMatchLength = Infinity;

                      for (let start = 0; start < phraseWords.length; start++) {
                          let abbrPos = 0;
                          let wordPos = start;

                          while (wordPos < phraseWords.length && abbrPos < abbrLetters.length) {
                              const word = phraseWords[wordPos];
                              if (word.length > 0 && word[0].toUpperCase() === abbrLetters[abbrPos]) {
                                  abbrPos++;
                              }
                              wordPos++;
                          }

                          // If we matched all abbreviation letters
                          if (abbrPos === abbrLetters.length) {
                              const matchLength = wordPos - start;
                              if (matchLength < bestMatchLength) {
                                  bestMatchStart = start;
                                  bestMatchLength = matchLength;
                              }
                          }
                      }

                      // If we found a good match
                      if (bestMatchStart !== -1) {
                          const relevantWords = phraseWords.slice(bestMatchStart, bestMatchStart + bestMatchLength);
                          definitionMap[abbr] = relevantWords.join(' ');
                          break;
                      }
                  }
              }
          } catch (error) {
              console.error(`Error processing abbreviation ${abbr}:`, error);
          }
      });

      // Pattern 2: "Abbreviations: ABBR, definition; ABBR2, definition2"
      try {
          const abbreviationSections = text.match(/Abbreviations:([^.]+)/g) || [];

          abbreviationSections.forEach(section => {
              // Remove the "Abbreviations:" prefix
              const content = section.replace(/^Abbreviations:/, '').trim();

              // Split by semicolons
              const pairs = content.split(';');

              pairs.forEach(pair => {
                  // Handle both "ABBR, definition" and "ABBR = definition" formats
                  const pairMatch = pair.match(/^\s*([A-Z0-9-]+)\s*(?:,|=|:)\s*(.+)$/);

                  if (pairMatch) {
                      const [, abbrFromSection, definition] = pairMatch;

                      // Check if this is one of our identified abbreviations
                      if (abbreviations.includes(abbrFromSection)) {
                          definitionMap[abbrFromSection] = definition.trim();
                      }
                  }
              });
          });

          // Look for exact "XYZ = full definition" patterns
          const exactDefinitions = text.match(/\b([A-Z][A-Z0-9-]{1,7})\s*(?:=|is|means|:)\s*["']?([^".;:)]+)["']?/g) || [];
          for (const def of exactDefinitions) {
              const match = def.match(/\b([A-Z][A-Z0-9-]{1,7})\s*(?:=|is|means|:)\s*["']?([^".;:)]+)["']?/);
              if (match && match[1] && match[2]) {
                  const abbr = match[1];
                  const definition = match[2].trim();

                  if (abbreviations.includes(abbr)) {
                      definitionMap[abbr] = definition;
                  }
              }
          }

          // Check for specific patterns like "information for use (IFU)" with exact wording match
          for (const abbr of abbreviations) {
              if (!definitionMap[abbr] || definitionMap[abbr] === "") {
                  // Special case for common abbreviations
                  if (abbr === "IFU") {
                      const ifuMatch = text.match(/information\s+for\s+use\s+\(IFU\)/i);
                      if (ifuMatch) {
                          definitionMap[abbr] = "information for use";
                      }
                  }
              }
          }
      } catch (error) {
          console.error("Error processing abbreviation sections:", error);
      }

      return definitionMap;
  }

  /**
   * Main function to find abbreviations in the document
   */
  async function findAbbreviations() {
      return Word.run(async (context) => {
          // Get the document body
          const body = context.document.body;
          body.load("text");

          await context.sync();

          // Get the full text content
          const text = body.text;

          // Create a Set to store unique abbreviations (removes duplicates)
          const abbreviations = new Set();
          const toExclude = new Set();

          // First find all hyphenated abbreviations to prevent their parts from being included separately
          const hyphenRegex = /\b([A-Z]+)-(\d[A-Z]?)\b/g;
          const hyphenatedAbbreviations = [];
          let hyphenMatch;

          while ((hyphenMatch = hyphenRegex.exec(text)) !== null) {
              const fullMatch = hyphenMatch[0];   // Example: "EQ-5D"
              const beforeHyphen = hyphenMatch[1]; // Example: "EQ"

              hyphenatedAbbreviations.push(fullMatch);
              toExclude.add(beforeHyphen); // Add the first part to exclusion list
          }

          // Identify potential titles (consecutive capitalized words)
          const titleRegex = /\b([A-Z]{2,}(\s+[A-Z]{2,}){2,})\b/g;
          const potentialTitles = [];
          let titleMatch;

          while ((titleMatch = titleRegex.exec(text)) !== null) {
              potentialTitles.push(titleMatch[0]);
          }

          // Add words from titles to exclusions
          potentialTitles.forEach(title => {
              title.split(/\s+/).forEach(word => {
                  if (word.length > 1) {
                      toExclude.add(word);
                  }
              });
          });

          // Common words to exclude
          const commonWords = ["ACRONYM", "AND", "FOR", "THE", "OF", "IN", "TO",
              "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
              "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
              "W", "X", "Y", "Z"];

          // Add common words to exclusions
          commonWords.forEach(word => toExclude.add(word));

          // Find standard acronyms (2+ uppercase letters)
          const acronymMatches = text.match(/\b[A-Z][A-Z]+\b/g) || [];

          // Find mixed-case abbreviations (like CoE, RoI)
          const mixedCaseMatches = text.match(/\b(?:[A-Z](?:[a-z]?[A-Z]){1,}[a-z]?)\b/g) || [];

          // Find acronyms with periods (U.S.A.)
          const periodMatches = text.match(/\b(?:[A-Z]\.){2,}[A-Z]?\b/g) || [];

          // Process all potential abbreviations
          [...acronymMatches, ...mixedCaseMatches, ...periodMatches].forEach(match => {
              // Skip if in exclusion list
              if (toExclude.has(match)) return;

              // Ignore single-character matches
              if (match.length < 2) return;

              // Ignore if it has more than 2 consecutive numbers
              if (/\d{3,}/.test(match)) return;

              // Ignore if it has more than one lowercase letter in a row
              if (/[a-z]{2,}/.test(match)) return;

              // Handle plurals (e.g., RCTs -> RCT)
              if (match.endsWith('s') && match.length > 2) {
                  const singular = match.slice(0, -1);
                  if (singular.match(/^[A-Z]+$/)) {
                      abbreviations.add(singular);
                      return;
                  }
              }

              // Add to set (handles duplicates automatically)
              abbreviations.add(match);
          });

          // Add all hyphenated abbreviations to the final set
          hyphenatedAbbreviations.forEach(abbr => {
              abbreviations.add(abbr);
          });

          // Convert to Array and sort alphabetically
          const sortedAbbreviations = Array.from(abbreviations).sort();

          // Find definitions for each abbreviation
          const definitions = findDefinitions(text, sortedAbbreviations);

          // Normalize capitalization of definitions
          Object.keys(definitions).forEach(abbr => {
              const definition = definitions[abbr];
              if (definition && definition.length > 0) {
                  // Capitalize first letter, leave the rest as is
                  definitions[abbr] = definition.charAt(0).toUpperCase() + definition.slice(1);
              }
          });

          // Log the results
          console.log("Found abbreviations with definitions:");
          console.log(definitions);
          console.log(`Total unique abbreviations found: ${sortedAbbreviations.length}`);

          // Search for the paragraph containing "Abbreviations"
          const searchResults = body.search("Abbreviations", { matchWholeWord: true });
          searchResults.load("items");
          await context.sync();

          if (searchResults.items.length > 0) {
              const targetParagraph = searchResults.items[0];
              // Apply style "NICE Heading 1" to the found paragraph
              // targetParagraph.style = "NICE Heading 1";

              // Build table data: header row + one row per abbreviation
              const tableData = [["Abbreviation", "Definition"]];
              sortedAbbreviations.forEach(abbr => {
                  tableData.push([abbr, definitions[abbr]]);
              });

              // Insert table after the "Abbreviations" paragraph and get the table object
              const table = targetParagraph.insertTable(tableData.length, 2, Word.InsertLocation.after, tableData);
              table.load("id");
              await context.sync();

              // Apply style "NICE Table text" to the table's range
              const tableRange = table.getRange();
              tableRange.style = "NICE Table text";
          }

          await context.sync();
          return definitions;
      });
  }
});