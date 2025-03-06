/**
 * Function to find abbreviations in a Word document
 * Uses jQuery for event handling and Office.js for document access
 */
$(document).ready(function () {
    // Attach click event handler to the abbreviation button
    $("#abbrev").on("click", findAbbreviations);
  
    /**
     * Find definitions for the identified abbreviations
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
    function findAbbreviations() {
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
  
        // ----- New code to insert table under "Abbreviations" paragraph -----
        // Search for the paragraph containing "Abbreviations"
        const searchResults = body.search("Abbreviations", { matchWholeWord: true });
        searchResults.load("items");
        await context.sync();
  
        if (searchResults.items.length > 0) {
          const targetParagraph = searchResults.items[0];
          // Apply style "NICE Heading 1" to the found paragraph
          targetParagraph.style = "NICE Heading 1";
  
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
        // ---------------------------------------------------------------------
  
        await context.sync();
        return definitions;
      }).catch(error => {
        console.error("Error finding abbreviations:", error);
      });
    }
  });
  