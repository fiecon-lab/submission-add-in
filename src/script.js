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
                  // Focus on paragraph styles (most common in UI)
                  return (
                      (style.type === Word.StyleType.paragraph || 
                       style.type === Word.StyleType.character) &&
                      !style.nameLocal.startsWith('_') &&
                      (style.nameLocal.startsWith('NICE') || 
                       style.nameLocal.includes('Heading') ||
                       style.nameLocal.includes('Title') ||
                       style.builtin)
                  );
              });

              // Populate select elements
              populateStyleSelects(visibleStyles);
          } catch (error) {
              console.error('Error:', error);
          }
      });
  });

  // Attach event handlers for buttons
  $("#blind-btn").on("click", () => tryCatch(blind));
  $("#abbrev").on("click", () => tryCatch(findAbbreviations));
  $("#abbrev-table").on("click", () => tryCatch(findTableAbbreviations));

  // Store known abbreviations globally
  const knownAbbreviations = new Set();
  
  // Initialize with industry-specific abbreviations
  const industryAbbreviations = {
    "HEOR": "Health Economics and Outcomes Research",
    "HTA": "Health Technology Assessment/Appraisal",
    "ICER": "Incremental Cost-Effectiveness Ratio",
    "QALY": "Quality-Adjusted Life Year",
    "DALY": "Disability-Adjusted Life Year",
    "CER": "Cost-Effectiveness Ratio",
    "BIA": "Budget Impact Analysis",
    "CUA": "Cost-Utility Analysis",
    "CEA": "Cost-Effectiveness Analysis/Comparative Effectiveness Analysis",
    "CBA": "Cost-Benefit Analysis",
    "CMA": "Cost-Minimisation Analysis",
    "PRO": "Patient-Reported Outcome",
    "PROM": "Patient-Reported Outcome Measure",
    "RWE": "Real-World Evidence",
    "RWD": "Real-World Data",
    "NICE": "National Institute for Health and Care Excellence (UK)",
    "SMC": "Scottish Medicines Consortium",
    "CADTH": "Canadian Agency for Drugs and Technologies in Health",
    "FDA": "Food and Drug Administration (US)",
    "EMA": "European Medicines Agency",
    "MHRA": "Medicines and Healthcare products Regulatory Agency (UK)",
    "NHS": "National Health Service (UK)",
    "CMS": "Centers for Medicare & Medicaid Services (US)",
    "P&R": "Pricing and Reimbursement",
    "TPP": "Target Product Profile",
    "VBP": "Value-Based Pricing",
    "MEA": "Managed Entry Agreement",
    "PAS": "Patient Access Scheme (UK)",
    "CED": "Coverage with Evidence Development",
    "HCP": "Healthcare Professional",
    "KOL": "Key Opinion Leader",
    "GPP": "Global Pricing Paper",
    "SLR": "Systematic Literature Review",
    "MA": "Meta-Analysis",
    "NMA": "Network Meta-Analysis",
    "ITC": "Indirect Treatment Comparison",
    "MAIC": "Matching-Adjusted Indirect Comparison",
    "STC": "Simulated Treatment Comparison",
    "ITT": "Intention-to-Treat",
    "PP": "Per Protocol",
    "AE": "Adverse Event",
    "SAE": "Serious Adverse Event",
    "RCT": "Randomised Controlled Trial",
    "SoC": "Standard of Care",
    "LoT": "Line of Therapy",
    "QoL": "Quality of Life",
    "HRQoL": "Health-Related Quality of Life",
    "DRG": "Diagnosis-Related Group",
    "HRG": "Healthcare Resource Group",
    "ALOS": "Average Length of Stay",
    "MCO": "Managed Care Organization (US)",
    "PBM": "Pharmacy Benefit Manager (US)",
    "IDN": "Integrated Delivery Network (US)",
    "CCG": "Clinical Commissioning Group (UK, historical)",
    "ICS": "Integrated Care System (UK)",
    "ICB": "Integrated Care Board (UK)",
    "T1D": "Type 1 Diabetes",
    "T2D": "Type 2 Diabetes",
    "DM": "Diabetes Mellitus",
    "GDM": "Gestational Diabetes Mellitus",
    "HbA1c": "Haemoglobin A1c (glycated haemoglobin)",
    "FPG": "Fasting Plasma Glucose",
    "NSCLC": "Non-Small Cell Lung Cancer",
    "SCLC": "Small Cell Lung Cancer",
    "mCRC": "Metastatic Colorectal Cancer",
    "HCC": "Hepatocellular Carcinoma",
    "RCC": "Renal Cell Carcinoma",
    "BC": "Breast Cancer",
    "mBC": "Metastatic Breast Cancer",
    "TNBC": "Triple-Negative Breast Cancer",
    "PCa": "Prostate Cancer",
    "mPCa": "Metastatic Prostate Cancer",
    "NHL": "Non-Hodgkin Lymphoma",
    "MMy": "Multiple Myeloma",
    "AML": "Acute Myeloid Leukaemia",
    "CLL": "Chronic Lymphocytic Leukaemia",
    "PFS": "Progression-Free Survival",
    "OS": "Overall Survival",
    "ORR": "Objective Response Rate",
    "DOR": "Duration of Response",
    "CR": "Complete Response",
    "PR": "Partial Response",
    "AD": "Alzheimer's Disease",
    "MCI": "Mild Cognitive Impairment",
    "PD": "Parkinson's Disease",
    "MS": "Multiple Sclerosis",
    "RRMS": "Relapsing-Remitting Multiple Sclerosis",
    "PPMS": "Primary Progressive Multiple Sclerosis",
    "SPMS": "Secondary Progressive Multiple Sclerosis",
    "ALS": "Amyotrophic Lateral Sclerosis",
    "HD": "Huntington's Disease",
    "MMSE": "Mini-Mental State Examination",
    "CVD": "Cardiovascular Disease",
    "CHF": "Congestive Heart Failure",
    "MI": "Myocardial Infarction",
    "ACS": "Acute Coronary Syndrome",
    "AF": "Atrial Fibrillation",
    "HTN": "Hypertension",
    "PAD": "Peripheral Arterial Disease",
    "MACE": "Major Adverse Cardiovascular Events",
    "RA": "Rheumatoid Arthritis",
    "PsA": "Psoriatic Arthritis",
    "AS": "Ankylosing Spondylitis",
    "SLE": "Systemic Lupus Erythematosus",
    "IBD": "Inflammatory Bowel Disease",
    "CD": "Crohn's Disease",
    "UC": "Ulcerative Colitis",
    "NASH": "Non-Alcoholic Steatohepatitis",
    "COPD": "Chronic Obstructive Pulmonary Disease",
    "CKD": "Chronic Kidney Disease",
    "ESRD": "End-Stage Renal Disease",
    "HIV": "Human Immunodeficiency Virus",
    "HCV": "Hepatitis C Virus",
    "HBV": "Hepatitis B Virus",
    "TB": "Tuberculosis",
    "CE": "Cost-Effectiveness",
    "CONSORT": "Consolidated Standards of Reporting Trials",
    "DPD": "Drug Pricing Database",
    "ECDRP": "European Commission Decision Reliance Procedure",
    "EQ-5D": "EuroQol 5-Dimension",
    "ID": "Identification",
    "IFU": "Information for Use",
    "LYG": "Life Years Gained",
    "NHB": "Net Health Benefit",
    "PbR": "Payment by Results",
    "RIS": "Research Information Systems",
    "STA": "Single Technology Appraisal",
    "SmPC": "Summary of Product Characteristics",
    "TA": "Technology Appraisal",
    "CEM": "Cost-Effectiveness Model",
    "BIM": "Budget Impact Model",
    "PSA": "Probabilistic Sensitivity Analysis",
    "DSA": "Deterministic Sensitivity Analysis",
    "OWSA": "One-Way Sensitivity Analysis",
    "TWSA": "Two-Way Sensitivity Analysis",
    "PSM": "Partitioned Survival Model",
    "STM": "State Transition Model",
    "DES": "Discrete Event Simulation",
    "MM": "Markov Model",
    "TTO": "Time Trade-Off",
    "SG": "Standard Gamble",
    "WTP": "Willingness To Pay",
    "PSS": "Personal Social Services",
    "DCE": "Discrete Choice Experiment",
    "VOI": "Value of Information",
    "EVPI": "Expected Value of Perfect Information",
    "EVPPI": "Expected Value of Partial Perfect Information",
    "EVSI": "Expected Value of Sample Information",
    "INMB": "Incremental Net Monetary Benefit",
    "ICUR": "Incremental Cost-Utility Ratio",
    "GDP": "Gross Domestic Product",
    "HR": "Hazard Ratio",
    "OR": "Odds Ratio",
    "RR": "Relative Risk",
    "CI": "Confidence Interval",
    "CrI": "Credible Interval",
    "AIC": "Akaike Information Criterion",
    "BIC": "Bayesian Information Criterion",
    "MSM": "Multi-State Model",
    "DAM": "Decision Analytic Model",
    "HUI": "Health Utilities Index",
    "SF-6D": "Short-Form Six-Dimension",
    "VAS": "Visual Analogue Scale",
    "AUC": "Area Under the Curve",
    "K-M": "Kaplan-Meier",
    "UK": "United Kingdom",
    "USA": "United States of America",
    "US": "United States",
    "EU": "European Union",
    "EU-5": "France, Germany, Italy, Spain, United Kingdom",
    "EU-4": "France, Germany, Italy, Spain",
    "LATAM": "Latin America",
    "APAC": "Asia-Pacific",
    "EMEA": "Europe, Middle East, and Africa",
    "ROW": "Rest of World",
    "FR": "France",
    "DE": "Germany",
    "IT": "Italy",
    "ES": "Spain",
    "JP": "Japan",
    "CN": "China",
    "AU": "Australia",
    "CA": "Canada",
    "CH": "Switzerland",
    "SE": "Sweden",
    "DK": "Denmark",
    "NO": "Norway",
    "FI": "Finland",
    "NL": "Netherlands",
    "BE": "Belgium",
    "AT": "Austria",
    "IE": "Ireland",
    "PT": "Portugal",
    "GR": "Greece",
    "BR": "Brazil",
    "MX": "Mexico",
    "RU": "Russia",
    "IN": "India",
    "KR": "South Korea",
    "TW": "Taiwan",
    "BRICS": "Brazil, Russia, India, China, South Africa",
    "PBAC": "Pharmaceutical Benefits Advisory Committee (Australia)",
    "MSAC": "Medical Services Advisory Committee (Australia)",
    "ICD": "International Classification of Diseases"
  };

  // Initialize known abbreviations with industry-specific ones
  Object.keys(industryAbbreviations).forEach(abbr => knownAbbreviations.add(abbr));

  // Also store the definitions for later use
  const knownDefinitions = new Map(Object.entries(industryAbbreviations));

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
   * Helper function to populate style select elements
   */
  function populateStyleSelects(styles) {
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
          styles.sort((a, b) => a.nameLocal.localeCompare(b.nameLocal));
          
          styles.forEach(style => {
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
      }

      if ($blindSelect.length) {
          populateSelect($blindSelect, 'NICE blinded');
      }
  }

  /**
   * Progress indicator utility
   */
  const ProgressIndicator = {
      currentOperation: null,

      start: function(operation) {
          this.currentOperation = operation;
          $('#modal-overlay .status-header span').text(operation);
          $('#modal-progress-bar').css('width', '0%');
          $('#modal-progress-status').text('Preparing...');
          $('#modal-overlay').show();
          $('button, select').prop('disabled', true);
      },

      update: function(percent, status) {
          $('#modal-progress-bar').css('width', percent + '%');
          $('#modal-progress-status').text(status);
      },

      finish: function(status = 'Complete', delay = 3000) {
          this.update(100, status);
          setTimeout(() => {
              $('#modal-overlay').hide();
              $('button, select').prop('disabled', false);
              this.currentOperation = null;
          }, delay);
      },

      error: function(errorMessage) {
          this.update(100, `Error: ${errorMessage}`);
          setTimeout(() => {
              $('#modal-overlay').hide();
              $('button, select').prop('disabled', false);
              this.currentOperation = null;
          }, 5000);
      }
  };

  /**
   * Function to blind confidential content
   */
  async function blind() {
      ProgressIndicator.start('Blinding document');

      await Word.run(async (context) => {
          try {
              let old_style = $('#confidential').val();
              let new_style = $('#blind').val();
              
              console.log("Starting style update...");
              let foundCount = 0;
              const PARAGRAPH_BATCH_SIZE = 20;
              const CHARACTER_BATCH_SIZE = 100;

              const body = context.document.body;
              body.load("text");
              await context.sync();

              ProgressIndicator.update(5, "Document loaded. Starting blinding process...");
              console.log("Document loaded. Starting blinding process...");
              
              // Process paragraphs with the target style
              foundCount += await processParagraphStyles(context, body, old_style, new_style, PARAGRAPH_BATCH_SIZE);
              
              // Process character-level styling
              foundCount += await processCharacterStyles(context, body, old_style, new_style, CHARACTER_BATCH_SIZE);
              
              // Final report
              const finalMessage = foundCount === 0 
                  ? "No instances of style " + old_style + " found"
                  : `Blinding complete. Updated ${foundCount} total instances from ${old_style} to ${new_style} style`;
              
              ProgressIndicator.finish(finalMessage);
              console.log(finalMessage);
          } catch (error) {
              console.error("Error during blinding:", error);
              ProgressIndicator.error("Failed to complete blinding process");
          }
      });
  }

  /**
   * Helper function to process paragraph styles
   */
  async function processParagraphStyles(context, body, oldStyle, newStyle, batchSize) {
      let foundCount = 0;
      ProgressIndicator.update(10, "Processing paragraphs with style: " + oldStyle);
      
      const paragraphsWithStyle = body.paragraphs;
      paragraphsWithStyle.load("items");
      await context.sync();
      
      const totalParagraphs = paragraphsWithStyle.items.length;
      ProgressIndicator.update(15, `Processing ${totalParagraphs} paragraphs in batches of ${batchSize}`);
      
      for (let i = 0; i < totalParagraphs; i += batchSize) {
          const batchEnd = Math.min(i + batchSize, totalParagraphs);
          const batchNumber = Math.floor(i/batchSize) + 1;
          const totalBatches = Math.ceil(totalParagraphs/batchSize);
          
          const progressPercent = 15 + Math.floor((i / totalParagraphs) * 25);
          ProgressIndicator.update(progressPercent, `Processing paragraph batch ${batchNumber}/${totalBatches}`);
          
          for (let j = i; j < batchEnd; j++) {
              paragraphsWithStyle.items[j].load("style, text");
          }
          await context.sync();
          
          for (let j = i; j < batchEnd; j++) {
              const para = paragraphsWithStyle.items[j];
              if (para.style === oldStyle) {
                  foundCount++;
                  para.style = newStyle;
                  const text = para.text;
                  const dashes = "-".repeat(text.trim().length);
                  para.insertText(dashes, Word.InsertLocation.replace);
              }
          }
          
          await context.sync();
      }
      
      ProgressIndicator.update(40, `Found ${foundCount} paragraphs with style ${oldStyle}`);
      return foundCount;
  }

  /**
   * Helper function to process character styles
   */
  async function processCharacterStyles(context, body, oldStyle, newStyle, batchSize) {
      let foundCount = 0;
      ProgressIndicator.update(45, "Processing character-level styling");
      
      const contentRanges = body.search("*", { matchWildcards: true });
      contentRanges.load("items");
      await context.sync();
      
      const totalRanges = contentRanges.items.length;
      ProgressIndicator.update(50, `Found ${totalRanges} content ranges to check for character styling`);
      
      for (let i = 0; i < totalRanges; i += batchSize) {
          const batchEnd = Math.min(i + batchSize, totalRanges);
          const progressPercent = 50 + Math.floor((i / totalRanges) * 45);
          const batchNumber = Math.floor(i/batchSize) + 1;
          const totalBatches = Math.ceil(totalRanges/batchSize);
          
          ProgressIndicator.update(progressPercent, `Processing character styles batch ${batchNumber}/${totalBatches}`);
          
          for (let j = i; j < batchEnd; j++) {
              contentRanges.items[j].load("text, style, font");
          }
          await context.sync();
          
          let batchFoundCount = 0;
          for (let j = i; j < batchEnd; j++) {
              const range = contentRanges.items[j];
              if (range.style === oldStyle || (range.font && range.font.style === oldStyle)) {
                  foundCount++;
                  batchFoundCount++;
                  range.style = newStyle;
                  const text = range.text;
                  const dashes = "-".repeat(text.trim().length);
                  range.insertText(dashes, Word.InsertLocation.replace);
              }
          }
          
          await context.sync();
          
          if (batchFoundCount > 0 && (batchNumber % 5 === 0 || batchNumber === totalBatches)) {
              ProgressIndicator.update(progressPercent, `Found ${batchFoundCount} styled ranges in batch ${batchNumber}/${totalBatches}`);
          }
      }
      
      ProgressIndicator.update(95, `Found ${foundCount} ranges with style ${oldStyle}`);
      return foundCount;
  }

  /**
   * Core abbreviation detection logic
   */
  function detectAbbreviations(text, options = { isTable: false, useKnownAbbreviations: true }) {
      const abbreviations = new Set();
      const toExclude = new Set();
      const contextMap = new Map(); // Track where each abbreviation appears
      const allPositions = new Map(); // Track all positions of all potential abbreviations

      // First, add any known abbreviations found in the text
      if (options.useKnownAbbreviations) {
          knownAbbreviations.forEach(known => {
              const regex = new RegExp(`\\b${known}\\b`, 'g');
              let match;
              while ((match = regex.exec(text)) !== null) {
                  contextMap.set(match.index, known);
                  if (!allPositions.has(known)) {
                      allPositions.set(known, []);
                  }
                  allPositions.get(known).push(match.index);
              }
          });
      }

      // Find hyphenated abbreviations
      const hyphenRegex = /\b([A-Z]+)-(\d[A-Z]?)\b/g;
      let hyphenMatch;

      while ((hyphenMatch = hyphenRegex.exec(text)) !== null) {
          const fullMatch = hyphenMatch[0];
          const beforeHyphen = hyphenMatch[1];
          toExclude.add(beforeHyphen);
          contextMap.set(hyphenMatch.index, fullMatch);
          if (!allPositions.has(fullMatch)) {
              allPositions.set(fullMatch, []);
          }
          allPositions.get(fullMatch).push(hyphenMatch.index);
      }

      // Exclude potential titles only for main document and tables without headers
      if (!options.isTable || options.isTable && !options.hasHeaders) {
          const titleRegex = /\b([A-Z]{2,}(\s+[A-Z]{2,}){2,})\b/g;
          let titleMatch;
          while ((titleMatch = titleRegex.exec(text)) !== null) {
              titleMatch[0].split(/\s+/).forEach(word => {
                  if (word.length > 1) toExclude.add(word);
              });
          }
      }

      // Common words to exclude
      const commonWords = ["ACRONYM", "AND", "FOR", "THE", "OF", "IN", "TO",
          "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
          "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
          "W", "X", "Y", "Z"];
      commonWords.forEach(word => toExclude.add(word));

      // Find all potential abbreviations
      const patterns = [
          /\b[A-Z][A-Z]+\b/g,                    // Standard acronyms
          /\b(?:[A-Z](?:[a-z]?[A-Z]){1,}[a-z]?)\b/g,  // Mixed-case
          /\b(?:[A-Z]\.){2,}[A-Z]?\b/g,         // With periods
          /\b[A-Z][A-Z0-9-]+\b/g                // With numbers and hyphens
      ];

      patterns.forEach(pattern => {
          let match;
          while ((match = pattern.exec(text)) !== null) {
              const abbr = match[0];
              if (toExclude.has(abbr)) continue;
              if (abbr.length < 2) continue;
              if (/\d{3,}/.test(abbr)) continue;
              if (/[a-z]{2,}/.test(abbr)) continue;

              // Track this potential abbreviation's position
              if (!allPositions.has(abbr)) {
                  allPositions.set(abbr, []);
              }
              allPositions.get(abbr).push(match.index);
          }
      });

      // Now analyze all found positions to determine valid abbreviations
      allPositions.forEach((positions, abbr) => {
          let isValid = true;
          
          positions.forEach(pos => {
              // Check if this position is part of a larger abbreviation
              contextMap.forEach((existingAbbr, existingPos) => {
                  if (existingPos <= pos && 
                      existingPos + existingAbbr.length >= pos + abbr.length &&
                      existingAbbr !== abbr) {
                      isValid = false;
                  }
              });
          });

          if (isValid) {
              if (abbr.endsWith('s') && abbr.length > 2) {
                  const singular = abbr.slice(0, -1);
                  if (singular.match(/^[A-Z]+$/)) {
                      abbreviations.add(singular);
                      positions.forEach(pos => contextMap.set(pos, singular));
                  }
              } else {
                  abbreviations.add(abbr);
                  positions.forEach(pos => contextMap.set(pos, abbr));
              }
          }
      });

      // Update known abbreviations with new findings
      if (options.useKnownAbbreviations) {
          abbreviations.forEach(abbr => knownAbbreviations.add(abbr));
      }

      return Array.from(abbreviations).sort();
  }

  /**
   * Find definitions for identified abbreviations
   */
  function findDefinitions(text, abbreviations) {
      const definitionMap = {};
      abbreviations.forEach(abbr => {
          // First check if we have a known definition
          if (knownDefinitions.has(abbr)) {
              definitionMap[abbr] = knownDefinitions.get(abbr);
              return;
          }
          definitionMap[abbr] = "";
      });

      // Pattern 1: "Full Name (ABBR)"
      abbreviations.forEach(abbr => {
          const pattern = new RegExp(`([^(]+)\\(${abbr}\\)`, 'gi');
          const matches = [];
          let match;

          while ((match = pattern.exec(text)) !== null) {
              matches.push(match);
          }

          for (const m of matches) {
              const beforeBrackets = m[1].trim();
              const words = beforeBrackets.split(/[\s-]+/).filter(word =>
                  word.length > 0 &&
                  !['and', 'or', 'the', 'of', 'for', 'in', 'on', 'by', 'to', 'with', 'a', 'an'].includes(word.toLowerCase())
              );

              const firstLetters = words.map(word => word[0].toUpperCase()).join('');
              if (firstLetters === abbr.toUpperCase()) {
                  definitionMap[abbr] = beforeBrackets;
                  break;
              }
          }

          // Try flexible matching if no exact match found
          if (!definitionMap[abbr]) {
              const parenthesesPattern = new RegExp(`([^(]{3,100})\\(${abbr}\\)`, 'gi');
              let parenthesesMatch;

              while ((parenthesesMatch = parenthesesPattern.exec(text)) !== null) {
                  const phraseBeforeBrackets = parenthesesMatch[1].trim();
                  const phraseWords = phraseBeforeBrackets.split(/[\s-]+/).filter(w => w.length > 0);
                  const abbrLetters = abbr.toUpperCase().split('');

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

                      if (abbrPos === abbrLetters.length) {
                          const matchLength = wordPos - start;
                          if (matchLength < bestMatchLength) {
                              bestMatchStart = start;
                              bestMatchLength = matchLength;
                          }
                      }
                  }

                  if (bestMatchStart !== -1) {
                      const relevantWords = phraseWords.slice(bestMatchStart, bestMatchStart + bestMatchLength);
                      definitionMap[abbr] = relevantWords.join(' ');
                      break;
                  }
              }
          }
      });

      // Pattern 2: "Abbreviations: ABBR, definition; ABBR2, definition2"
      const abbreviationSections = text.match(/Abbreviations:([^.]+)/g) || [];
      abbreviationSections.forEach(section => {
          const content = section.replace(/^Abbreviations:/, '').trim();
          const pairs = content.split(';');

          pairs.forEach(pair => {
              const pairMatch = pair.match(/^\s*([A-Z0-9-]+)\s*(?:,|=|:)\s*(.+)$/);
              if (pairMatch && abbreviations.includes(pairMatch[1])) {
                  definitionMap[pairMatch[1]] = pairMatch[2].trim();
              }
          });
      });

      // Pattern 3: "XYZ = full definition"
      const exactDefinitions = text.match(/\b([A-Z][A-Z0-9-]{1,7})\s*(?:=|is|means|:)\s*["']?([^".;:)]+)["']?/g) || [];
      for (const def of exactDefinitions) {
          const match = def.match(/\b([A-Z][A-Z0-9-]{1,7})\s*(?:=|is|means|:)\s*["']?([^".;:)]+)["']?/);
          if (match && match[1] && match[2] && abbreviations.includes(match[1])) {
              definitionMap[match[1]] = match[2].trim();
          }
      }

      // Special cases
      if (abbreviations.includes("IFU")) {
          const ifuMatch = text.match(/information\s+for\s+use\s+\(IFU\)/i);
          if (ifuMatch) {
              definitionMap["IFU"] = "information for use";
          }
      }

      // Normalize capitalization
      Object.keys(definitionMap).forEach(abbr => {
          const definition = definitionMap[abbr];
          if (definition && definition.length > 0) {
              definitionMap[abbr] = definition.charAt(0).toUpperCase() + definition.slice(1);
          }
      });

      return definitionMap;
  }

  /**
   * Main function to find abbreviations in the document
   */
  async function findAbbreviations() {
      ProgressIndicator.start('Finding abbreviations');

      return Word.run(async (context) => {
          try {
              const body = context.document.body;
              body.load("text");
              await context.sync();

              ProgressIndicator.update(20, "Analyzing document text...");

              const text = body.text;
              // Clear known abbreviations before document-wide scan
              knownAbbreviations.clear();
              const abbreviations = detectAbbreviations(text, { 
                  isTable: false, 
                  useKnownAbbreviations: true 
              });

              ProgressIndicator.update(50, "Finding definitions...");

              const definitions = findDefinitions(text, abbreviations);

              console.log("Found abbreviations with definitions:", definitions);
              console.log(`Total unique abbreviations found: ${abbreviations.length}`);

              ProgressIndicator.update(70, "Creating abbreviations table...");

              const searchResults = body.search("Abbreviations", { matchWholeWord: true });
              searchResults.load("items");
              await context.sync();

              if (searchResults.items.length > 0) {
                  const targetParagraph = searchResults.items[0];
                  const tableData = [["Abbreviation", "Definition"]];
                  abbreviations.forEach(abbr => {
                      tableData.push([abbr, definitions[abbr]]);
                  });

                  const table = targetParagraph.insertTable(tableData.length, 2, Word.InsertLocation.after, tableData);
                  table.load("id");
                  await context.sync();

                  const tableRange = table.getRange();
                  tableRange.style = "NICE Table text";
                  await context.sync();
              }

              ProgressIndicator.finish(`Found ${abbreviations.length} abbreviations`);
              return definitions;
          } catch (error) {
              console.error("Error finding abbreviations:", error);
              ProgressIndicator.error("Failed to process abbreviations");
              throw error;
          }
      });
  }

  /**
   * Function to find abbreviations in tables
   */
  async function findTableAbbreviations() {
      ProgressIndicator.start('Processing tables');

      return Word.run(async (context) => {
          try {
              const tables = context.document.body.tables;
              tables.load("items");
              await context.sync();

              const totalTables = tables.items.length;
              let processedTables = 0;
              
              for (let i = 0; i < tables.items.length; i++) {
                  const table = tables.items[i];
                  processedTables++;
                  
                  ProgressIndicator.update(
                      Math.round((processedTables / totalTables) * 80),
                      `Processing table ${processedTables} of ${totalTables}`
                  );

                  if (table.rowCount <= 1) continue;

                  const range = table.getRange();
                  const afterRange = range.insertParagraph("", Word.InsertLocation.after);
                  const nextParagraph = afterRange.getNext();
                  nextParagraph.load("text");
                  await context.sync();
                  
                  afterRange.delete();
                  await context.sync();

                  const nextParagraphText = nextParagraph.text.trim();
                  if (nextParagraphText.startsWith("Abbreviations") && 
                      nextParagraphText !== "Abbreviations" && 
                      (nextParagraphText.includes("-") || nextParagraphText.includes(";"))) {
                      continue;
                  }
                  
                  range.load("text");
                  await context.sync();
                  
                  const tableText = range.text;
                  const hasHeaders = table.rowCount > 1 && tableText.split('\n')[0].toUpperCase() === tableText.split('\n')[0];
                  const abbreviations = detectAbbreviations(tableText, { 
                      isTable: true, 
                      hasHeaders,
                      useKnownAbbreviations: true 
                  });
                  
                  if (abbreviations.length > 0) {
                      const paragraph = table.insertParagraph("", Word.InsertLocation.after);
                      const definitions = findDefinitions(tableText, abbreviations);
                      
                      const formattedAbbreviations = "Abbreviations: " + 
                          abbreviations.map(abbr => `${abbr} - ${definitions[abbr] || ""}`).join("; ");
                      
                      paragraph.insertText(formattedAbbreviations, Word.InsertLocation.replace);
                      paragraph.style = "NICE Footnote";
                      await context.sync();
                  }
              }
              
              ProgressIndicator.finish(`Processed ${totalTables} tables`);
          } catch (error) {
              console.error("Error scanning tables:", error);
              ProgressIndicator.error("Failed to process tables");
          }
      });
  }
});