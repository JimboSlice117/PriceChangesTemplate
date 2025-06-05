// @ts-nocheck
/**
 * Cross-Platform Price Management System
 * Complete Apps Script implementation
 * Accuracy-focused version with enhanced SKU matching and performance optimizations.
 * Batch processing has been removed. SKU Match Review sidebar has been removed.
 * Includes a dedicated "Price Changes" tab.
 *
 * Changes:
 * - Fixed 'reviewRequiredMatchesOverall' typo in runSkuMatching.
 * - Modified logic to effectively ignore "Clean Sku" columns for matching by not populating/using cleanSkuMaps.
 * - Added a local cache for extractSkuAttributesAndCore results within accurateFindBestMatch_optimized.
 * - Enhanced error logging and added more try-catch blocks for robustness.
 * - Added isNaN checks for numeric conversions.
 */

// ---------------- GLOBAL CONSTANTS ----------------
const SS = SpreadsheetApp.getActiveSpreadsheet();
const PLATFORM_TABS = {
  EBAY: { name: "eBay Export", headers: ["Action", "Item number", "Custom label (SKU)", "Start price"] },
  AMAZON: { name: "Amazon Export", headers: ["seller-sku", "price"] },
  SHOPIFY: { name: "Shopify Export", headers: ["Variant SKU", "Variant Price", "Variant Compare At Price", "Variant Cost"] },
  INFLOW: { name: "inFlow Export", headers: ["Name", "UnitPrice", "Cost"] },
  SELLERCLOUD: { name: "SellerCloud Export", headers: ["ProductID", "MAPPrice", "SitePrice", "SiteCost"] },
  REVERB: { name: "Reverb Export", headers: ["sku", "price", "condition"] }
};
const ANALYSIS_PLATFORMS = ['AMAZON', 'EBAY', 'SHOPIFY', 'REVERB']; // Platforms to include in analysis
const BSTOCK_CATEGORIES = {
  'BA': 0.95, 'BB': 0.90, 'BC': 0.85, 'BD': 0.80,
  'NOACC': 'special', 'AA': 'special'
};
const COLOR_ABBREVIATIONS = {
  'BK': 'BLACK', 'BL': 'BLUE', 'BR': 'BROWN', 'GY': 'GRAY', 'GR': 'GREEN', 'IV': 'IVORY',
  'OR': 'ORANGE', 'PK': 'PINK', 'RD': 'RED', 'TN': 'TAN', 'WH': 'WHITE', 'YL': 'YELLOW',
  'PL': 'PURPLE', 'NV': 'NAVY', 'CH': 'CHERRY', 'NT': 'NATURAL', 'MH': 'MAHOGANY', 'WN': 'WALNUT',
  'BLACK':'BLACK', 'BLUE':'BLUE', 'BROWN':'BROWN', 'GRAY':'GRAY', 'GREEN':'GREEN', 'IVORY':'IVORY',
  'ORANGE':'ORANGE', 'PINK':'PINK', 'RED':'RED', 'TAN':'TAN', 'WHITE':'WHITE', 'YELLOW':'YELLOW',
  'PURPLE':'PURPLE', 'NAVY':'NAVY', 'CHERRY':'CHERRY', 'NATURAL':'NATURAL', 'MAHOGANY':'MAHOGANY', 'WALNUT':'WALNUT'
};
const COLORS = {
  HEADER_BG: '#4285F4', HEADER_TEXT: '#FFFFFF', AMAZON: '#FF9900', EBAY: '#E53238',
  SHOPIFY: '#96BF48', INFLOW: '#00A1E0', SELLERCLOUD: '#6441A5', REVERB: '#F5756C',
  POSITIVE: '#0F9D58', NEGATIVE: '#DB4437', WARNING: '#F4B400',
  PRICE_UP: '#f4cccc', PRICE_DOWN: '#d9ead3'
};
const skuNormalizeCache = {}; // Global cache for conservativeNormalizeSku
let errorLogSheet = null; // Global variable for error log sheet

// ---------------- UTILITY FUNCTIONS ----------------

function getBStockInfo(sku) {
  if (!sku || (typeof sku !== 'string' && typeof sku !== 'number')) return null;
  const upperSku = String(sku).toUpperCase();
  for (const [bStockType, multiplierOrSpecial] of Object.entries(BSTOCK_CATEGORIES)) {
    if (upperSku.includes('-' + bStockType) || upperSku.startsWith(bStockType + '-') || upperSku.endsWith('-' + bStockType) ) {
      let actualMultiplier; let isSpecial = false;
      if (typeof multiplierOrSpecial === 'number') {
        actualMultiplier = multiplierOrSpecial;
      } else {
        isSpecial = true;
        if (bStockType === 'NOACC') actualMultiplier = BSTOCK_CATEGORIES['BC']; // Default NOACC to BC multiplier
        else if (bStockType === 'AA') actualMultiplier = 0.98; // AA specific multiplier
        else actualMultiplier = 0.85; // Default for other 'special' types
      }
      return { type: bStockType, multiplier: actualMultiplier, sku: sku, isSpecial: isSpecial };
    }
  }
  return null;
}

function getColumnIndices(headers, columnNames) {
  const indices = {};
  columnNames.forEach(name => {
    indices[name] = headers.indexOf(name);
  });
  return indices;
}

function getPlatformVariations(platform) {
  const variations = [platform, platform.charAt(0) + platform.slice(1).toLowerCase(), platform.toLowerCase()];
  if (platform === 'SELLERCLOUD') variations.push('ellerCloud');
  if (platform === 'INFLOW') variations.push('inFlow');
  return variations.filter(Boolean);
}

function initializeErrorLogSheet() {
    if (!errorLogSheet) {
        errorLogSheet = SS.getSheetByName('Error Log');
        if (!errorLogSheet) {
            errorLogSheet = SS.insertSheet('Error Log');
            errorLogSheet.appendRow(['Timestamp', 'Function', 'Error Message', 'SKU/Item', 'Details']);
            SpreadsheetApp.flush(); // Ensure sheet is created
        }
    }
    return errorLogSheet;
}

function logError(functionName, errorMessage, skuOrItem = '', details = '', showAlert = false, uiInstance = null) {
    try {
        initializeErrorLogSheet();
        const timestamp = new Date();
        Logger.log(`ERROR in ${functionName} at ${timestamp}: ${errorMessage}. Item: ${skuOrItem}. Details: ${details ? (details.stack || details) : ''}`);
        if (errorLogSheet) {
            errorLogSheet.appendRow([timestamp, functionName, errorMessage, skuOrItem, details ? (details.stack || details.toString()) : '']);
        }
        if (showAlert && uiInstance) {
            uiInstance.alert('Error', `An error occurred in ${functionName}: ${errorMessage}`, uiInstance.ButtonSet.OK);
        }
    } catch (e) {
        Logger.log(`Failed to write to Error Log: ${e.toString()}`);
    }
}

function longestCommonSubstring(str1, str2) {
  if (!str1 || !str2) return '';
  const s1 = [...str1], s2 = [...str2];
  const matrix = Array(s1.length + 1).fill(null).map(() => Array(s2.length + 1).fill(0));
  let maxLength = 0, endPosition = 0;

  for (let i = 1; i <= s1.length; i++) {
    for (let j = 1; j <= s2.length; j++) {
      if (s1[i - 1] === s2[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1] + 1;
        if (matrix[i][j] > maxLength) {
          maxLength = matrix[i][j];
          endPosition = i;
        }
      }
    }
  }
  return str1.substring(endPosition - maxLength, endPosition);
}

// ---------------- ACCURACY-FOCUSED SKU MATCHING FUNCTIONS ----------------

function conservativeNormalizeSku(sku) {
  if (!sku || (typeof sku !== 'string' && typeof sku !== 'number')) return '';
  const skuStr = String(sku);
  if (skuNormalizeCache[skuStr]) return skuNormalizeCache[skuStr];

  let normalized = skuStr.toUpperCase()
    .replace(/[^A-Z0-9\-]/g, '') // Keep only alphanumeric and hyphens
    .replace(/\-+/g, '-')       // Replace multiple hyphens with a single one
    .replace(/^-|-$/g, '');     // Remove leading/trailing hyphens

  skuNormalizeCache[skuStr] = normalized;
  return normalized;
}

function conservativeExtractCore(sku) {
  if (!sku) return '';
  let core = sku;
  // More specific prefixes based on common patterns. Order might matter if prefixes can overlap.
  const obviousPrefixes = [
    /^EMG-/, /^SHU-/, /^BOSS-/, /^MCK-/, /^GGACC-/, /^1SV-/,
    /^AA-/, // Shopify specific often, but could be generic
    /^360H-/, /^8SI-/ // Amazon specific prefixes
  ];
  for (const prefix of obviousPrefixes) {
    if (prefix.test(core)) {
      core = core.replace(prefix, '');
      break; // Assume only one such prefix
    }
  }
  // Suffixes
  const obviousSuffixes = [
    /-FOL$/, /-FOLIOS?$/ // Amazon FOL suffix
  ];
  for (const suffix of obviousSuffixes) {
    if (suffix.test(core)) {
      core = core.replace(suffix, '');
      break; // Assume only one such suffix
    }
  }
  return core.replace(/^-|-$/g, ''); // Clean up any leading/trailing hyphens from replacements
}

function extractSkuAttributesAndCore(rawSku) {
    if (!rawSku || (typeof rawSku !== 'string' && typeof rawSku !== 'number')) {
        return { originalSku: rawSku, normalizedSku: '', coreSku: '', bStock: null, color: null };
    }

    const originalSkuUpper = String(rawSku).toUpperCase();
    const normalizedForProcessing = conservativeNormalizeSku(originalSkuUpper); // Basic normalization
    const bStockInfo = getBStockInfo(originalSkuUpper); // Check for B-Stock based on original SKU structure

    let tempSkuForCore = normalizedForProcessing;

    // Remove B-Stock identifiers from the SKU string that will be used to find the color and core
    // This helps isolate the core product SKU from B-Stock indicators
    if (bStockInfo) {
        const bStockType = bStockInfo.type;
        // Define patterns to remove the B-Stock type. Handle cases where it might be at start, end, or middle.
        const patternsToRemove = [
            '-' + bStockType + '-', // Middle: ABC-BA-XYZ -> ABC-XYZ
            '-' + bStockType,       // End:   ABC-BA -> ABC
            bStockType + '-',       // Start: BA-ABC -> ABC
        ];
        for (const pattern of patternsToRemove) {
            if (tempSkuForCore.includes(pattern)) {
                // More careful replacement to avoid issues if bStockType itself contains regex characters (unlikely with current BSTOCK_CATEGORIES)
                // Replace with a single hyphen if it was in the middle, or empty if at start/end
                tempSkuForCore = tempSkuForCore.replace(new RegExp(pattern.replace(/-/g, '\\-'), 'g'), pattern === `-${bStockType}-` ? '-' : '');
            }
        }
        tempSkuForCore = tempSkuForCore.replace(/^-|-$/g, '').replace(/\-\-/g,'-'); // Clean up resulting hyphens
    }

    let foundColor = null;
    let skuAfterColorRemoval = tempSkuForCore;

    // Sort color abbreviations by length (descending) to match longer ones first (e.g., "CHERRY" before "CH")
    const sortedColorAbbrs = Object.keys(COLOR_ABBREVIATIONS).sort((a, b) => b.length - a.length);

    for (const abbr of sortedColorAbbrs) {
        const upperAbbr = abbr.toUpperCase();
        let replaced = false;

        // Try matching as a suffix with a preceding hyphen (e.g., SKU-COLOR)
        if (skuAfterColorRemoval.endsWith('-' + upperAbbr)) {
            skuAfterColorRemoval = skuAfterColorRemoval.substring(0, skuAfterColorRemoval.length - (upperAbbr.length + 1));
            foundColor = COLOR_ABBREVIATIONS[abbr];
            replaced = true;
        }
        // Try matching as a direct suffix, but be more careful
        // Only if it's a longer abbreviation or not preceded by another letter/number (to avoid partial matches like "BINDER" -> "RED")
        if (!replaced && skuAfterColorRemoval.endsWith(upperAbbr)) {
            const precedingCharIndex = skuAfterColorRemoval.length - upperAbbr.length - 1;
            // Condition: Is it the start of the SKU, or is the preceding char not alphanumeric, or is it a long color name?
            if (precedingCharIndex < 0 || !/[A-Z0-9]/.test(skuAfterColorRemoval.charAt(precedingCharIndex)) || upperAbbr.length > 2) { // Min length for non-hyphenated color
                skuAfterColorRemoval = skuAfterColorRemoval.substring(0, skuAfterColorRemoval.length - upperAbbr.length);
                foundColor = COLOR_ABBREVIATIONS[abbr];
                replaced = true;
            }
        }
        if (replaced) break; // Found a color, stop searching
    }
    skuAfterColorRemoval = skuAfterColorRemoval.replace(/-$/, ''); // Remove trailing hyphen if any

    // Extract core after B-Stock and color removal
    let finalCore = conservativeExtractCore(skuAfterColorRemoval);
    finalCore = finalCore.replace(/^-|-$/g, ''); // Final cleanup of core

    return {
        originalSku: rawSku,
        normalizedSku: normalizedForProcessing, // The result of conservativeNormalizeSku on original
        coreSku: finalCore,
        bStock: bStockInfo,
        color: foundColor
    };
}

function levenshteinDistance(a, b) {
  if (!a && !b) return 0;
  if (!a) return b.length;
  if (!b) return a.length;

  const matrix = Array(b.length + 1).fill(null).map(() => Array(a.length + 1).fill(0));
  for (let i = 0; i <= a.length; i++) matrix[0][i] = i;
  for (let j = 0; j <= b.length; j++) matrix[j][0] = j;

  for (let j = 1; j <= b.length; j++) {
    for (let i = 1; i <= a.length; i++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[j][i] = Math.min(
        matrix[j][i - 1] + 1,        // Deletion
        matrix[j - 1][i] + 1,        // Insertion
        matrix[j - 1][i - 1] + cost  // Substitution
      );
    }
  }
  return matrix[b.length][a.length];
}

function strictSimilarityCheck(coreSku1, coreSku2) {
    if (!coreSku1 && !coreSku2) return { score: 5, reason: 'Both cores empty (attribute-only SKUs)'}; // Both are empty, could be attribute-only SKUs
    if (!coreSku1 || !coreSku2) return { score: 0, reason: 'Empty Core SKU' }; // One is empty, no match

    const len1 = coreSku1.length;
    const len2 = coreSku2.length;

    // Handle very short core SKUs (e.g. "A", "BC")
    if (len1 < 2 || len2 < 2) { // Adjusted minimum length, e.g. single characters are too ambiguous
        if (coreSku1 === coreSku2) return { score: 90, reason: "Exact short core match" };
        const shortDist = levenshteinDistance(coreSku1, coreSku2);
        if (shortDist <=1 && Math.max(len1, len2) <=2) return {score: 70, reason: `Near exact short core (dist ${shortDist})`}; // e.g. "A" vs "B" or "AB" vs "AC"
        return { score: 0, reason: `Core too short for non-exact (c1:${len1}, c2:${len2})` };
    }

    // Length ratio check: if one core is much shorter than the other, it's unlikely a good match
    const lengthRatio = Math.min(len1, len2) / Math.max(len1, len2);
    if (lengthRatio < 0.45) return { score: 0, reason: `Core length ratio too low (${lengthRatio.toFixed(2)})` }; // Stricter ratio

    if (coreSku1 === coreSku2) return { score: 95, reason: `Exact core match (${coreSku1})` };

    const coreDistance = levenshteinDistance(coreSku1, coreSku2);
    const coreMaxLength = Math.max(len1, len2);
    const coreSimilarity = ((coreMaxLength - coreDistance) / coreMaxLength);

    if (coreSimilarity >= 0.88) return { score: Math.min(90, 65 + Math.round(coreSimilarity * 30)), reason: `High core similarity: ${Math.round(coreSimilarity * 100)}% (dist ${coreDistance})` }; // Adjusted scoring
    // Check for containment (e.g., "PART" in "PARTXYZ" or "XYZPART")
    if (len1 >= 3 && coreSku2.includes(coreSku1)) return { score: 88, reason: `Core1 (${coreSku1}) in Core2 (${coreSku2})` };
    if (len2 >= 3 && coreSku1.includes(coreSku2)) return { score: 88, reason: `Core2 (${coreSku2}) in Core1 (${coreSku1})` };

    // Longest common substring check
    const commonSub = longestCommonSubstring(coreSku1, coreSku2);
    if (commonSub.length >= Math.max(2, Math.min(len1, len2) * 0.50)) { // At least half of the shorter SKU
        const overlapRatioCore = commonSub.length / Math.min(len1, len2);
        if (overlapRatioCore >= 0.60) return { score: Math.min(85, 55 + Math.round(overlapRatioCore * 35)), reason: `Strong core overlap: ${commonSub} (${Math.round(overlapRatioCore * 100)}%)`};
    }
    
    if (coreSimilarity >= 0.75) return { score: Math.min(80, 50 + Math.round(coreSimilarity * 35)), reason: `Good core similarity: ${Math.round(coreSimilarity * 100)}% (dist ${coreDistance})` };

    return { score: Math.max(0, Math.round(coreSimilarity * 60)), reason: `Low core similarity: ${Math.round(coreSimilarity*100)}% (dist ${coreDistance})` }; // Scaled down low similarity
}


function conservativePlatformMatch(mfrSkuNormalized, platformSkuNormalized, platform) {
    if (!mfrSkuNormalized || !platformSkuNormalized) return { score: 0, reason: 'Empty SKU for platform match' };

    let score = 0;
    let reason = '';
    const mfrCorePlat = conservativeExtractCore(mfrSkuNormalized); // Use the basic core extractor
    const platCorePlat = conservativeExtractCore(platformSkuNormalized);

    // Platform-specific prefix/suffix rules (these are high-confidence overrides)
    switch (platform) {
        case 'AMAZON':
            if (platformSkuNormalized.endsWith('-FOL') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Amazon FOL suffix, base match'; }
            else if ((platformSkuNormalized.startsWith('360H-') || platformSkuNormalized.startsWith('8SI-')) && platformSkuNormalized.substring(platformSkuNormalized.indexOf('-')+1) === mfrSkuNormalized) { score = 96; reason = 'Amazon prefix, exact mfr SKU'; }
            break;
        case 'SELLERCLOUD': // Often uses EMG- prefix
            if (platformSkuNormalized.startsWith('EMG-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'SC EMG prefix, base match'; }
            else if (platformSkuNormalized.startsWith('EMG-') && platformSkuNormalized.substring(4) === mfrSkuNormalized) { score = 96; reason = 'SC EMG prefix, exact mfr SKU';}
            break;
        case 'REVERB': // Prefixes like SHU-, BOSS-, MCK-
            if (platformSkuNormalized.startsWith('SHU-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Reverb SHU prefix, base match'; }
            else if (platformSkuNormalized.startsWith('BOSS-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Reverb BOSS prefix, base match'; }
            else if (platformSkuNormalized.startsWith('MCK-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Reverb MCK prefix, base match'; }
            break;
        case 'EBAY': // GGACC- prefix
             if (platformSkuNormalized.startsWith('GGACC-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'eBay GGACC prefix, base match'; }
            break;
        case 'SHOPIFY': // AA- prefix for non-BStock
            if (platformSkuNormalized.startsWith('AA-') && mfrCorePlat === platCorePlat && mfrCorePlat.length >= 4 && !getBStockInfo(platformSkuNormalized)) {
                score = 92; reason = 'Shopify AA prefix (non-BStock), base match (len>=4)';
            }
            break;
        case 'INFLOW': // 1SV- prefix
            if (platformSkuNormalized.startsWith('1SV-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'inFlow 1SV prefix, base match'; }
            break;
    }

    // B-Stock matching consistency check (boosts score if B-stock types match, or if one is new and other is B-stock of same core)
    const mfrBstockPlat = getBStockInfo(mfrSkuNormalized); // Check B-Stock on normalized (already uppercased)
    const platBstockPlat = getBStockInfo(platformSkuNormalized);

    if (mfrBstockPlat && platBstockPlat && mfrBstockPlat.type === platBstockPlat.type) {
        // Both are B-Stock of the same type, check if underlying SKUs match
        let baseMfr = mfrSkuNormalized.replace('-'+mfrBstockPlat.type, '').replace(mfrBstockPlat.type+'-','').replace(/^-|-$/g, '');
        let basePlat = platformSkuNormalized.replace('-'+platBstockPlat.type, '').replace(platBstockPlat.type+'-','').replace(/^-|-$/g, '');
        if (baseMfr === basePlat && baseMfr.length > 0) { // Ensure base is not empty
            score = Math.max(score, 97); // Very high confidence
            reason = (reason ? reason + "; " : "") + `Platform B-Stock (${mfrBstockPlat.type}) base match`;
        }
    } else if (!mfrBstockPlat && platBstockPlat) { // MFR is new, Platform is B-Stock
        let basePlat = platformSkuNormalized.replace('-'+platBstockPlat.type, '').replace(platBstockPlat.type+'-','').replace(/^-|-$/g, '');
        if (mfrSkuNormalized === basePlat && mfrSkuNormalized.length > 0) {
            score = Math.max(score, 96); // High confidence
            reason = (reason ? reason + "; " : "") + `Platform B-Stock (${platBstockPlat.type}) matches new MFR SKU base`;
        }
    }
    // Could add a case for mfrBstockPlat && !platBstockPlat if needed, but often MFR sheet is "new"

    return { score: score, reason: reason };
}


function validateMatch_optimized(mfrSkuRaw, mfrAtt, platformSkuRaw, platAtt, platform) {
    if (!mfrSkuRaw || !platformSkuRaw) return { valid: false, confidence: 0, reason: 'Empty SKU provided to validateMatch' };

    mfrAtt = mfrAtt || extractSkuAttributesAndCore(mfrSkuRaw); // Ensure attributes are populated
    platAtt = platAtt || extractSkuAttributesAndCore(platformSkuRaw);

    // Highest confidence: Exact raw match (case-insensitive)
    if (String(mfrSkuRaw).toUpperCase() === String(platformSkuRaw).toUpperCase()) {
        return { valid: true, confidence: 100, reason: 'Exact raw match (case-insensitive)' };
    }

    // High confidence: Exact normalized match (after basic conservative normalization)
    if (mfrAtt.normalizedSku && platAtt.normalizedSku && mfrAtt.normalizedSku === platAtt.normalizedSku) {
        return { valid: true, confidence: 99, reason: `Exact normalized match (${mfrAtt.normalizedSku})` };
    }
    
    // Handle very short SKUs after initial checks (normalized can be short)
    if ((mfrAtt.normalizedSku && mfrAtt.normalizedSku.length < 3) || (platAtt.normalizedSku && platAtt.normalizedSku.length < 3)) {
        const dist = levenshteinDistance(mfrAtt.normalizedSku, platAtt.normalizedSku);
        const maxL = Math.max(mfrAtt.normalizedSku.length, platAtt.normalizedSku.length);
        if (dist === 0 && maxL > 0) return { valid: true, confidence: 98, reason: `Short exact normalized (${mfrAtt.normalizedSku})`}; // e.g. "AB" == "AB"
        if (dist <= 1 && maxL <= 4 && maxL > 0) { // Max length of 4 for this rule, distance 1
             return { valid: true, confidence: 70 + (5-maxL)*2, reason: `Short SKU near match (dist ${dist}, len ${maxL})` }; // e.g. "ABC" vs "ABD"
        }
        const sim = maxL > 0 ? (maxL - dist) / maxL : 0;
        if (sim >= 0.60 && maxL > 0) return { valid: 'NEEDS_REVIEW', confidence: Math.max(50, Math.round(sim * 60)), reason: `Short SKU, mod. similarity ${Math.round(sim * 100)}%` };
        return { valid: false, confidence: 0, reason: `SKU too short (MfrN:${mfrAtt.normalizedSku}, PlatN:${platAtt.normalizedSku})` };
    }


    // --- Core Similarity and Attribute Bonus/Penalty ---
    const platformMatchResult = conservativePlatformMatch(mfrAtt.normalizedSku, platAtt.normalizedSku, platform);
    const coreSimilarityResult = strictSimilarityCheck(mfrAtt.coreSku, platAtt.coreSku);
    let attributeBonus = 0;
    let attributeReason = "";
    let coreIdenticalAndAttributesDiffer = false; // Flag for specific scenario

    // Bonus if core SKUs are identical (and not empty)
    if (mfrAtt.coreSku && platAtt.coreSku && mfrAtt.coreSku === platAtt.coreSku && mfrAtt.coreSku !== "") {
        attributeBonus += 5; // Small bonus for identical non-empty cores
    }

    // B-Stock comparison
    if (mfrAtt.bStock && platAtt.bStock) { // Both have B-Stock info
        if (mfrAtt.bStock.type === platAtt.bStock.type) {
            attributeBonus += 15; attributeReason += `B-Stock type match (${mfrAtt.bStock.type}). `;
        } else {
            attributeBonus -= 10; attributeReason += `B-Stock type mismatch (${mfrAtt.bStock.type} vs ${platAtt.bStock.type}). `;
            if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true;
        }
    } else if (mfrAtt.bStock || platAtt.bStock) { // Only one has B-Stock info
        const bStockSource = mfrAtt.bStock ? `MFR (${mfrAtt.bStock.type})` : `Plat (${platAtt.bStock.type})`;
        if (mfrAtt.coreSku === platAtt.coreSku && mfrAtt.coreSku !== "") { // If cores match, this difference might be expected (e.g. new vs b-stock of same item)
             attributeBonus += 8; attributeReason += `Expected B-Stock diff (${bStockSource}). `;
        } else {
             attributeBonus -= 5; attributeReason += `B-Stock presence diff (${bStockSource}). `;
        }
        if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true;
    }

    // Color comparison
    if (mfrAtt.color && platAtt.color) { // Both have color info
        if (mfrAtt.color === platAtt.color) {
            attributeBonus += 15; attributeReason += `Color match (${mfrAtt.color}). `;
        } else {
            attributeBonus -= 10; attributeReason += `Color mismatch (${mfrAtt.color} vs ${platAtt.color}). `;
            if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true;
        }
    } else if (mfrAtt.color || platAtt.color) { // Only one has color info
        const colorSource = mfrAtt.color ? `MFR (${mfrAtt.color})` : `Plat (${platAtt.color})`;
         if (mfrAtt.coreSku === platAtt.coreSku && mfrAtt.coreSku !== "") { // If cores match, this difference might be expected
            attributeBonus += 8; attributeReason += `Expected Color diff (${colorSource}). `;
        } else {
            attributeBonus -= 5; attributeReason += `Color presence diff (${colorSource}). `;
        }
        if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true;
    }

    attributeBonus = Math.max(-20, Math.min(attributeBonus, 30)); // Cap bonus/penalty

    let finalScore = 0;
    let finalReason = "";

    if (coreSimilarityResult.score > 0) {
        finalReason = `CoreSim(${coreSimilarityResult.score}%): ${coreSimilarityResult.reason}. `;
        if (attributeReason) finalReason += `Attr: ${attributeReason}(Bonus ${attributeBonus}). `;
        finalScore = coreSimilarityResult.score + attributeBonus;
    } else {
        // Fallback to full normalized similarity if core sim is 0 (e.g. cores are too different or one is empty)
        const fullNormalizedDistance = levenshteinDistance(mfrAtt.normalizedSku, platAtt.normalizedSku);
        const fullNormalizedMaxLength = Math.max(mfrAtt.normalizedSku.length, platAtt.normalizedSku.length);
        const fullNormalizedSimilarity = fullNormalizedMaxLength > 0 ? ((fullNormalizedMaxLength - fullNormalizedDistance) / fullNormalizedMaxLength) : 0;
        let baseScore = Math.round(fullNormalizedSimilarity * 60); // Max 60 from this path
        finalReason = `Low CoreSim. FullNormSim(${Math.round(fullNormalizedSimilarity*100)}%). `;
        if (attributeReason) finalReason += `Attr: ${attributeReason}(Bonus ${attributeBonus}). `;
        finalScore = baseScore + attributeBonus;
    }
    
    if (coreIdenticalAndAttributesDiffer && finalScore > 70) { // Add note if cores were identical but attributes caused issues
        finalReason += "Cores matched but attributes differed. ";
    }

    // Consider platform-specific match score
    if (platformMatchResult.score > 0) {
        // If platform specific score is significantly higher or very high itself, it can override
        if (platformMatchResult.score > finalScore + 10 || (platformMatchResult.score >= 90 && finalScore < 90) ) {
            finalScore = platformMatchResult.score;
            finalReason = `Platform Specific: ${platformMatchResult.reason} (Core/Attr: ${finalScore > 0 ? finalScore : 'N/A'})`;
        } else if (finalScore < 50 && platformMatchResult.score > finalScore) { // If core/attr score is very low, platform hint can boost
            finalScore = platformMatchResult.score;
            finalReason = `Platform Hint: ${platformMatchResult.reason} (Low core/attr).`;
        } else { // Otherwise, just note it and take the max
            finalReason += ` PlatformNote: ${platformMatchResult.reason} (Score ${platformMatchResult.score}).`;
            finalScore = Math.max(finalScore, platformMatchResult.score); // Take the higher of the two if not overriding
        }
    }
    
    finalScore = Math.min(98, Math.max(0, Math.round(finalScore))); // Cap final score (100 is for exact raw, 99 for exact norm)

    if (finalScore >= 85) return { valid: true, confidence: finalScore, reason: finalReason };
    if (finalScore >= 70) return { valid: 'NEEDS_REVIEW', confidence: finalScore, reason: `Review: ${finalReason}` };
    return { valid: false, confidence: finalScore, reason: `Low Confidence: ${finalReason}` };
}


function accurateFindBestMatch_optimized(rawMfrSku, mfrExtractedAttributes, platformItemsWithAttributes, platformSkuMap, platform) {
    // Local cache for extractSkuAttributesAndCore results for the current MFR SKU processing
    const localAttributeExtractionCache = new Map();
    function getCachedPlatformAttributes(platSku) {
        if (localAttributeExtractionCache.has(platSku)) {
            return localAttributeExtractionCache.get(platSku);
        }
        // For platform SKUs, we generally don't assume isMfgSku = true unless it's a specific scenario
        const attributes = extractSkuAttributesAndCore(platSku, false);
        localAttributeExtractionCache.set(platSku, attributes);
        return attributes;
    }


    const normalizedMfrSkuForMapLookup = mfrExtractedAttributes.normalizedSku;

    // 1. Check direct match on normalized SKU using platformSkuMap
    if (platformSkuMap && platformSkuMap[normalizedMfrSkuForMapLookup]) {
        const platItemContainer = platformSkuMap[normalizedMfrSkuForMapLookup];
        // Since platItemContainer.extractedAttributes should already be populated from preProcessAllPlatformData, use that directly
        const validation = validateMatch_optimized(rawMfrSku, mfrExtractedAttributes, platItemContainer.sku, platItemContainer.extractedAttributes, platform);
        if (validation.confidence >= 98) { // High threshold for this direct map lookup
            return {
                platformSku: platItemContainer.sku,
                currentPrice: platItemContainer.price,
                currentCost: platItemContainer.cost,
                confidenceScore: validation.confidence,
                matchType: 'Exact Normalized (Validated)',
                matchReason: validation.reason
            };
        }
    }

    // 2. Iterate through all platform items for more detailed matching (if no quick map match or low confidence)
    let bestMatch = null;
    let highestConfidence = 0;

    for (const itemContainer of platformItemsWithAttributes) {
        if (!itemContainer.sku) continue;
        
        // Use pre-calculated attributes if available, otherwise calculate (and cache locally if this loop is for the same mfrSku)
        // itemContainer.extractedAttributes should be populated by preProcessAllPlatformData
        const platformAttributes = itemContainer.extractedAttributes || getCachedPlatformAttributes(itemContainer.sku);
        if (!platformAttributes) continue;


        const validation = validateMatch_optimized(rawMfrSku, mfrExtractedAttributes, itemContainer.sku, platformAttributes, platform);

        if (validation.confidence > highestConfidence) {
            highestConfidence = validation.confidence;
            bestMatch = {
                platformSku: itemContainer.sku,
                currentPrice: itemContainer.price,
                currentCost: itemContainer.cost,
                confidenceScore: validation.confidence,
                matchType: validation.valid === true ? 'Validated-Strong' : (validation.valid === 'NEEDS_REVIEW' ? 'Validated-Needs-Review' : 'Validated-Low-Confidence'),
                matchReason: validation.reason
            };
        }
    }

    if (bestMatch) {
        if (bestMatch.confidenceScore >= 85) return bestMatch;
        // For scores between 70 and 84, flag for review.
        if (bestMatch.confidenceScore >= 70) {
            bestMatch.matchType = 'LOW-CONFIDENCE-REVIEW-REQUIRED'; // Overwrite match type
            bestMatch.matchReason = `REVIEW (Score ${bestMatch.confidenceScore}): ${bestMatch.matchReason}`;
            return bestMatch;
        }
        // Implicitly, if below 70, it's still null or a very low confidence match not returned.
    }

    return null; // No suitable match found
}


// ---------------- PRE-PROCESSING OF PLATFORM DATA ----------------
function preProcessAllPlatformData(platformDataRaw) {
  const platformSkuMaps = {};
  // const cleanSkuMaps = {}; // No longer populating or using cleanSkuMaps
  const preProcessedPlatformDataWithAttributes = {};

  Logger.log("Starting pre-processing of all platform data...");
  for (const platform in platformDataRaw) {
    platformSkuMaps[platform] = {};
    // cleanSkuMaps[platform] = {}; // Do not initialize if not used
    preProcessedPlatformDataWithAttributes[platform] = platformDataRaw[platform].map(item => {
      if (!item.sku) return { ...item, extractedAttributes: null /*, cleanExtractedAttributes: null*/ };

      const attributes = extractSkuAttributesAndCore(item.sku);
      // const itemWithAttributes = { ...item, extractedAttributes: attributes, cleanExtractedAttributes: null };
      const itemWithAttributes = { ...item, extractedAttributes: attributes };


      if (attributes && attributes.normalizedSku) {
        platformSkuMaps[platform][attributes.normalizedSku] = itemWithAttributes;
      }

      // Logic for Clean Sku removed as per user request to ignore it
      // if (item.cleanSku) {
      //   const cleanAttributes = extractSkuAttributesAndCore(item.cleanSku);
      //   itemWithAttributes.cleanExtractedAttributes = cleanAttributes;
      //   if (cleanAttributes && cleanAttributes.normalizedSku) {
      //     cleanSkuMaps[platform][cleanAttributes.normalizedSku] = itemWithAttributes;
      //   }
      // }
      return itemWithAttributes;
    });
    Logger.log(`Platform ${platform} (for pre-processing): Processed ${preProcessedPlatformDataWithAttributes[platform].length} items.`);
  }
  Logger.log("Full platform data pre-processing complete.");
  // return { preProcessedPlatformDataWithAttributes, platformSkuMaps, cleanSkuMaps };
  return { preProcessedPlatformDataWithAttributes, platformSkuMaps };
}


// ---------------- CORE MATCHING LOGIC (SINGLE RUN) ----------------
function performFullMatching(manufacturerData, preProcessedPlatformData, platformSkuMaps /*, cleanSkuMaps - removed */) {
  const matchResults = [];
  Object.keys(skuNormalizeCache).forEach(key => delete skuNormalizeCache[key]); // Clear global normalization cache at start of full matching

  const localAttributeExtractionCache = new Map(); // Cache for extractSkuAttributesAndCore for the current MFR item

  manufacturerData.forEach((mfrItem, index) => {
    localAttributeExtractionCache.clear(); // Clear for each new MFR item

    if (index > 0 && index % 100 === 0) {
        Logger.log(`Processing MFR SKU ${index + 1} of ${manufacturerData.length}: ${mfrItem.manufacturerSku}`);
        SpreadsheetApp.flush(); 
    }
    const rawMfrSku = mfrItem.manufacturerSku;
    
    // Cache MFR SKU's attribute extraction as it's used for every platform
    let mfrExtractedAttributes;
    if (localAttributeExtractionCache.has(rawMfrSku)) {
        mfrExtractedAttributes = localAttributeExtractionCache.get(rawMfrSku);
    } else {
        mfrExtractedAttributes = extractSkuAttributesAndCore(rawMfrSku);
        localAttributeExtractionCache.set(rawMfrSku, mfrExtractedAttributes);
    }

    const result = {
      manufacturerSku: rawMfrSku,
      msrp: mfrItem.msrp,
      map: mfrItem.map,
      dealerPrice: mfrItem.dealerPrice,
      matches: {}
    };

    for (const platform in preProcessedPlatformData) {
      try {
        result.matches[platform] = accurateFindBestMatch_optimized(
          rawMfrSku,
          mfrExtractedAttributes, // Pass the cached attributes
          preProcessedPlatformData[platform],
          platformSkuMaps[platform],
          // cleanSkuMaps[platform], // Removed
          platform
        );
      } catch (error) {
        logError('performFullMatching (platform loop)', `Error matching ${rawMfrSku} on ${platform}`, rawMfrSku, error, false);
        result.matches[platform] = null; // Ensure match is null on error
      }
    }
    matchResults.push(result);
  });
  return matchResults;
}


// ---------------- MENU FUNCTIONS ---------------- 
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Price Management')
    .addItem('Safe Setup System (Preserve Data)', 'safeSetupPriceManagementSystem')
    .addSeparator()
    .addItem('Run Accurate SKU Matching', 'runSkuMatching')
    .addItem('Validate Match Quality', 'validateMatchQuality')
    .addItem('Save Matches to History', 'saveMatches')
    .addSeparator()
    .addItem('Update Price Analysis', 'updatePriceAnalysis')
    .addItem('Generate Export Files', 'generateExports')
    .addSeparator()
    .addItem('Generate CP Listings', 'generateCPListings')
    .addItem('Identify Discontinued Items', 'identifyDiscontinuedItems')
    .addToUi();
}

// ---------------- SETUP FUNCTIONS ---------------- 
function safeSetupPriceManagementSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Safe Setup', 'Create missing tabs & preserve data?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const existingSheets = {};
  SS.getSheets().forEach(sheet => { existingSheets[sheet.getName()] = true; });

  if (!existingSheets['Manufacturer Price Sheet']) createManufacturerPriceSheet();
  if (!existingSheets['SKU Matching Engine']) createSkuMatchingEngineTab();
  if (!existingSheets['Price Analysis Dashboard']) createPriceAnalysisDashboard();
  if (!existingSheets['Match History']) createMatchHistoryTab();
  if (!existingSheets['Instructions']) createInstructionsTab();
  if (!existingSheets['CP Listings']) createCPListingsTab();
  if (!existingSheets['Discontinued']) createDiscontinuedTab();
  if (!existingSheets['Price Changes']) createPriceChangesTab(); // Added for the new tab
  if (!existingSheets['Error Log']) initializeErrorLogSheet(); // Ensure Error Log is set up


  Object.keys(PLATFORM_TABS).forEach(platform => {
    const tabName = PLATFORM_TABS[platform].name;
    if (!existingSheets[tabName]) createExportTab(platform);
  });
  ui.alert('Setup Complete', 'Missing tabs created/verified. Data preserved.', ui.ButtonSet.OK);
}

function createManufacturerPriceSheet() {
  const sheet = SS.insertSheet('Manufacturer Price Sheet');
  const headers = ['Manufacturer SKU', 'UPC', 'MSRP', 'MAP', 'Dealer Price'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, headers.length).merge().setValue('Paste MFR price sheet data below. Ensure "Manufacturer SKU", "MAP", and "Dealer Price" columns are present.').setFontStyle('italic').setHorizontalAlignment('center');
  sheet.getRange(3, 1, 100, headers.length).setBackground('#f3f3f3'); // Example styling for data area
  sheet.setColumnWidth(1, 200); // MFR SKU
  sheet.setColumnWidth(2, 150); // UPC
  sheet.setColumnWidth(3, 100); // MSRP
  sheet.setColumnWidth(4, 100); // MAP
  sheet.setColumnWidth(5, 100); // Dealer Price
}

function createSkuMatchingEngineTab() {
  const sheet = SS.insertSheet('SKU Matching Engine');
  const headers = ['Manufacturer SKU', 'MSRP', 'MAP', 'Dealer Price'];
  Object.keys(PLATFORM_TABS).forEach(platform => {
    headers.push(`${platform} SKU`, `${platform} Confidence`, `${platform} Current Price`);
  });
  headers.push('Status', 'Notes'); // For overall status and manual notes
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, headers.length).merge().setValue('Matches shown here. Use "Run Accurate SKU Matching" to populate.').setFontStyle('italic').setHorizontalAlignment('center');
  // Set column widths
  sheet.setColumnWidth(1, 200); // MFR SKU
  sheet.setColumnWidth(2, 80);  // MSRP
  sheet.setColumnWidth(3, 80);  // MAP
  sheet.setColumnWidth(4, 80);  // Dealer Price
  let colIndex = 5;
  Object.keys(PLATFORM_TABS).forEach(() => {
    sheet.setColumnWidth(colIndex, 200);     // Platform SKU
    sheet.setColumnWidth(colIndex + 1, 100); // Confidence
    sheet.setColumnWidth(colIndex + 2, 100); // Current Price
    colIndex += 3;
  });
  sheet.setColumnWidth(colIndex, 100); // Status
  sheet.setColumnWidth(colIndex + 1, 200); // Notes
}

function createPriceAnalysisDashboard() {
    const sheet = SS.insertSheet('Price Analysis Dashboard');
    // Title
    sheet.getRange(1, 1, 1, 12).merge().setValue('PRICE ANALYSIS DASHBOARD')
        .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center').setBackground('#e0e0e0');
    // Summary Metrics Area
    sheet.getRange(2, 1, 1, 12).merge().setValue('SUMMARY METRICS')
        .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f3f3f3');
    sheet.getRange(3,1).setValue('Total Products:'); sheet.getRange(3,2).setValue('0');
    sheet.getRange(3,4).setValue('Price Increases:'); sheet.getRange(3,5).setValue('0');
    sheet.getRange(3,7).setValue('Price Decreases:'); sheet.getRange(3,8).setValue('0');
    sheet.getRange(3,10).setValue('Average Change:'); sheet.getRange(3,11).setValue('0%');
    sheet.getRange(4,1).setValue('MAP Violations:'); sheet.getRange(4,2).setValue('0');
    sheet.getRange(4,4).setValue('High Impact Changes:');sheet.getRange(4,5).setValue('0');
    sheet.getRange(4,7).setValue('Unmatched Products:');sheet.getRange(4,8).setValue('0');
    sheet.getRange(4,10).setValue('B-Stock Changes:');sheet.getRange(4,11).setValue('0');
    // Bold labels for summary
    sheet.getRange(3,1,2,1).setFontWeight('bold'); sheet.getRange(3,4,2,1).setFontWeight('bold');
    sheet.getRange(3,7,2,1).setFontWeight('bold'); sheet.getRange(3,10,2,1).setFontWeight('bold');

    // Headers for detailed analysis
    const analysisHeaders = ['SKU', 'MAP', 'Dealer', 'B-Stock', 'Platform', 'Current', 'New', 'Change $', 'Change %', 'Status'];
    sheet.getRange(6, 1, 1, analysisHeaders.length).setValues([analysisHeaders])
        .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(5,1,1,analysisHeaders.length).merge().setValue('Use "Update Price Analysis" to refresh.')
        .setFontStyle('italic').setHorizontalAlignment('center');

    // Column Widths for analysis table
    sheet.setColumnWidth(1, 180); // SKU
    sheet.setColumnWidth(2, 80);  // MAP
    sheet.setColumnWidth(3, 80);  // Dealer
    sheet.setColumnWidth(4, 80);  // B-Stock
    sheet.setColumnWidth(5, 100); // Platform
    sheet.setColumnWidth(6, 100); // Current Price
    sheet.setColumnWidth(7, 100); // New Price
    sheet.setColumnWidth(8, 100); // Change $
    sheet.setColumnWidth(9, 100); // Change %
    sheet.setColumnWidth(10, 120); // Status
    sheet.getRange(7,1,100,analysisHeaders.length).setBackground('#f8f9fa'); // Example styling
}

function createMatchHistoryTab() {
  const sheet = SS.insertSheet('Match History');
  const headers = ['Manufacturer SKU', 'Platform', 'Platform SKU', 'Match Type', 'Confidence', 'Date Added', 'User Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2,1,1,headers.length).merge().setValue('Confirmed SKU matches. Use "Save Matches" to populate.').setFontStyle('italic').setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 100); sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 100); sheet.setColumnWidth(6, 120); sheet.setColumnWidth(7, 200);
  sheet.getRange(3,1,100,headers.length).setBackground('#f8f9fa');
}

function createInstructionsTab() {
  const sheet = SS.insertSheet('Instructions');
  const instructions = [
    ['CROSS-PLATFORM PRICE MANAGEMENT SYSTEM'],
    ['This system helps match manufacturer SKUs to various online sales platforms and analyze pricing.'],
    ['1. Manufacturer Price Sheet: Paste your master price list here. Ensure "Manufacturer SKU", "MAP", and "Dealer Price" columns exist.'],
    ['2. Platform Databases: (If used directly) This sheet should contain exports from your sales platforms. Column headers should identify the platform (e.g., "Amazon SKU", "eBay Price"). The script attempts to auto-detect these.'],
    ['3. Run Accurate SKU Matching: From the "Price Management" menu, this will populate the "SKU Matching Engine" tab.'],
    ['4. SKU Matching Engine: Review matches here. Confidence scores indicate match quality. Status helps track items.'],
    ['5. Update Price Analysis: Analyzes prices based on new MAP/Dealer prices and current platform prices. Updates "Price Analysis Dashboard" and "Price Changes" tab.'],
    ['6. Generate Export Files: Creates formatted files for each platform based on the price analysis.'],
    ['Other functions: Save Matches (to history), Validate Quality (re-checks matches), CP Listings (finds unlisted items), Identify Discontinued.'],
    ['Error Log: Check this tab for any errors encountered during script execution.']
  ];
  for (let i = 0; i < instructions.length; i++) {
    sheet.getRange(i + 1, 1).setValue(instructions[i][0]);
    if(i>1) sheet.getRange(i+1,1).setWrap(true);
  }
  sheet.getRange(1,1).setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 800); // Wider for instructions
}
function createCPListingsTab() {
    const sheet = SS.insertSheet('CP Listings');
    const headers = ['Manufacturer SKU', 'UPC', 'MSRP', 'MAP', 'Dealer Price', 'Amazon', 'eBay', 'Shopify', 'Reverb', 'inFlow', 'SellerCloud', 'Action Needed'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(2,1,1,headers.length).merge().setValue('Items from MFR sheet needing listing on one or more platforms.').setFontStyle('italic').setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 100);sheet.setColumnWidth(4, 100);sheet.setColumnWidth(5, 100);
    for (let i=6; i<=11; i++) sheet.setColumnWidth(i,100); // Platform columns
    sheet.setColumnWidth(12, 200); // Action needed
    sheet.getRange(3,1,100,headers.length).setBackground('#f8f9fa');
}
function createDiscontinuedTab() {
    const sheet = SS.insertSheet('Discontinued');
    const headers = ['Platform SKU', 'Platform', 'Brand', 'Current Price', 'Last Updated', 'Confidence', 'Status', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(2,1,1,headers.length).merge().setValue('Platform SKUs not found in current MFR sheet (potential discontinued).').setFontStyle('italic').setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 100); sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 120); sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 150); sheet.setColumnWidth(8, 250);
    sheet.getRange(3,1,100,headers.length).setBackground('#f8f9fa');
    return sheet;
}

function createExportTab(platform) {
  const platformInfo = PLATFORM_TABS[platform];
  if (!platformInfo) {
      logError('createExportTab', `Platform info not found for ${platform}`);
      return;
  }
  const tabName = platformInfo.name;
  const headers = platformInfo.headers;
  const sheet = SS.insertSheet(tabName);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLORS[platform] || '#cccccc').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, headers.length).merge()
    .setValue('NOT READY - Run "Generate Export Files"').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f4b400').setFontColor('#000000');
  sheet.getRange(3, 1, 1, headers.length).merge()
    .setValue('Data ready to export to ' + tabName.replace(' Export', '')).setFontStyle('italic').setHorizontalAlignment('center');
  for (let i = 0; i < headers.length; i++) {
    sheet.setColumnWidth(i + 1, 150);
  }
  sheet.getRange(4,1,100,headers.length).setBackground('#f8f9fa');
}

function createPriceChangesTab() {
  const sheetName = "Price Changes";
  let sheet = SS.getSheetByName(sheetName);
  if (sheet) {
    Logger.log(`Sheet "${sheetName}" already exists. Content will be overwritten or appended.`);
  } else {
    sheet = SS.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" created.`);
  }

  sheet.clearContents(); // Always clear for fresh data

  // Price Increases Section
  sheet.getRange("A1").setValue("PRICE INCREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_UP);
  const headersUp = ["MFR SKU", "Platform", "Old Price", "New Price", "Change $", "Change %"];
  sheet.getRange(2, 1, 1, headersUp.length).setValues([headersUp]).setFontWeight("bold").setBackground("#f2f2f2");
  sheet.setColumnWidth(1, 200); // MFR SKU
  sheet.setColumnWidth(2, 100); // Platform
  sheet.setColumnWidth(3, 100); // Old Price
  sheet.setColumnWidth(4, 100); // New Price
  sheet.setColumnWidth(5, 100); // Change $
  sheet.setColumnWidth(6, 100); // Change %

  // Determine start row for decreases dynamically, ensuring space
  // This calculation will be done when populating, as last row changes
  // For setup, just define the header section for decreases conceptually
  const placeholderStartRowDecreases = 50; // Placeholder, actual row determined in populatePriceChangesTab
  sheet.getRange(placeholderStartRowDecreases, 1).setValue("PRICE DECREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_DOWN);
  const headersDown = ["MFR SKU", "Platform", "Old Price", "New Price", "Change $", "Change %"];
  sheet.getRange(placeholderStartRowDecreases + 1, 1, 1, headersDown.length).setValues([headersDown]).setFontWeight("bold").setBackground("#f2f2f2");
  
  return sheet;
}


// ---------------- SKU MATCHING EXECUTION (SINGLE RUN) ----------------
function runSkuMatching() {
  const ui = SpreadsheetApp.getUi();
  initializeErrorLogSheet(); // Ensure error log sheet is ready
  const response = ui.alert(
    'Run Accurate SKU Matching',
    'This will match all manufacturer SKUs to platform SKUs using accuracy-focused algorithms. This may take some time for large datasets. Previous results on "SKU Matching Engine" will be cleared. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const matchingSheet = SS.getSheetByName('SKU Matching Engine');
  if (!matchingSheet) {
    logError('runSkuMatching', 'SKU Matching Engine tab not found. Run Safe Setup System first.', '', '', true, ui);
    return;
  }

  // Clear previous results more thoroughly
  const lastRowContent = matchingSheet.getLastRow();
  if (lastRowContent >= 3) { // Header is row 1, status/message is row 2
    matchingSheet.getRange(3, 1, lastRowContent - 2, matchingSheet.getLastColumn()).clearContent().clearDataValidations().clearNote();
  }
   if (matchingSheet.getMaxRows() > 2 && matchingSheet.getMaxRows() > (lastRowContent || 2) ) { 
     // Clear any rows below the actual content that might have old data/formatting
     matchingSheet.getRange((lastRowContent || 2) + 1, 1, matchingSheet.getMaxRows() - (lastRowContent || 2), matchingSheet.getLastColumn()).clearContent().clearDataValidations().clearNote();
   }


  const statusCell = matchingSheet.getRange(2, 1, 1, matchingSheet.getLastColumn());
  statusCell.merge().setValue('PROCESSING - Running SKU matching...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center');
  SpreadsheetApp.flush();

  try {
    const startTime = new Date().getTime();
    Object.keys(skuNormalizeCache).forEach(key => delete skuNormalizeCache[key]); // Clear global normalization cache

    statusCell.setValue('PROCESSING - Loading manufacturer data...'); SpreadsheetApp.flush();
    const manufacturerData = getManufacturerData(); // From "Manufacturer Price Sheet"
    if (!manufacturerData || manufacturerData.length === 0) {
      statusCell.setValue('ERROR - No manufacturer data found in "Manufacturer Price Sheet".').setBackground('#f4cccc');
      logError('runSkuMatching', 'No manufacturer data found.', '', '', true, ui);
      return;
    }

    const mfrTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`PROCESSING - Loaded ${manufacturerData.length} MFR SKUs in ${mfrTime.toFixed(1)}s. Pre-processing platform data...`); SpreadsheetApp.flush();
    
    const rawPlatformData = getPlatformDataFromStructuredSheet(); // From "Platform Databases"
     if (!rawPlatformData || Object.keys(rawPlatformData).length === 0 || !Object.values(rawPlatformData).some(p => p.length > 0)) {
      statusCell.setValue('ERROR - No platform data found or parsed from "Platform Databases" sheet.').setBackground('#f4cccc');
      logError('runSkuMatching', 'No platform data found or parsed.', '', '', true, ui);
      return;
    }

    const { preProcessedPlatformDataWithAttributes, platformSkuMaps /*, cleanSkuMaps - removed */ } = preProcessAllPlatformData(rawPlatformData);
    
    let totalPlatformSkus = 0; 
    for (const platform in preProcessedPlatformDataWithAttributes) {
        totalPlatformSkus += preProcessedPlatformDataWithAttributes[platform].length;
    }

    const dataLoadTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`PROCESSING - Loaded & pre-processed ${totalPlatformSkus} platform SKUs in ${dataLoadTime.toFixed(1)}s (total). Starting matching...`); SpreadsheetApp.flush();

    const matchResults = performFullMatching(manufacturerData, preProcessedPlatformDataWithAttributes, platformSkuMaps /*, cleanSkuMaps - removed */);

    let totalMatches = 0, exactMatches = 0, highConfidenceMatches = 0, mediumConfidenceMatches = 0, lowConfidenceMatches = 0, reviewRequiredMatches = 0;
    matchResults.forEach(result => {
      Object.values(result.matches).forEach(match => {
        if (match) {
          totalMatches++;
          const confidence = match.confidenceScore;
          if (confidence >= 95) exactMatches++;
          else if (confidence >= 85) highConfidenceMatches++;
          else if (confidence >= 70) mediumConfidenceMatches++; // 70-84
          else lowConfidenceMatches++; // Below 70
          
          // Check for review needed based on matchType or confidence score range
          if (match.matchType && (match.matchType.includes('REVIEW') || (confidence < 85 && confidence >= 70))) {
            reviewRequiredMatches++;
          }
        }
      });
    });

    const matchingTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`PROCESSING - Found ${totalMatches} potential matches in ${matchingTime.toFixed(1)}s. Updating sheet...`); SpreadsheetApp.flush();
    
    appendMatchResultsToSheet(matchResults, matchingSheet); // Clears and appends in one go now
    
    const finalRowCount = matchingSheet.getLastRow() - 2; // Data starts at row 3
    if (finalRowCount > 0) {
        addConditionalFormattingToMatchingSheet(matchingSheet, finalRowCount);
    }

    const totalTime = (new Date().getTime() - startTime) / 1000;
    // *** FIXED: reviewRequiredMatchesOverall changed to reviewRequiredMatches ***
    statusCell.setValue(`MATCHING COMPLETE - ${matchResults.length} MFR SKUs processed in ${totalTime.toFixed(1)}s. ${totalMatches} matches. ${reviewRequiredMatches} for review.`).setBackground('#d9ead3');
    ui.alert('Accurate SKU Matching Complete',
             `${matchResults.length} MFR SKUs processed in ${totalTime.toFixed(1)}s.\n` +
             `Total matches: ${totalMatches}\n` +
             `Exact (95-100%): ${exactMatches}\nHigh (85-94%): ${highConfidenceMatches}\n` +
             `Medium (70-84%): ${mediumConfidenceMatches}\nLow (<70%): ${lowConfidenceMatches}\n` +
             `${reviewRequiredMatches} matches require review.`, ui.ButtonSet.OK);

  } catch (error) {
    logError('runSkuMatching', 'Main process failed', '', error, true, ui);
    statusCell.setValue('ERROR - ' + error.toString()).setBackground('#f4cccc');
  }
}


// ---------------- DATA FETCHING FUNCTIONS ----------------
function getManufacturerData() {
  const sheetName = 'Manufacturer Price Sheet';
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`${sheetName} tab not found`);
  }
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return []; // No data

  const headers = values[0].map(h => String(h).trim());
  const skuIndex = headers.indexOf('Manufacturer SKU');
  const upcIndex = headers.indexOf('UPC');
  const msrpIndex = headers.indexOf('MSRP');
  const mapIndex = headers.indexOf('MAP'); // Essential
  const dealerPriceIndex = headers.indexOf('Dealer Price'); // Essential

  if (skuIndex === -1) throw new Error('"Manufacturer SKU" column not found in MFR Sheet');
  if (mapIndex === -1) throw new Error('"MAP" column not found in MFR Sheet');
  if (dealerPriceIndex === -1) throw new Error('"Dealer Price" column not found in MFR Sheet');

  const data = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sku = row[skuIndex];
    if (!sku) continue; // Skip rows without SKU

    try {
        let dealerPriceVal = row[dealerPriceIndex];
        let mapVal = row[mapIndex];

        let dealerPrice = !isNaN(parseFloat(dealerPriceVal)) ? parseFloat(dealerPriceVal) : null;
        let map = !isNaN(parseFloat(mapVal)) ? parseFloat(mapVal) : null;
        
        // Fallback for MAP if empty/zero but dealer price is valid
        if ((!map || map === 0) && dealerPrice && dealerPrice > 0) {
            map = parseFloat((dealerPrice / 0.85).toFixed(2)); // Example: MAP is 15% above dealer
        }


        data.push({
            manufacturerSku: String(sku),
            upc: upcIndex >= 0 && row[upcIndex] ? String(row[upcIndex]) : "",
            msrp: msrpIndex >= 0 && !isNaN(parseFloat(row[msrpIndex])) ? parseFloat(row[msrpIndex]) : null,
            map: map,
            dealerPrice: dealerPrice
        });
    } catch (e) {
        logError('getManufacturerData', `Error processing row ${i+1} for SKU ${sku}`, sku, e);
    }
  }
  return data;
}

function getPlatformDataFromStructuredSheet() {
  const sheetName = 'Platform Databases'; // Standardized sheet name
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    logError('getPlatformDataFromStructuredSheet', `${sheetName} tab not found. Please create it or ensure the name is exact.`);
    return {}; // Return empty object if sheet not found
  }
  const allSheetValues = sheet.getDataRange().getValues();
  if (allSheetValues.length <= 1) { // Need at least header row + data
      logError('getPlatformDataFromStructuredSheet', `No data found in ${sheetName}.`);
      return {};
  }

  const platformData = { AMAZON: [], EBAY: [], INFLOW: [], REVERB: [], SHOPIFY: [], SELLERCLOUD: [] };
  const platformHeaderIndices = {}; // Stores the starting column index for each platform's data block

  // --- Flexible Header Detection ---
  // Attempt to find platform headers in the first few rows
  // Example: "Amazon" or "Amazon SKU" could be in row 1, "seller-sku" in row 2 under that block.
  const headerRowsToScan = Math.min(5, allSheetValues.length); // Scan up to 5 rows for headers
  const mainPlatformHeaders = Object.keys(PLATFORM_TABS); // AMAZON, EBAY, etc.

  // Find main platform blocks first (e.g., "AMAZON" column header)
  for (let col = 0; col < allSheetValues[0].length; col++) {
      for (let row = 0; row < headerRowsToScan; row++) {
          const cellValue = String(allSheetValues[row][col]).toUpperCase().trim();
          for (const platformKey of mainPlatformHeaders) {
              if (cellValue.includes(platformKey) && platformHeaderIndices[platformKey] === undefined) { // Found a platform block
                  platformHeaderIndices[platformKey] = col; // Mark the starting column
                  break; 
              }
          }
      }
  }
  
  // For each identified platform block, find its specific sub-headers (SKU, Price, Cost, etc.)
  let dataStartRow = -1; // Dynamically determine where actual data starts

  for (const platform in platformHeaderIndices) {
    const startCol = platformHeaderIndices[platform];
    if (startCol === undefined) {
      Logger.log(`Warning: Platform ${platform} main header not found in '${sheetName}' header rows.`);
      continue;
    }

    let skuColOffset = -1, priceColOffset = -1, costColOffset = -1, conditionColOffset = -1;
    // Clean Sku is intentionally not sought here as per user request to ignore it.
    
    // Scan for sub-headers within this platform's column range and a few rows down
    for (let hRow = 0; hRow < headerRowsToScan; hRow++) {
        // Iterate columns only within the presumed block of the current platform or slightly beyond if not structured perfectly
        // Stop if we hit the start of another known platform's block
        for (let hCol = startCol; hCol < allSheetValues[hRow].length; hCol++) {
            // Check if hCol is the start of another platform block to avoid reading its headers
            let isAnotherPlatformBlock = false;
            for(const otherPlat in platformHeaderIndices){
                if(otherPlat !== platform && platformHeaderIndices[otherPlat] === hCol) {
                    isAnotherPlatformBlock = true;
                    break;
                }
            }
            if(isAnotherPlatformBlock && hCol > startCol) break; // Moved into another platform's territory

            const colName = String(allSheetValues[hRow][hCol]).trim().toLowerCase(); // Standardize for matching
            const offset = hCol - startCol;

            // Update dataStartRow if this is the deepest header row found so far
            if (dataStartRow < hRow + 1) dataStartRow = hRow + 1;

            switch (platform) {
              case 'AMAZON':
                if (colName === 'seller-sku') skuColOffset = offset;
                if (colName === 'price') priceColOffset = offset;
                break;
              case 'EBAY':
                if (colName === 'custom label (sku)' || colName === 'custom label') skuColOffset = offset;
                if (colName === 'start price') priceColOffset = offset;
                break;
              case 'INFLOW':
                if (colName === 'name' || colName === 'product name') skuColOffset = offset; // 'Name' is often SKU in InFlow
                if (colName === 'unitprice' || colName === 'unit price') priceColOffset = offset;
                if (colName === 'cost') costColOffset = offset;
                break;
              case 'REVERB':
                if (colName === 'sku') skuColOffset = offset;
                if (colName === 'price') priceColOffset = offset;
                if (colName === 'condition') conditionColOffset = offset;
                break;
              case 'SHOPIFY':
                if (colName === 'variant sku') skuColOffset = offset;
                if (colName === 'variant price') priceColOffset = offset;
                if (colName === 'variant cost') costColOffset = offset; // Or 'cost per item'
                break;
              case 'SELLERCLOUD':
                if (colName === 'productid' || colName === 'product id') skuColOffset = offset;
                if (colName === 'siteprice' || colName === 'site price') priceColOffset = offset;
                if (colName === 'sitecost' || colName === 'site cost') costColOffset = offset;
                break;
            }
        }
    }
    
    if (dataStartRow === -1 || dataStartRow >= allSheetValues.length) {
        logError('getPlatformDataFromStructuredSheet', `Could not determine data start row or no data rows in ${sheetName}.`);
        continue; // Skip this platform if data start is invalid
    }


    if (skuColOffset === -1) {
      Logger.log(`Warning: SKU column not found for platform ${platform} in '${sheetName}'. This platform will be skipped.`);
      continue;
    }

    // Extract data for the platform
    for (let rowIdx = dataStartRow; rowIdx < allSheetValues.length; rowIdx++) {
      const rowData = allSheetValues[rowIdx];
      const sku = (startCol + skuColOffset < rowData.length) ? String(rowData[startCol + skuColOffset]).trim() : null;
      if (!sku) continue; // Skip if no SKU

      try {
          let price = null;
          if (priceColOffset !== -1 && (startCol + priceColOffset < rowData.length) && rowData[startCol + priceColOffset] !== '') {
              const rawPrice = rowData[startCol + priceColOffset];
              price = !isNaN(parseFloat(rawPrice)) ? parseFloat(rawPrice) : null;
              if (price === null) Logger.log(`Warning: Invalid price '${rawPrice}' for ${platform} SKU ${sku}.`);
          }
          
          let cost = null;
          if (costColOffset !== -1 && (startCol + costColOffset < rowData.length) && rowData[startCol + costColOffset] !== '') {
              const rawCost = rowData[startCol + costColOffset];
              cost = !isNaN(parseFloat(rawCost)) ? parseFloat(rawCost) : null;
               if (cost === null) Logger.log(`Warning: Invalid cost '${rawCost}' for ${platform} SKU ${sku}.`);
          }
          
          let condition = null;
          if (conditionColOffset !== -1 && (startCol + conditionColOffset < rowData.length)) {
              condition = String(rowData[startCol + conditionColOffset]).trim();
          }

          platformData[platform].push({
            platform: platform,
            sku: sku,
            price: price,
            cost: cost,
            // cleanSku: null, // Not extracted
            condition: condition
          });
      } catch (e) {
          logError('getPlatformDataFromStructuredSheet (data extraction)', `Error processing row ${rowIdx +1} for ${platform} SKU ${sku}`, sku, e);
      }
    }
  }
  return platformData;
}


// ---------------- OUTPUT & FORMATTING ---------------- 
function appendMatchResultsToSheet(batchMatchResults, sheet) {
  if (!batchMatchResults || batchMatchResults.length === 0) {
    Logger.log("No results in this batch to append.");
    return;
  }

  // Clear previous results from row 3 downwards
  const lastRowContent = sheet.getLastRow();
  if (lastRowContent >= 3) {
    sheet.getRange(3, 1, lastRowContent - 2, sheet.getLastColumn()).clearContent().clearDataValidations().clearNote();
  }
   if (sheet.getMaxRows() > 2 && sheet.getMaxRows() > (lastRowContent || 2) ) { 
     sheet.getRange((lastRowContent || 2) + 1, 1, sheet.getMaxRows() - (lastRowContent || 2), sheet.getLastColumn()).clearContent().clearDataValidations().clearNote();
   }


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowsToAppend = [];

  batchMatchResults.forEach(result => {
    const row = Array(headers.length).fill(''); // Initialize empty row
    headers.forEach((header, colIndex) => {
      try {
          if (header === 'Manufacturer SKU') row[colIndex] = result.manufacturerSku;
          else if (header === 'MSRP') row[colIndex] = result.msrp !== null && !isNaN(result.msrp) ? result.msrp : '';
          else if (header === 'MAP') row[colIndex] = result.map !== null && !isNaN(result.map) ? result.map : '';
          else if (header === 'Dealer Price') row[colIndex] = result.dealerPrice !== null && !isNaN(result.dealerPrice) ? result.dealerPrice : '';
          else if (header === 'Status') {
            // Determine status based on matches
            const hasAutoMatch = Object.values(result.matches).some(match => match && match.confidenceScore >= 85);
            const needsReview = Object.values(result.matches).some(match => match && match.matchType && (match.matchType.includes('REVIEW') || (match.confidenceScore < 85 && match.confidenceScore >= 70)));
            const hasAnyMatch = Object.values(result.matches).some(match => match);

            if (needsReview) row[colIndex] = 'Review Required';
            else if (hasAutoMatch) row[colIndex] = 'Auto-matched';
            else if (hasAnyMatch) row[colIndex] = 'Partial Matches'; // Low confidence, but some match found
            else row[colIndex] = 'No Match';
          }
          else if (header === 'Notes') row[colIndex] = ''; // Placeholder for manual notes
          else {
            const parts = header.split(' '); // e.g. "AMAZON SKU", "EBAY Confidence"
            if (parts.length >= 2) {
              const platform = parts[0].toUpperCase(); // Ensure platform key matches (AMAZON, EBAY)
              const field = parts.slice(1).join(' '); // "SKU", "Confidence", "Current Price"
              
              if (result.matches[platform]) {
                const match = result.matches[platform];
                if (field === 'SKU') row[colIndex] = match.platformSku;
                else if (field === 'Confidence') row[colIndex] = match.confidenceScore !== null && !isNaN(match.confidenceScore) ? parseFloat(match.confidenceScore.toFixed(2)) : '';
                else if (field === 'Current Price') row[colIndex] = match.currentPrice !== null && !isNaN(match.currentPrice) ? parseFloat(match.currentPrice.toFixed(2)) : '';
              }
            }
          }
      } catch (e) {
          logError('appendMatchResultsToSheet (cell population)', `Error populating cell for header "${header}" for MFR SKU ${result.manufacturerSku}`, result.manufacturerSku, e);
          row[colIndex] = 'ERROR_IN_CELL'; // Mark cell with error
      }
    });
    rowsToAppend.push(row);
  });

  if (rowsToAppend.length > 0) {
    const startRowForAppend = 3; // Data always starts at row 3 after clearing
    sheet.getRange(startRowForAppend, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    Logger.log(`Appended ${rowsToAppend.length} rows to SKU Matching Engine.`);
  }
}

function addConditionalFormattingToMatchingSheet(sheet, rowCount) {
  if (!sheet || rowCount <= 0) return;
  try {
      sheet.clearConditionalFormatRules(); // Clear existing rules for this sheet
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rules = [];
      const dataStartRow = 3; // Data starts at row 3

      headers.forEach((header, index) => {
        if (header.includes('Confidence')) {
          const range = sheet.getRange(dataStartRow, index + 1, rowCount, 1);
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(95).setBackground('#d9ead3').setRanges([range]).build()); // Green - Exact/High
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(85, 94.99).setBackground('#cfe2f3').setRanges([range]).build());      // Blue - Good
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(70, 84.99).setBackground('#fff2cc').setRanges([range]).build());      // Yellow - Needs Review
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).whenNumberLessThan(70).setBackground('#f4cccc').setRanges([range]).build()); // Red - Low/No
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#efefef').setRanges([range]).build());           // Grey - No match / 0 confidence
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#ffffff').setRanges([range]).build()); // White - Empty
        }
      });

      const statusColIndex = headers.indexOf('Status');
      if (statusColIndex >= 0) {
        const range = sheet.getRange(dataStartRow, statusColIndex + 1, rowCount, 1);
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Review Required').setBackground(COLORS.WARNING).setRanges([range]).build()); // Yellow
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Partial Matches').setBackground('#fff2cc').setRanges([range]).build()); // Light Yellow
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Auto-matched').setBackground(COLORS.POSITIVE).setFontColor('#FFFFFF').setRanges([range]).build()); // Green
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('No Match').setBackground('#efefef').setRanges([range]).build()); // Grey
      }
      if (rules.length > 0) {
        sheet.setConditionalFormatRules(rules);
        Logger.log("Applied conditional formatting to SKU Matching Engine.");
      }
  } catch(e) {
      logError('addConditionalFormattingToMatchingSheet', 'Failed to apply conditional formatting', '', e);
  }
}

// ---------------- MATCH REVIEW & UPDATE (SIDEBAR) - REMOVED ----------------
// The functions showMatchReviewSidebar, getMatchesForReviewAccurate, getConservativeSkuSuggestions_optimized, and updateMatch
// have been removed as this feature is no longer needed. Manual review will be done directly on the "SKU Matching Engine" sheet.

// ---------------- OTHER CORE FUNCTIONS ----------------
function validateMatchQuality() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Validate Match Quality', 'Analyze existing matches in "SKU Matching Engine" and flag potentially incorrect ones based on current logic. This may add notes or change status. Continue?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    if (!matchingSheet) { 
        logError('validateMatchQuality', 'SKU Matching Engine tab not found.', '', '', true, ui);
        return;
    }

    const matchingData = matchingSheet.getDataRange().getValues();
    if (matchingData.length < 3) {
        ui.alert('Info', 'No data to validate in SKU Matching Engine.', ui.ButtonSet.OK);
        return;
    }
    const headers = matchingData[0];
    const mfrSkuIndex = headers.indexOf('Manufacturer SKU');
    const statusIndex = headers.indexOf('Status');
    const notesIndex = headers.indexOf('Notes');

    if (mfrSkuIndex === -1) { 
        logError('validateMatchQuality', 'Manufacturer SKU column not found.', '', '', true, ui);
        return;
    }
    if (notesIndex === -1) {
        logError('validateMatchQuality', '"Notes" column not found. Please ensure it exists for quality feedback.', '', '', true, ui);
        return; // Notes column is essential for this function
    }


    const platformColumns = {}; // Stores {platformName: sku_col_index}
    const platformConfidenceColumns = {}; // Stores {platformName: confidence_col_index}
    headers.forEach((header, index) => {
      if (header.includes(' SKU') && !header.includes('Manufacturer')) {
          platformColumns[header.split(' ')[0].toUpperCase()] = index;
      }
      if (header.includes(' Confidence')) {
          platformConfidenceColumns[header.split(' ')[0].toUpperCase()] = index;
      }
    });

    let totalMatchesChecked = 0, suspiciousMatches = 0, goodMatches = 0;
    const suspiciousRowsInfo = []; // To update notes later in batch

    // Start from row 3 (index 2) as row 1 is header, row 2 is status
    for (let i = 2; i < matchingData.length; i++) {
      const row = matchingData[i];
      const mfrSkuRaw = row[mfrSkuIndex];
      if (!mfrSkuRaw) continue; // Skip empty MFR SKU rows

      const mfrAttributes = extractSkuAttributesAndCore(mfrSkuRaw);

      for (const platform in platformColumns) {
        const platformSkuRaw = row[platformColumns[platform]];
        const currentConfidenceRaw = platformConfidenceColumns[platform] !== undefined ? row[platformConfidenceColumns[platform]] : '0';
        const currentConfidence = parseFloat(currentConfidenceRaw);

        if (!platformSkuRaw || isNaN(currentConfidence) || currentConfidence < 1) continue; // Skip if no platform SKU or confidence is invalid/zero

        totalMatchesChecked++;
        const platformAttributes = extractSkuAttributesAndCore(platformSkuRaw);
        const validation = validateMatch_optimized(mfrSkuRaw, mfrAttributes, platformSkuRaw, platformAttributes, platform);

        // Flag if current high confidence (>=85) now re-validates to low (<85)
        // Or if current medium (70-84) now re-validates significantly lower (<70 or much lower in medium)
        let flagAsSuspicious = false;
        if (currentConfidence >= 85 && validation.confidence < 85) {
            flagAsSuspicious = true;
        } else if (currentConfidence >= 70 && validation.confidence < 70) { // Was reviewable, now low
             flagAsSuspicious = true;
        }


        if (flagAsSuspicious) {
          suspiciousMatches++;
          suspiciousRowsInfo.push({
            rowNum: i + 1, // 1-based row index for sheet
            platform: platform, mfrSku: mfrSkuRaw, platformSku: platformSkuRaw,
            originalConfidence: currentConfidence.toFixed(1), newConfidence: validation.confidence.toFixed(1), 
            reason: validation.reason
          });
        } else if (validation.confidence >= 85) {
          goodMatches++;
        }
      }
    }

    // Update notes and status for suspicious matches
    suspiciousRowsInfo.forEach(info => {
      let currentNotes = matchingSheet.getRange(info.rowNum, notesIndex + 1).getValue().toString();
      const warningNote = `QUALITY CHECK (${new Date().toLocaleDateString()}): ${info.platform} match (${info.platformSku}) re-validated to ${info.newConfidence}% (was ${info.originalConfidence}%). Reason: ${info.reason}`;
      
      // Avoid duplicate quality check notes for the same platform on the same row
      const qualityCheckPrefix = `QUALITY CHECK (${new Date().toLocaleDateString()}): ${info.platform}`;
      if (!currentNotes.includes(qualityCheckPrefix)) {
        matchingSheet.getRange(info.rowNum, notesIndex + 1).setValue((currentNotes ? currentNotes + '; ' : '') + warningNote);
        if (statusIndex !== -1) { // Update status if Status column exists
            matchingSheet.getRange(info.rowNum, statusIndex + 1).setValue('Quality Review Required');
        }
      }
    });

    if (suspiciousRowsInfo.length > 0) {
        applyQualityValidationFormatting(matchingSheet); // Apply specific formatting for these rows
    }

    let message = `Match Quality Validation Complete\nTotal Matches Checked: ${totalMatchesChecked}, Good Re-validations: ${goodMatches}, Suspicious: ${suspiciousMatches}\n`;
    if (suspiciousMatches > 0) message += `Flagged matches updated in Notes & Status. Please review.`;
    else message += `All existing high-confidence matches seem consistent with current logic.`;
    ui.alert('Validation Complete', message, ui.ButtonSet.OK);

  } catch (error) { 
      logError('validateMatchQuality', 'Error during validation', '', error, true, ui);
  }
}

function applyQualityValidationFormatting(sheet) {
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const notesIndex = headers.indexOf('Notes');
    const statusIndex = headers.indexOf('Status');
    const dataStartRow = 3;
    const numDataRows = sheet.getLastRow() - dataStartRow + 1;

    if (numDataRows <= 0) return;

    const existingRules = sheet.getConditionalFormatRules(); // Get existing rules to preserve them
    const newRules = [];

    if (notesIndex >= 0) {
        const notesRange = sheet.getRange(dataStartRow, notesIndex + 1, numDataRows, 1);
        newRules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('QUALITY CHECK')
            .setBackground('#fff2cc') // Light yellow for notes containing "QUALITY CHECK"
            .setRanges([notesRange])
            .build());
    }
    if (statusIndex >= 0) {
        const statusRange = sheet.getRange(dataStartRow, statusIndex + 1, numDataRows, 1);
        newRules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('Quality Review Required')
            .setBackground('#f4cccc') // Light red for status
            .setRanges([statusRange])
            .build());
    }
    
    // It's better to add to existing rules if possible, or manage sets of rules carefully.
    // For simplicity here, we are adding. Be mindful of rule limits if many are applied.
    sheet.setConditionalFormatRules(existingRules.concat(newRules));
    Logger.log("Applied/updated quality validation formatting.");
}


function saveMatches() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Save Matches to History', 
    'Save current matches from "SKU Matching Engine" (with confidence 70%+) to the "Match History" tab? This will append new, unique matches. Continue?', 
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    const historySheet = SS.getSheetByName('Match History');
    if (!matchingSheet || !historySheet) { 
        logError('saveMatches', 'Required sheets (SKU Matching Engine or Match History) not found.', '', '', true, ui);
        return;
    }

    const matchingData = matchingSheet.getDataRange().getValues();
    if (matchingData.length < 3) { // Header, status, then data
        ui.alert('Info', 'No match data found in SKU Matching Engine to save.', ui.ButtonSet.OK);
        return;
    }
    const matchHeaders = matchingData[0];
    const mfrSkuIndex = matchHeaders.indexOf('Manufacturer SKU');
    const statusIndexMatch = matchHeaders.indexOf('Status'); // From SKU Matching Engine
    if (mfrSkuIndex === -1) {
        logError('saveMatches', 'MFR SKU column not found in Matching Engine.', '', '', true, ui);
        return;
    }

    // Map platform column indices from SKU Matching Engine
    const platformMeta = {}; // { PLATFORM: {skuIndex: index, confidenceIndex: index}, ... }
     matchHeaders.forEach((header, index) => {
      const parts = header.split(' ');
      if (parts.length > 1) {
        const platform = parts[0].toUpperCase(); // Ensure uppercase key
        if (PLATFORM_TABS[platform]) { // Check if it's a known platform
          if (!platformMeta[platform]) platformMeta[platform] = {};
          if (parts.slice(1).join(' ') === 'SKU') platformMeta[platform].skuIndex = index;
          else if (parts.slice(1).join(' ') === 'Confidence') platformMeta[platform].confidenceIndex = index;
        }
      }
    });

    const historyHeaders = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
    const histMfrSkuIdx = historyHeaders.indexOf('Manufacturer SKU');
    const histPlatformIdx = historyHeaders.indexOf('Platform');
    const histPlatSkuIdx = historyHeaders.indexOf('Platform SKU');
    const histMatchTypeIdx = historyHeaders.indexOf('Match Type');
    const histConfidenceIdx = historyHeaders.indexOf('Confidence');
    const histDateIdx = historyHeaders.indexOf('Date Added');
    const histUserNotesIdx = historyHeaders.indexOf('User Notes'); // For any future manual notes in history

    if (histMfrSkuIdx === -1 || histPlatformIdx === -1 || histPlatSkuIdx === -1) {
        logError('saveMatches', 'Essential columns missing in Match History sheet.', '', '', true, ui);
        return;
    }
    
    // Load existing history to prevent duplicates
    const existingHistory = {}; // Key: "MFRSKU_PLATFORM_PLATFORMSKU"
    const historyData = historySheet.getDataRange().getValues();
    for (let i = 1; i < historyData.length; i++) { // Skip header
        const key = `${historyData[i][histMfrSkuIdx]}_${historyData[i][histPlatformIdx]}_${historyData[i][histPlatSkuIdx]}`;
        if (historyData[i][histMfrSkuIdx] && historyData[i][histPlatformIdx] && historyData[i][histPlatSkuIdx]) { //Ensure key parts are valid
             existingHistory[key] = true;
        }
    }

    const rowsToAdd = [];
    const currentDate = new Date();

    for (let i = 2; i < matchingData.length; i++) { // Start from data rows in Matching Engine
      const row = matchingData[i];
      const mfrSku = row[mfrSkuIndex];
      if (!mfrSku) continue;
      const statusOnMatchSheet = statusIndexMatch >=0 ? String(row[statusIndexMatch]) : '';

      for (const platformKey in platformMeta) {
        const meta = platformMeta[platformKey];
        if (meta.skuIndex === undefined || meta.confidenceIndex === undefined) continue; // Ensure indices exist

        const platformSku = row[meta.skuIndex];
        const confidenceRaw = row[meta.confidenceIndex];
        if (!platformSku || confidenceRaw === '') continue; // Skip if no platform SKU or confidence

        const confidence = parseFloat(confidenceRaw);
        if (isNaN(confidence) || confidence < 70) continue; // Save only if confidence is 70+

        const historyKey = `${mfrSku}_${platformKey}_${platformSku}`;
        if (existingHistory[historyKey]) continue; // Skip if already in history

        let matchType = 'Auto'; // Default
        if (statusOnMatchSheet.toLowerCase().includes('manual')) matchType = 'Manual'; // If status indicates manual intervention
        else if (confidence >= 95) matchType = 'Exact/High';
        else if (confidence >= 85) matchType = 'High Confidence';
        else if (confidence >= 70) matchType = 'Medium Confidence';
        
        const historyRow = Array(historyHeaders.length).fill('');
        historyRow[histMfrSkuIdx] = mfrSku;
        historyRow[histPlatformIdx] = platformKey;
        historyRow[histPlatSkuIdx] = platformSku;
        historyRow[histMatchTypeIdx] = matchType;
        historyRow[histConfidenceIdx] = confidence.toFixed(2);
        historyRow[histDateIdx] = currentDate;
        if (histUserNotesIdx !== -1) historyRow[histUserNotesIdx] = ''; // Placeholder for notes

        rowsToAdd.push(historyRow);
        existingHistory[historyKey] = true; // Add to current batch's existing history to prevent dupes within this run
      }
    }

    if (rowsToAdd.length > 0) {
      historySheet.getRange(historySheet.getLastRow() + 1, 1, rowsToAdd.length, historyHeaders.length).setValues(rowsToAdd);
      ui.alert('Success', `${rowsToAdd.length} new unique matches (70%+) saved to Match History.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Info', 'No new unique matches (70%+) found to save to Match History.', ui.ButtonSet.OK);
    }

  } catch (error) {
      logError('saveMatches', 'Error saving matches', '', error, true, ui);
  }
}

function getDashboardIndices(headers) {
  return {
    sku: headers.indexOf('SKU'),
    map: headers.indexOf('MAP'),
    dealer: headers.indexOf('Dealer'),
    bStock: headers.indexOf('B-Stock'),
    platform: headers.indexOf('Platform'),
    current: headers.indexOf('Current'),
    newPrice: headers.indexOf('New'),
    changeAmt: headers.indexOf('Change $'),
    changePct: headers.indexOf('Change %'),
    status: headers.indexOf('Status')
  };
}

function updateDashboardSummary(sheet, metrics) {
  sheet.getRange(3, 2).setValue(metrics.totalProducts);
  sheet.getRange(3, 5).setValue(metrics.priceIncreases);
  sheet.getRange(3, 8).setValue(metrics.priceDecreases);
  sheet.getRange(3, 11).setValue(metrics.avgChange).setNumberFormat('0.0%');
  sheet.getRange(4, 2).setValue(metrics.mapViolations);
  sheet.getRange(4, 5).setValue(metrics.highImpactChanges);
  sheet.getRange(4, 8).setValue(metrics.unmatched);
  sheet.getRange(4, 11).setValue(metrics.bStockChanges);
}

function updatePriceAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Update Price Analysis', 
    'This will analyze price changes based on "SKU Matching Engine" and "Manufacturer Price Sheet". Results will update the "Price Analysis Dashboard" and "Price Changes" tab. Continue?', 
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let dashboardSheet;
  let priceChangesSheet;
  try {
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    dashboardSheet = SS.getSheetByName('Price Analysis Dashboard');
    priceChangesSheet = SS.getSheetByName('Price Changes'); 
    const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet');

    if (!matchingSheet || !dashboardSheet || !mfrSheet) { 
        logError('updatePriceAnalysis', 'Required sheets (SKU Matching Engine, Price Analysis Dashboard, Manufacturer Price Sheet) not found.', '', '', true, ui);
        return;
    }
    if (!priceChangesSheet) { // Ensure Price Changes tab exists, create if not
        priceChangesSheet = createPriceChangesTab(); 
        if (!priceChangesSheet) {
             logError('updatePriceAnalysis', 'Could not create or find "Price Changes" sheet.', '', '', true, ui); return;
        }
    }

    // Status updates
    dashboardSheet.getRange(1, 1, 1, dashboardSheet.getLastColumn()).merge().setValue('PROCESSING - Analyzing prices...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center');
    priceChangesSheet.getRange("A1").setValue("PROCESSING - Updating Price Change Lists...").setFontWeight("bold").setBackground(COLORS.WARNING);
    SpreadsheetApp.flush();

    // Get data from SKU Matching Engine
    const matchingData = matchingSheet.getDataRange().getValues();
    if (matchingData.length < 3) {
        dashboardSheet.getRange(1,1).setValue('PRICE ANALYSIS DASHBOARD - No data in SKU Matching Engine');
        logError('updatePriceAnalysis', 'No data in SKU Matching Engine.', '', '', true, ui);
        return;
    }
    const matchingHeaders = matchingData[0];
    const mfrSkuIndexMatch = matchingHeaders.indexOf('Manufacturer SKU');
    const mapIndexMatch = matchingHeaders.indexOf('MAP'); // MAP from MFR sheet, carried to matching engine
    const dealerPriceIndexMatch = matchingHeaders.indexOf('Dealer Price'); // Dealer Price from MFR

    // Identify platform-specific columns in SKU Matching Engine
    const platformMeta = {}; // { PLATFORM: {skuIndex: index, priceIndex: index, confidenceIndex: index}, ... }
    matchingHeaders.forEach((header, index) => {
        const parts = header.split(' ');
        if (parts.length > 1) {
            const platform = parts[0].toUpperCase();
            // Only consider platforms defined for analysis (e.g., those with pricing data)
            if (ANALYSIS_PLATFORMS.includes(platform) && PLATFORM_TABS[platform]) {
                if (!platformMeta[platform]) platformMeta[platform] = {};
                if (parts.slice(1).join(' ') === 'SKU') platformMeta[platform].skuIndex = index;
                else if (parts.slice(1).join(' ') === 'Current Price') platformMeta[platform].priceIndex = index;
                else if (parts.slice(1).join(' ') === 'Confidence') platformMeta[platform].confidenceIndex = index;
            }
        }
    });

    resetDashboardLayout(dashboardSheet); // Clear old data and reset layout
    const dashboardHeaders = dashboardSheet.getRange(6, 1, 1, dashboardSheet.getLastColumn()).getValues()[0];
    const dashIndices = getDashboardIndices(dashboardHeaders);

    const analysisDataForDashboard = [];
    const priceUpItemsForPriceChangeTab = [];
    const priceDownItemsForPriceChangeTab = [];

    let totalProductsAnalyzed = 0, priceIncreases = 0, priceDecreases = 0, mapViolations = 0;
    let highImpactChanges = 0, unmatchedProductsThisRun = 0, totalPercentChangeValue = 0;
    let validProductCountForAvg = 0, bStockChanges = 0;

    for (let i = 2; i < matchingData.length; i++) { // Start from data rows in Matching Engine
      const row = matchingData[i];
      const mfrSku = String(row[mfrSkuIndexMatch]);
      if (!mfrSku) continue;
      
      totalProductsAnalyzed++;

      const newMapPrice = parseFloat(row[mapIndexMatch]);
      const newDealerPrice = parseFloat(row[dealerPriceIndexMatch]);
      if (isNaN(newMapPrice) || isNaN(newDealerPrice)) {
          logError('updatePriceAnalysis', `Invalid MAP or Dealer Price for MFR SKU ${mfrSku} in Matching Engine. Row ${i+1}.`);
          continue; // Skip if essential MFR prices are not numbers
      }

      const bStockInfo = getBStockInfo(mfrSku);
      const isBStockProduct = bStockInfo !== null;
      let platformProcessedForThisMfrSku = false; // To count unmatched products correctly

      for (const platformKey in platformMeta) {
        const meta = platformMeta[platformKey];
        if (!meta.skuIndex || !meta.priceIndex || !meta.confidenceIndex) continue; // Ensure all needed columns are mapped

        const platformSku = row[meta.skuIndex];
        const currentPlatformPriceRaw = row[meta.priceIndex];
        const confidenceRaw = row[meta.confidenceIndex];
        
        const confidence = parseFloat(confidenceRaw);
        if (!platformSku || isNaN(confidence) || confidence < 85) { // Only analyze high-confidence matches (85%+)
            // Handled below for unmatchedProductsThisRun check
            continue; 
        }
        platformProcessedForThisMfrSku = true; // At least one platform has a high-confidence match

        const currentPlatformPrice = !isNaN(parseFloat(currentPlatformPriceRaw)) ? parseFloat(currentPlatformPriceRaw) : null;
        if (currentPlatformPrice === null && currentPlatformPriceRaw !== '') { // Log if price was not empty but not a number
            logError('updatePriceAnalysis', `Invalid current platform price '${currentPlatformPriceRaw}' for ${platformKey} SKU ${platformSku}.`);
        }


        let calculatedNewPrice;
        let priceStatus = '';

        if (isBStockProduct) {
          calculatedNewPrice = newMapPrice * bStockInfo.multiplier; // Use B-Stock multiplier on MAP
          priceStatus = `${bStockInfo.type} ${bStockInfo.isSpecial ? "Special" : "B-Stock"}`;
        } else {
          calculatedNewPrice = newMapPrice; // For new items, new price is MAP
        }
        calculatedNewPrice = parseFloat(calculatedNewPrice.toFixed(2)); // Round to 2 decimal places

        let changeAmount = 0;
        let changePercent = 0;

        if (currentPlatformPrice !== null && !isNaN(calculatedNewPrice)) {
          changeAmount = parseFloat((calculatedNewPrice - currentPlatformPrice).toFixed(2));
          changePercent = currentPlatformPrice !== 0 ? (changeAmount / currentPlatformPrice) * 100 : (calculatedNewPrice > 0 ? Infinity : 0);

          if (changeAmount > 0.001) { // Price Increase
            priceIncreases++;
            priceUpItemsForPriceChangeTab.push([mfrSku, platformKey, currentPlatformPrice, calculatedNewPrice, changeAmount, isFinite(changePercent) ? changePercent / 100 : 'N/A']);
          } else if (changeAmount < -0.001) { // Price Decrease
            priceDecreases++;
            priceDownItemsForPriceChangeTab.push([mfrSku, platformKey, currentPlatformPrice, calculatedNewPrice, changeAmount, isFinite(changePercent) ? changePercent / 100 : 'N/A']);
          }

          // High impact and MAP violation checks
          if (Math.abs(changePercent) > 10 || Math.abs(changeAmount) > 10) { highImpactChanges++; priceStatus += (priceStatus ? '; ' : '') + 'High Impact'; }
          if (!isBStockProduct && newMapPrice > 0 && currentPlatformPrice < (newMapPrice - 0.001) ) { mapViolations++; priceStatus += (priceStatus ? '; ' : '') + 'MAP Violation'; }
          if (isBStockProduct && Math.abs(changeAmount) > 0.01) bStockChanges++;
          if (isFinite(changePercent)) { totalPercentChangeValue += changePercent; validProductCountForAvg++; }

        } else if (currentPlatformPrice === null && !isNaN(calculatedNewPrice)){ // No current price, but new price calculated (new listing)
            priceStatus += (priceStatus ? '; ' : '') + 'New Listing/Price';
            priceIncreases++; // Treat as an increase for tracking purposes
            priceUpItemsForPriceChangeTab.push([mfrSku, platformKey, 'N/A', calculatedNewPrice, calculatedNewPrice, 'New Item']);
        } else if (currentPlatformPrice !== null && isNaN(calculatedNewPrice)) {
             logError('updatePriceAnalysis', `Could not calculate new price for ${platformKey} MFR SKU ${mfrSku}. Current: ${currentPlatformPrice}`);
             priceStatus += (priceStatus ? '; ' : '') + 'New Price Calc Error';
        }


        // Prepare row for Price Analysis Dashboard
        const analysisRow = Array(dashboardHeaders.length).fill('');
        if (dashIndices.sku !== -1) analysisRow[dashIndices.sku] = mfrSku;
        if (dashIndices.map !== -1) analysisRow[dashIndices.map] = newMapPrice;
        if (dashIndices.dealer !== -1) analysisRow[dashIndices.dealer] = newDealerPrice;
        if (dashIndices.bStock !== -1) analysisRow[dashIndices.bStock] = isBStockProduct ? bStockInfo.type : '';
        if (dashIndices.platform !== -1) analysisRow[dashIndices.platform] = platformKey;
        if (dashIndices.current !== -1) analysisRow[dashIndices.current] = currentPlatformPrice !== null ? currentPlatformPrice : '';
        if (dashIndices.newPrice !== -1) analysisRow[dashIndices.newPrice] = !isNaN(calculatedNewPrice) ? calculatedNewPrice : '';
        if (dashIndices.changeAmt !== -1) analysisRow[dashIndices.changeAmt] = changeAmount;
        if (dashIndices.changePct !== -1) analysisRow[dashIndices.changePct] = isFinite(changePercent) ? changePercent / 100 : (changePercent === Infinity ? 'New Item' : '');
        if (dashIndices.status !== -1) analysisRow[dashIndices.status] = priceStatus;
        analysisDataForDashboard.push(analysisRow);
      }
      if (!platformProcessedForThisMfrSku && totalProductsAnalyzed > 0) { // Checked after all platforms for this MFR SKU
          unmatchedProductsThisRun++;
      }
    }

    // Update Price Analysis Dashboard
    if (analysisDataForDashboard.length > 0) {
      const dataRange = dashboardSheet.getRange(7, 1, analysisDataForDashboard.length, dashboardHeaders.length);
      dataRange.setValues(analysisDataForDashboard);
      // Apply number formats to currency and percentage columns
      const formatColumns = [ 
          { header: 'MAP', format: '$#,##0.00' }, { header: 'Dealer', format: '$#,##0.00' }, 
          { header: 'Current', format: '$#,##0.00' }, { header: 'New', format: '$#,##0.00' }, 
          { header: 'Change $', format: '$#,##0.00' }, { header: 'Change %', format: '0.00%' }
      ];
      formatColumns.forEach(colInfo => { 
          const colIndex = dashboardHeaders.indexOf(colInfo.header); 
          if (colIndex !== -1) {
              dashboardSheet.getRange(7, colIndex + 1, analysisDataForDashboard.length, 1).setNumberFormat(colInfo.format);
          }
      });
    }
    const metrics = {
        totalProducts: totalProductsAnalyzed,
        priceIncreases,
        priceDecreases,
        avgChange: validProductCountForAvg > 0 ? (totalPercentChangeValue / validProductCountForAvg / 100) : 0,
        mapViolations,
        highImpactChanges,
        unmatched: unmatchedProductsThisRun,
        bStockChanges
    };
    updateDashboardSummary(dashboardSheet, metrics);
    applyDashboardFormatting(dashboardSheet, analysisDataForDashboard.length);
    dashboardSheet.getRange(1, 1, 1, dashboardSheet.getLastColumn()).merge().setValue('PRICE ANALYSIS DASHBOARD - COMPLETE - Last Updated: ' + new Date().toLocaleString()).setBackground('#e0e0e0');

    // Populate Price Changes Tab
    populatePriceChangesTab(priceChangesSheet, priceUpItemsForPriceChangeTab, priceDownItemsForPriceChangeTab);

    ui.alert('Analysis Complete', 
             `Price analysis updated.\nProducts analyzed from Matching Engine: ${totalProductsAnalyzed}.\nPrice Increases: ${priceIncreases}, Price Decreases: ${priceDecreases}.\nCheck "Price Analysis Dashboard" and "Price Changes" tabs.`, 
             ui.ButtonSet.OK);

  } catch (error) {
    logError('updatePriceAnalysis', 'Error during price analysis', '', error, true, ui);
    if (dashboardSheet) dashboardSheet.getRange(1, 1, 1, dashboardSheet.getLastColumn()).merge().setValue('ERROR - ' + error.toString()).setBackground('#f4cccc');
    if (priceChangesSheet) priceChangesSheet.getRange("A1").setValue("ERROR during price analysis. Check logs.").setBackground(COLORS.NEGATIVE);
  }
}

function populatePriceChangesTab(sheet, priceUpItems, priceDownItems) {
  if (!sheet) {
      logError('populatePriceChangesTab', 'Price Changes sheet not found.');
      return;
  }
  sheet.clearContents(); // Clear previous content

  const headers = ["MFR SKU", "Platform", "Old Price", "New Price", "Change $", "Change %"];
  const currencyFormat = "$#,##0.00";
  const percentFormat = "0.00%";

  // Price Increases
  sheet.getRange("A1").setValue("PRICE INCREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_UP);
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f2f2f2");
  
  let currentRow = 3;
  if (priceUpItems.length > 0) {
    sheet.getRange(currentRow, 1, priceUpItems.length, headers.length).setValues(priceUpItems);
    // Format columns 3, 4, 5 (Old Price, New Price, Change $) as currency
    sheet.getRange(currentRow, 3, priceUpItems.length, 3).setNumberFormat(currencyFormat);
    // Format column 6 (Change %) as percentage
    sheet.getRange(currentRow, 6, priceUpItems.length, 1).setNumberFormat(percentFormat);
    currentRow += priceUpItems.length;
  } else {
    sheet.getRange(currentRow, 1).setValue("No high-confidence price increases found.").setFontStyle("italic");
    currentRow++;
  }

  currentRow += 2; // Add some space before decreases section

  // Price Decreases
  sheet.getRange(currentRow, 1).setValue("PRICE DECREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_DOWN);
  currentRow++;
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f2f2f2");
  currentRow++;

  if (priceDownItems.length > 0) {
    sheet.getRange(currentRow, 1, priceDownItems.length, headers.length).setValues(priceDownItems);
    // Format columns 3, 4, 5 (Old Price, New Price, Change $) as currency
    sheet.getRange(currentRow, 3, priceDownItems.length, 3).setNumberFormat(currencyFormat);
    // Format column 6 (Change %) as percentage
    sheet.getRange(currentRow, 6, priceDownItems.length, 1).setNumberFormat(percentFormat);
  } else {
    sheet.getRange(currentRow, 1).setValue("No high-confidence price decreases found.").setFontStyle("italic");
  }
  
  // Auto-resize columns for better readability
  for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
  }
  Logger.log("Price Changes tab populated.");
}


function resetDashboardLayout(sheet) { 
  if (!sheet) return;
  // Clear data area (from row 7 down)
  if (sheet.getLastRow() >= 7) {
    sheet.getRange(7, 1, sheet.getLastRow() - 6, sheet.getMaxColumns())
         .clear({contentsOnly: true, formatOnly: true, validationsOnly: true, commentsOnly: true, conditionalFormatRulesOnly: true});
  }
  // Clear summary metrics area (rows 2-4 for content, keep row 1 title)
  sheet.getRange(2, 1, 3, sheet.getMaxColumns()).clear({contentsOnly: true, formatOnly: true}); 
  sheet.clearConditionalFormatRules(); // Clear all conditional formats on the sheet

  // Re-establish fixed layout elements
  sheet.getRange(2, 1, 1, 12).merge().setValue('SUMMARY METRICS')
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f3f3f3');
  const summaryMetrics = [
      ['Total Products:', '0', '', 'Price Increases:', '0', '', 'Price Decreases:', '0', '', 'Average Change:', '0%'],
      ['MAP Violations:', '0', '', 'High Impact Changes:', '0', '', 'Unmatched Products:', '0', '', 'B-Stock Changes:', '0']
  ];
  sheet.getRange(3, 1, 2, 11).setValues(summaryMetrics); 
  sheet.getRange(3,1,2,1).setFontWeight('bold'); sheet.getRange(3,4,2,1).setFontWeight('bold');
  sheet.getRange(3,7,2,1).setFontWeight('bold'); sheet.getRange(3,10,2,1).setFontWeight('bold');

  const analysisHeaders = ['SKU', 'MAP', 'Dealer', 'B-Stock', 'Platform', 'Current', 'New', 'Change $', 'Change %', 'Status'];
  sheet.getRange(6, 1, 1, analysisHeaders.length).setValues([analysisHeaders])
      .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(5,1,1,analysisHeaders.length).merge().setValue('Use "Update Price Analysis" to refresh data below.')
      .setFontStyle('italic').setHorizontalAlignment('center').setBackground(null); // Clear background for this message
}

function applyDashboardFormatting(sheet, rowCount) {
    if (!sheet || rowCount <= 0) return;
    sheet.clearConditionalFormatRules(); // Start fresh for dashboard rules
    const headers = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dataStartRow = 7;

    const changeAmtColIndex = headers.indexOf('Change $') + 1;
    const changePctColIndex = headers.indexOf('Change %') + 1;
    const statusColIndex = headers.indexOf('Status') + 1;
    const rules = [];

    // Formatting for 'Change $' column (Price increases are negative for profit, decreases positive)
    if (changeAmtColIndex > 0) {
        const range = sheet.getRange(dataStartRow, changeAmtColIndex, rowCount, 1);
        // Price went UP (calculatedNew > currentPlatformPrice, so changeAmount is POSITIVE) - potentially good
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(COLORS.POSITIVE).setRanges([range]).build()); // Green text for price up
        // Price went DOWN (calculatedNew < currentPlatformPrice, so changeAmount is NEGATIVE) - potentially bad
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(COLORS.NEGATIVE).setRanges([range]).build());   // Red text for price down
    }

    // Formatting for 'Change %' column for significant changes
    if (changePctColIndex > 0) {
        const range = sheet.getRange(dataStartRow, changePctColIndex, rowCount, 1);
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.10).setBackground('#fce5cd').setRanges([range]).build()); // Large positive % change (Orange BG)
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.05, 0.10).setBackground('#fff2cc').setRanges([range]).build()); // Medium positive % change (Yellow BG)
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.10).setBackground('#d9ead3').setRanges([range]).build()); // Large negative % change (Green BG)
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.10, -0.05).setBackground('#e2f0d9').setRanges([range]).build());// Medium negative % change (Light Green BG)
    }

    // Formatting for 'Status' column
    if (statusColIndex > 0) {
        const range = sheet.getRange(dataStartRow, statusColIndex, rowCount, 1);
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('High Impact').setBackground(COLORS.WARNING).setRanges([range]).build());         // Yellow BG
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('MAP Violation').setBackground(COLORS.NEGATIVE).setFontColor('#FFFFFF').setRanges([range]).build());// Red BG, White text
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('B-Stock').setBackground(COLORS.INFLOW).setFontColor('#FFFFFF').setRanges([range]).build());       // Blue BG, White text
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Special').setBackground(COLORS.SHOPIFY).setFontColor('#FFFFFF').setRanges([range]).build());     // Green BG, White text (Shopify green)
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('New Listing/Price').setBackground('#cfe2f3').setRanges([range]).build()); // Light Blue BG
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Error').setBackground('#780000').setFontColor('#FFFFFF').setRanges([range]).build());// Dark Red BG
    }

    if (rules.length > 0) {
        sheet.setConditionalFormatRules(rules);
        Logger.log("Applied conditional formatting to Price Analysis Dashboard.");
    }
}

function generateExports() {
  const ui = SpreadsheetApp.getUi();
  const EXPORT_CONFIDENCE_THRESHOLD = 85; // Only export matches with this confidence or higher
  const response = ui.alert('Generate Export Files', 
    `Generate export files for each platform? This uses prices from the "Price Analysis Dashboard" for matches with ${EXPORT_CONFIDENCE_THRESHOLD}%+ confidence. Continue?`, 
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let dashboardSheet;
  try {
    dashboardSheet = SS.getSheetByName('Price Analysis Dashboard');
    const matchingSheet = SS.getSheetByName('SKU Matching Engine'); // Need this for original platform SKUs
    const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet'); // Need for UPCs if available

    if (!matchingSheet || !mfrSheet || !dashboardSheet) {
      logError('generateExports', 'Required sheets (Dashboard, Matching Engine, MFR Sheet) not found.', '', '', true, ui);
      return;
    }

    const titleCell = dashboardSheet.getRange(1, 1); // Check if analysis has been run
    const titleText = titleCell.getValue().toString();
    if (!titleText.includes('COMPLETE')) {
      ui.alert('Price Analysis Required', 'Please run "Update Price Analysis" first to ensure prices are current before generating exports.', ui.ButtonSet.OK);
      return;
    }

    // --- Data from Price Analysis Dashboard ---
    const analysisData = dashboardSheet.getRange(7, 1, Math.max(1, dashboardSheet.getLastRow()-6), dashboardSheet.getLastColumn()).getValues();
    const analysisHeaders = dashboardSheet.getRange(6, 1, 1, dashboardSheet.getLastColumn()).getValues()[0];
    // Indices from Analysis Dashboard
    const anSkuIdx = analysisHeaders.indexOf('SKU');
    const anPlatformIdx = analysisHeaders.indexOf('Platform');
    const anNewPriceIdx = analysisHeaders.indexOf('New'); // This is the calculated new price
    const anMapIdx = analysisHeaders.indexOf('MAP');
    const anDealerIdx = analysisHeaders.indexOf('Dealer');
    const anBstockTypeIdx = analysisHeaders.indexOf('B-Stock'); // B-Stock type, e.g., "BA"

    // --- Data from SKU Matching Engine (for original platform SKU and confidence) ---
    const matchingData = matchingSheet.getDataRange().getValues();
    const matchingHeaders = matchingData[0];
    const matchMfrSkuIdx = matchingHeaders.indexOf('Manufacturer SKU');
    // Create a map for quick lookup: MFR_SKU -> {platform_sku_col_idx, confidence_col_idx} for each platform
    const platformColMapOnMatchingSheet = {};
    matchingHeaders.forEach((header, index) => {
        const parts = header.split(' ');
        if (parts.length > 1) {
            const platform = parts[0].toUpperCase();
            if (PLATFORM_TABS[platform]) {
                if (!platformColMapOnMatchingSheet[platform]) platformColMapOnMatchingSheet[platform] = {};
                if (parts.slice(1).join(' ') === 'SKU') platformColMapOnMatchingSheet[platform].skuIndex = index;
                if (parts.slice(1).join(' ') === 'Confidence') platformColMapOnMatchingSheet[platform].confidenceIndex = index;
            }
        }
    });
     // Map MFR SKU to its row index in matchingData for quick lookup
    const mfrSkuToMatchRowIndex = {};
    for(let i=2; i < matchingData.length; i++) { // Data starts at row 3 (index 2)
        if(matchingData[i][matchMfrSkuIdx]) {
            mfrSkuToMatchRowIndex[String(matchingData[i][matchMfrSkuIdx])] = i;
        }
    }


    // --- Data from Manufacturer Price Sheet (for UPC) ---
    const mfrPriceSheetData = mfrSheet.getDataRange().getValues();
    const mfrPriceSheetHeaders = mfrPriceSheetData[0];
    const mfrSkuIndexUPC = mfrPriceSheetHeaders.indexOf('Manufacturer SKU');
    const upcIndexUPC = mfrPriceSheetHeaders.indexOf('UPC');
    const upcMap = {}; // MFR_SKU -> UPC
    if (mfrSkuIndexUPC !== -1 && upcIndexUPC !== -1) {
      for (let i = 1; i < mfrPriceSheetData.length; i++) {
        if (mfrPriceSheetData[i][mfrSkuIndexUPC]) {
          upcMap[String(mfrPriceSheetData[i][mfrSkuIndexUPC])] = String(mfrPriceSheetData[i][upcIndexUPC] || '');
        }
      }
    }

    // --- Generate Exports for each Platform Tab ---
    for (const platformKey in PLATFORM_TABS) {
      const exportSheetName = PLATFORM_TABS[platformKey].name;
      let exportSheet = SS.getSheetByName(exportSheetName);
      if (!exportSheet) {
        createExportTab(platformKey); // Create if doesn't exist
        exportSheet = SS.getSheetByName(exportSheetName);
        if(!exportSheet) {
            logError('generateExports', `Error: Could not create/find export sheet for ${platformKey}`);
            continue; // Skip this platform
        }
      }
      const exportHeaders = PLATFORM_TABS[platformKey].headers;
      // Clear old data from export sheet (from row 4 down)
      if (exportSheet.getLastRow() > 3) {
        exportSheet.getRange(4, 1, exportSheet.getLastRow() - 3, exportHeaders.length).clearContent();
      }

      const exportRows = [];
      for (const analysisRow of analysisData) {
        const mfrSku = String(analysisRow[anSkuIdx]);
        const platformAnalyzed = String(analysisRow[anPlatformIdx]).toUpperCase(); // Platform from analysis dashboard

        if (!mfrSku || platformAnalyzed !== platformKey) continue; // Only process items for the current platformKey

        const matchRowIndex = mfrSkuToMatchRowIndex[mfrSku];
        if (matchRowIndex === undefined) {
            // Logger.log(`MFR SKU ${mfrSku} not found in SKU Matching Engine map during export generation for ${platformKey}.`);
            continue;
        }
        
        // Get platform SKU and confidence from Matching Engine sheet
        const platformSpecificMetaOnMatching = platformColMapOnMatchingSheet[platformKey];
        if (!platformSpecificMetaOnMatching || platformSpecificMetaOnMatching.skuIndex === undefined || platformSpecificMetaOnMatching.confidenceIndex === undefined) {
            // Logger.log(`Column mapping for ${platformKey} SKU or Confidence not found on Matching Engine sheet.`);
            continue;
        }

        const platformSkuOnMatch = String(matchingData[matchRowIndex][platformSpecificMetaOnMatching.skuIndex]);
        const confidence = parseFloat(matchingData[matchRowIndex][platformSpecificMetaOnMatching.confidenceIndex]);

        if (!platformSkuOnMatch || isNaN(confidence) || confidence < EXPORT_CONFIDENCE_THRESHOLD) continue; // Skip if no SKU, or low confidence

        // Get other necessary data from analysisRow
        const newPrice = parseFloat(analysisRow[anNewPriceIdx]);
        if (isNaN(newPrice)) {
            logError('generateExports', `New price is not a number for MFR SKU ${mfrSku}, Platform ${platformKey}. Value: ${analysisRow[anNewPriceIdx]}`);
            continue;
        }
        const mapPrice = parseFloat(analysisRow[anMapIdx]); // Already ensured valid in analysis
        const dealerPrice = parseFloat(analysisRow[anDealerIdx]); // Already ensured valid in analysis
        const bStockTypeFromAnalysis = String(analysisRow[anBstockTypeIdx]);
        const isBAsku = bStockTypeFromAnalysis && bStockTypeFromAnalysis !== ''; // Is it a B-Stock item?
        const bStockInfoForExport = isBAsku ? getBStockInfo(mfrSku) : null; // Get full B-Stock info if needed for conditions

        const exportRowData = createPlatformExportRow(
          platformKey, exportHeaders, platformSkuOnMatch, newPrice,
          0, // MSRP (not directly used from analysis, can be fetched if needed)
          mapPrice, dealerPrice, upcMap[mfrSku] || '', 
          isBAsku, bStockInfoForExport
        );
        if (exportRowData) exportRows.push(exportRowData);
      }

      // Write to export sheet
      if (exportRows.length > 0) {
        exportSheet.getRange(4, 1, exportRows.length, exportHeaders.length).setValues(exportRows);
        exportSheet.getRange(2, 1, 1, exportHeaders.length).merge()
                   .setValue(`READY FOR EXPORT - ${exportRows.length} items (${EXPORT_CONFIDENCE_THRESHOLD}%+) - ${new Date().toLocaleString()}`)
                   .setBackground('#d9ead3').setFontColor('#000000');
      } else {
        exportSheet.getRange(2, 1, 1, exportHeaders.length).merge()
                   .setValue(`NO DATA (${EXPORT_CONFIDENCE_THRESHOLD}%+ matches not found for ${platformKey})`)
                   .setBackground('#f4cccc').setFontColor('#000000');
      }
    }
    ui.alert('Exports Generated', `Export files updated with ${EXPORT_CONFIDENCE_THRESHOLD}%+ confidence matches using prices from the Analysis Dashboard.`, ui.ButtonSet.OK);
  } catch (error) {
    logError('generateExports', 'Error generating export files', '', error, true, ui);
  }
}

function createPlatformExportRow(platform, headers, platformSku, newPrice, msrp, map, dealerPrice, upc, isBAsku, bStockInfo) {
  const row = Array(headers.length).fill('');
  const formattedPrice = !isNaN(newPrice) ? parseFloat(newPrice.toFixed(2)) : '';
  const formattedMap = !isNaN(map) ? parseFloat(map.toFixed(2)) : '';
  const formattedDealer = !isNaN(dealerPrice) ? parseFloat(dealerPrice.toFixed(2)) : '';

  const idx = (name) => headers.indexOf(name); // Helper to find header index

  try {
      switch (platform) {
        case 'AMAZON':
          if (idx('seller-sku') !== -1) row[idx('seller-sku')] = platformSku;
          if (idx('price') !== -1) row[idx('price')] = formattedPrice;
          break;
        case 'EBAY':
          if (idx('Action') !== -1) row[idx('Action')] = 'Revise'; // Assuming revise for existing matched items
          if (idx('Item number') !== -1) row[idx('Item number')] = ''; // Item number usually needed for revise, may need lookup
          if (idx('Custom label (SKU)') !== -1) row[idx('Custom label (SKU)')] = platformSku;
          if (idx('Start price') !== -1) row[idx('Start price')] = formattedPrice;
          break;
        case 'SHOPIFY':
          if (idx('Variant SKU') !== -1) row[idx('Variant SKU')] = platformSku;
          if (idx('Variant Price') !== -1) row[idx('Variant Price')] = formattedPrice;
          // Compare At Price for B-Stock could be the original MAP of the new item
          if (idx('Variant Compare At Price') !== -1) row[idx('Variant Compare At Price')] = isBAsku ? formattedMap : ''; // Show MAP as compare for B-stock
          if (idx('Variant Cost') !== -1) row[idx('Variant Cost')] = formattedDealer;
          break;
        case 'INFLOW':
          if (idx('Name') !== -1) row[idx('Name')] = platformSku; // Assuming Name is SKU for InFlow
          if (idx('UnitPrice') !== -1) row[idx('UnitPrice')] = formattedPrice;
          if (idx('Cost') !== -1) row[idx('Cost')] = formattedDealer;
          break;
        case 'SELLERCLOUD':
          if (idx('ProductID') !== -1) row[idx('ProductID')] = platformSku;
          if (idx('MAPPrice') !== -1) row[idx('MAPPrice')] = formattedMap;
          if (idx('SitePrice') !== -1) row[idx('SitePrice')] = formattedPrice;
          if (idx('SiteCost') !== -1) row[idx('SiteCost')] = formattedDealer;
          break;
        case 'REVERB':
          if (idx('sku') !== -1) row[idx('sku')] = platformSku;
          if (idx('price') !== -1) row[idx('price')] = formattedPrice;
          if (idx('condition') !== -1) {
            if (isBAsku && bStockInfo) {
              // Map B-Stock type to Reverb conditions
              switch (bStockInfo.type) {
                case 'AA': row[idx('condition')] = 'Excellent'; break; // Or Mint if appropriate
                case 'BA': case 'BB': row[idx('condition')] = 'Very Good'; break;
                case 'BC': case 'BD': case 'NOACC': row[idx('condition')] = 'Good'; break;
                default: row[idx('condition')] = 'Good'; // Default B-Stock condition
              }
            } else {
              row[idx('condition')] = 'Brand New'; // For non-B-Stock items
            }
          }
          break;
        default:
          Logger.log("Unknown platform for export row creation: " + platform);
          return null; // Skip unknown platforms
      }
  } catch(e) {
      logError('createPlatformExportRow', `Error creating export row for ${platform}, SKU ${platformSku}`, platformSku, e);
      return null;
  }
  return row;
}

function generateCPListings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Generate CP Listings', 
    'Identify items from "Manufacturer Price Sheet" that are not listed (or have low confidence matches <85%) on some platforms. Results go to "CP Listings" tab. Continue?', 
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let cpSheet;
  try {
    const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet');
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    cpSheet = SS.getSheetByName('CP Listings');

    if (!mfrSheet || !matchingSheet || !cpSheet) {
      logError('generateCPListings', 'Required sheets (MFR, Matching Engine, CP Listings) not found.', '', '', true, ui);
      return;
    }
    
    // Clear CP Listings sheet (from row 3 down)
    if (cpSheet.getLastRow() > 2) {
        cpSheet.getRange(3, 1, cpSheet.getLastRow() - 2, cpSheet.getLastColumn()).clearContent();
    }

    cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge()
           .setValue('PROCESSING - Generating CP listings...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center');
    SpreadsheetApp.flush();

    // Get MFR Data
    const mfrDataValues = mfrSheet.getDataRange().getValues();
    const mfrHeaders = mfrDataValues[0].map(h => String(h).trim());
    const mfrSkuCol = mfrHeaders.indexOf('Manufacturer SKU');
    const upcCol = mfrHeaders.indexOf('UPC');
    const msrpCol = mfrHeaders.indexOf('MSRP');
    const mapCol = mfrHeaders.indexOf('MAP');
    const dealerPriceCol = mfrHeaders.indexOf('Dealer Price');
    if (mfrSkuCol === -1) {
        throw new Error('MFR SKU column not found in Manufacturer Price Sheet.');
    }

    // Get Matching Engine Data to check platform presence
    const matchingDataValues = matchingSheet.getDataRange().getValues();
    const matchingHeaders = matchingDataValues[0].map(h => String(h).trim());
    const matchingMfrSkuCol = matchingHeaders.indexOf('Manufacturer SKU');
    
    const platformPresence = {}; // mfrSku -> { platformKey: true/false }
    const platformColsInMatching = {}; // platformKey -> { skuColIndex, confidenceColIndex }
    Object.keys(PLATFORM_TABS).forEach(pKey => {
      platformColsInMatching[pKey] = {
        sku: matchingHeaders.indexOf(pKey + ' SKU'),
        confidence: matchingHeaders.indexOf(pKey + ' Confidence')
      };
    });

    for (let i = 2; i < matchingDataValues.length; i++) { // Start from data row
      const row = matchingDataValues[i];
      const mfrSku = String(row[matchingMfrSkuCol]);
      if (!mfrSku) continue;
      if (!platformPresence[mfrSku]) platformPresence[mfrSku] = {};
      
      Object.keys(PLATFORM_TABS).forEach(pKey => {
        const skuIdx = platformColsInMatching[pKey].sku;
        const confIdx = platformColsInMatching[pKey].confidence;
        if (skuIdx !== -1 && row[skuIdx] && confIdx !== -1) {
            const confidence = parseFloat(row[confIdx]);
            platformPresence[mfrSku][pKey] = (!isNaN(confidence) && confidence >= 85);
        } else {
            platformPresence[mfrSku][pKey] = false;
        }
      });
    }

    const cpDataRows = [];
    const cpHeaders = cpSheet.getRange(1, 1, 1, cpSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const cpHeaderMap = {}; // Header name to column index
    cpHeaders.forEach((h,i) => cpHeaderMap[h] = i);

    for (let i = 1; i < mfrDataValues.length; i++) { // Start from MFR data row
      const mfrRow = mfrDataValues[i];
      const mfrSku = String(mfrRow[mfrSkuCol]);
      if (!mfrSku) continue;

      const missingOnPlatforms = [];
      Object.keys(PLATFORM_TABS).forEach(pKey => {
        if (!platformPresence[mfrSku] || !platformPresence[mfrSku][pKey]) {
          missingOnPlatforms.push(pKey);
        }
      });

      if (missingOnPlatforms.length > 0) {
        const cpRow = Array(cpHeaders.length).fill('');
        cpRow[cpHeaderMap['Manufacturer SKU']] = mfrSku;
        if(cpHeaderMap['UPC'] !== undefined && upcCol !== -1) cpRow[cpHeaderMap['UPC']] = mfrRow[upcCol] || '';
        if(cpHeaderMap['MSRP'] !== undefined && msrpCol !== -1) cpRow[cpHeaderMap['MSRP']] = mfrRow[msrpCol];
        if(cpHeaderMap['MAP'] !== undefined && mapCol !== -1) cpRow[cpHeaderMap['MAP']] = mfrRow[mapCol];
        if(cpHeaderMap['Dealer Price'] !== undefined && dealerPriceCol !== -1) cpRow[cpHeaderMap['Dealer Price']] = mfrRow[dealerPriceCol];
        
        Object.keys(PLATFORM_TABS).forEach(pKey => {
          if(cpHeaderMap[pKey] !== undefined) { // Check if platform is a header in CP Listings
            cpRow[cpHeaderMap[pKey]] = (platformPresence[mfrSku] && platformPresence[mfrSku][pKey]) ? 'Yes' : 'No';
          }
        });
        if(cpHeaderMap['Action Needed'] !== undefined) {
            cpRow[cpHeaderMap['Action Needed']] = `List on: ${missingOnPlatforms.join(', ')}`;
        }
        cpDataRows.push(cpRow);
      }
    }

    if (cpDataRows.length > 0) {
      cpSheet.getRange(3, 1, cpDataRows.length, cpHeaders.length).setValues(cpDataRows);
      applyCPListingsFormatting(cpSheet, cpDataRows.length);
      cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge()
             .setValue(`CP LISTINGS - ${cpDataRows.length} items identified needing action - ${new Date().toLocaleString()}`).setBackground('#d9ead3');
    } else {
      cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge()
             .setValue('All items from Manufacturer sheet appear to be listed with high confidence across all platforms or no MFR items.').setBackground('#d9ead3');
    }
    ui.alert('CP Listings Generated', `${cpDataRows.length} items identified for cross-platform listing review.`, ui.ButtonSet.OK);

  } catch (error) {
    logError('generateCPListings', 'Error generating CP listings', '', error, true, ui);
    if (cpSheet) cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge().setValue('ERROR - ' + error.toString()).setBackground('#f4cccc');
  }
}

function applyCPListingsFormatting(sheet, rowCount) {
    if(!sheet || rowCount <= 0) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const dataStartRow = 3;
    let rules = sheet.getConditionalFormatRules(); // Get existing rules to add to them

    Object.keys(PLATFORM_TABS).forEach(pKey => {
        const colIndex = headers.indexOf(pKey); // Platform name itself is a header
        if (colIndex !== -1) {
            const range = sheet.getRange(dataStartRow, colIndex + 1, rowCount, 1);
            rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#d9ead3').setRanges([range]).build()); // Green for Yes
            rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#f4cccc').setRanges([range]).build());  // Red for No
        }
    });
    sheet.setConditionalFormatRules(rules);

    // Format price columns
    ['MSRP', 'MAP', 'Dealer Price'].forEach(colName => {
        const colIndex = headers.indexOf(colName);
        if (colIndex !== -1) {
            sheet.getRange(dataStartRow, colIndex + 1, rowCount, 1).setNumberFormat('$#,##0.00');
        }
    });
    Logger.log("Applied CP Listings formatting.");
}


function identifyDiscontinuedItems() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Identify Discontinued Items', 
    'Analyze platform SKUs (from "Platform Databases") that are not found in the "Manufacturer Price Sheet". This may indicate discontinued items. Results go to "Discontinued" tab. Continue?', 
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let discontinuedSheet;
  try {
    const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet');
    const platformDbSheet = SS.getSheetByName('Platform Databases'); // Source of all platform SKUs
    discontinuedSheet = SS.getSheetByName('Discontinued');
    if (!discontinuedSheet) discontinuedSheet = createDiscontinuedTab(); // Create if not exists

    if (!mfrSheet || !platformDbSheet) {
      logError('identifyDiscontinuedItems', 'Required sheets (MFR Price Sheet, Platform Databases) not found.', '', '', true, ui);
      return;
    }
    
    // Clear Discontinued sheet (from row 3 down)
    if (discontinuedSheet.getLastRow() > 2) {
        discontinuedSheet.getRange(3, 1, discontinuedSheet.getLastRow() - 2, discontinuedSheet.getLastColumn()).clearContent();
    }

    discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge()
                   .setValue('PROCESSING - Identifying discontinued items...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center');
    SpreadsheetApp.flush();

    // Get all MFR SKUs (original, normalized, and core) for efficient lookup
    const mfrDataValues = mfrSheet.getDataRange().getValues();
    const mfrHeaders = mfrDataValues[0].map(h => String(h).trim());
    const mfrSkuCol = mfrHeaders.indexOf('Manufacturer SKU');
    if (mfrSkuCol === -1) throw new Error('Manufacturer SKU column not found in Manufacturer Price Sheet.');

    const mfrOriginalSkuSet = new Set(); // Stores UPPERCASE original MFR SKUs
    const mfrNormalizedSkuSet = new Set(); // Stores conservatively normalized MFR SKUs
    const mfrCoreSkuSet = new Set();       // Stores extracted core SKUs from MFR items

    for (let i = 1; i < mfrDataValues.length; i++) {
      const rawMfrSku = String(mfrDataValues[i][mfrSkuCol]);
      if (rawMfrSku) {
        mfrOriginalSkuSet.add(rawMfrSku.toUpperCase());
        const mfrAttributes = extractSkuAttributesAndCore(rawMfrSku); // Use the refined extraction
        if (mfrAttributes.normalizedSku) mfrNormalizedSkuSet.add(mfrAttributes.normalizedSku);
        if (mfrAttributes.coreSku) mfrCoreSkuSet.add(mfrAttributes.coreSku);
      }
    }

    // Get all platform items from Platform Databases sheet
    const platformDataAll = getPlatformDataFromStructuredSheet(); // This returns {AMAZON: [{sku, price,...}], EBAY: [...]}
    if (Object.keys(platformDataAll).length === 0) {
        discontinuedSheet.getRange(2,1,1,discontinuedSheet.getLastColumn()).merge().setValue('No platform data found in "Platform Databases" sheet.').setBackground('#d9ead3');
        ui.alert('Info', 'No platform data found to analyze for discontinued items.', ui.ButtonSet.OK);
        return;
    }

    const discontinuedDataRows = [];
    const discHeaders = discontinuedSheet.getRange(1,1,1, discontinuedSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const discHeaderMap = {}; // Header name to column index
    discHeaders.forEach((h,i) => discHeaderMap[h] = i);
    const currentDate = new Date();

    for (const platformKey in platformDataAll) {
      platformDataAll[platformKey].forEach(item => {
        if (!item.sku) return;
        const platOriginalUpper = String(item.sku).toUpperCase();
        const platAttributes = extractSkuAttributesAndCore(item.sku); // Use refined extraction

        let foundInMfr = false;
        if (mfrOriginalSkuSet.has(platOriginalUpper)) foundInMfr = true;
        else if (platAttributes.normalizedSku && mfrNormalizedSkuSet.has(platAttributes.normalizedSku)) foundInMfr = true;
        else if (platAttributes.coreSku && mfrCoreSkuSet.has(platAttributes.coreSku)) {
            // If core matches, check if attributes (BStock, Color) are also compatible or if it's a simple variant
            // This helps avoid flagging a B-stock version of an active MFR SKU as discontinued.
            // A more sophisticated check here could compare mfrAtt.bstock with platAtt.bstock if core matches.
            // For now, a core match is considered "found" to be conservative about flagging as discontinued.
            foundInMfr = true; 
        }
        // Could add a stricter check: if core matches, do attributes suggest it's a *different* base item?

        if (!foundInMfr) {
          const row = Array(discHeaders.length).fill('');
          row[discHeaderMap['Platform SKU']] = item.sku;
          row[discHeaderMap['Platform']] = platformKey;
          row[discHeaderMap['Brand']] = extractBrandFromSku(item.sku, platformKey); // Helper to guess brand
          row[discHeaderMap['Current Price']] = item.price !== null && !isNaN(item.price) ? item.price : '';
          row[discHeaderMap['Last Updated']] = currentDate; // When this check was run
          row[discHeaderMap['Status']] = 'Potential Discontinued';
          // Confidence & Notes are placeholders for now
          if (discHeaderMap['Confidence'] !== undefined) row[discHeaderMap['Confidence']] = ''; 
          if (discHeaderMap['Notes'] !== undefined) row[discHeaderMap['Notes']] = 'Not found in current MFR sheet by SKU, normalized SKU, or core SKU.';
          discontinuedDataRows.push(row);
        }
      });
    }

    if (discontinuedDataRows.length > 0) {
      discontinuedSheet.getRange(3, 1, discontinuedDataRows.length, discHeaders.length).setValues(discontinuedDataRows);
      applyDiscontinuedFormatting(discontinuedSheet, discontinuedDataRows.length);
      discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge()
                     .setValue(`${discontinuedDataRows.length} potential discontinued items found - ${currentDate.toLocaleString()}`).setBackground('#d9ead3');
    } else {
      discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge()
                     .setValue('No potential discontinued items found based on current criteria.').setBackground('#d9ead3');
    }
    ui.alert('Discontinued Item Analysis Complete', 
             `${discontinuedDataRows.length} items found on platforms but not in the Manufacturer Price Sheet have been listed in the "Discontinued" tab.`, 
             ui.ButtonSet.OK);

  } catch (error) {
    logError('identifyDiscontinuedItems', 'Error identifying discontinued items', '', error, true, ui);
    if (discontinuedSheet) discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge().setValue('ERROR - ' + error.toString()).setBackground('#f4cccc');
  }
}

function extractBrandFromSku(sku, platform) {
    if (!sku) return 'Unknown';
    const upperSku = String(sku).toUpperCase();

    // Common brand prefixes often seen in SKUs. Order by length (longer first) if some are substrings of others.
    // This is a heuristic and might need refinement based on actual SKU patterns.
    const brandPrefixes = {
        'FENDER': 'Fender', 'GIBSON': 'Gibson', 'SQUIER': 'Squier', 'EPIPHONE': 'Epiphone', 
        'IBANEZ': 'Ibanez', 'YAMAHA': 'Yamaha', 'PRS': 'PRS', 'MARTIN': 'Martin', 'TAYLOR': 'Taylor',
        'BOSS': 'Boss', 'ROLAND': 'Roland', 'KORG': 'Korg', 'MOOG': 'Moog',
        'SHURE': 'Shure', 'SENNHEISER': 'Sennheiser', 'AKG': 'AKG', 'RODE': 'Rode',
        'MACKIE': 'Mackie', 'BEHRINGER': 'Behringer', 'PRESONUS': 'Presonus', 'FOCUSRITE': 'Focusrite',
        'EMG': 'EMG', 'DIMARZIO': 'DiMarzio', 'SEYMOUR DUNCAN': 'Seymour Duncan', 'SDPICKUPS': 'Seymour Duncan',
        'GODIN': 'Godin', 'G&L': 'G&L', 'GRETSCH': 'Gretsch', 'JACKSON': 'Jackson', 'CHARVEL': 'Charvel', 'EVH': 'EVH',
        'PEAVEY': 'Peavey', 'ORANGE': 'Orange Amps', 'MARSHALL': 'Marshall', 'VOX': 'Vox', 'LINE 6': 'Line 6',
        'DUNLOP': 'Dunlop', 'MXR': 'MXR', 'ELECTRO-HARMONIX': 'Electro-Harmonix', 'EHX': 'Electro-Harmonix',
        'STRYMON': 'Strymon', 'KEELEY': 'Keeley', 'WALRUS AUDIO': 'Walrus Audio', 'JHS': 'JHS Pedals',
        'ZILDJIAN': 'Zildjian', 'SABIAN': 'Sabian', 'PAISTE': 'Paiste', 'MEINL': 'Meinl',
        'PEARL': 'Pearl Drums', 'TAMA': 'Tama', 'DW': 'DW Drums', 'LUDWIG': 'Ludwig',
        // Shorter or common prefixes (ensure these don't incorrectly match longer names)
        'SHU-': 'Shure', 'MCK-': 'Mackie', 'GGC-': 'Godin', // Platform-specific prefixes if relevant
        'FEND': 'Fender', 'GIB': 'Gibson', 'IBZ': 'Ibanez', 'YMH': 'Yamaha', 
        'MART': 'Martin', 'TAYL': 'Taylor', 'RLND': 'Roland', 'SENNH': 'Sennheiser',
        'PRESO': 'Presonus', 'FOCU': 'Focusrite', 'SEYM': 'Seymour Duncan'
    };

    const parts = upperSku.split(/[-_ ]/); // Split by common delimiters
    if (parts.length > 0) {
        const firstPart = parts[0];
        if (brandPrefixes[firstPart]) return brandPrefixes[firstPart]; // Direct match on first part
        // Check if first part starts with a known prefix key (longer keys first)
        for (const pfxKey of Object.keys(brandPrefixes).sort((a,b) => b.length - a.length)) {
            if (firstPart.startsWith(pfxKey)) return brandPrefixes[pfxKey];
        }
    }
    // Fallback: check if any part of the SKU contains a known brand key (longer keys first)
    for (const brandKey of Object.keys(brandPrefixes).sort((a,b) => b.length - a.length)) {
        if (upperSku.includes(brandKey)) return brandPrefixes[brandKey];
    }
    return 'Unknown'; // Default if no brand identified
}

function applyDiscontinuedFormatting(sheet, rowCount) {
    if(!sheet || rowCount <= 0) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const dataStartRow = 3;
    let rules = sheet.getConditionalFormatRules(); // Get existing rules

    // Platform column formatting
    const platformCol = headers.indexOf('Platform');
    if (platformCol !== -1) {
        const range = sheet.getRange(dataStartRow, platformCol + 1, rowCount, 1);
        Object.keys(PLATFORM_TABS).forEach(pKey => {
            if (COLORS[pKey]) { // Use platform colors if defined
                const bgColor = COLORS[pKey];
                // Determine font color for readability based on background
                let fontColor = '#FFFFFF'; // Default white
                if (bgColor.match(/^#[0-9A-F]{6}$/i)) { // Basic hex check
                    const r = parseInt(bgColor.substring(1,3), 16);
                    const g = parseInt(bgColor.substring(3,5), 16);
                    const b = parseInt(bgColor.substring(5,7), 16);
                    // Simple brightness check (Luma formula)
                    if ((r*0.299 + g*0.587 + b*0.114) > 150) fontColor = '#000000'; // Use black for light backgrounds
                }
                rules.push(SpreadsheetApp.newConditionalFormatRule()
                    .whenTextEqualTo(pKey)
                    .setBackground(bgColor)
                    .setFontColor(fontColor)
                    .setRanges([range])
                    .build());
            }
        });
    }
    sheet.setConditionalFormatRules(rules);

    // Format Price column as currency
    const priceCol = headers.indexOf('Current Price');
    if (priceCol !== -1) {
        sheet.getRange(dataStartRow, priceCol + 1, rowCount, 1).setNumberFormat('$#,##0.00');
    }
    // Format Date column
    const dateCol = headers.indexOf('Last Updated');
    if (dateCol !== -1) {
        sheet.getRange(dataStartRow, dateCol + 1, rowCount, 1).setNumberFormat('M/d/yyyy h:mm:ss');
    }
    Logger.log("Applied Discontinued formatting.");
}