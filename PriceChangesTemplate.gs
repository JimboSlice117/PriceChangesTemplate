// @ts-nocheck
/**
 * Cross-Platform Price Management System
 * Complete Apps Script implementation
 * Accuracy-focused version with enhanced SKU matching and performance optimizations.
 * Batch processing has been removed. SKU Match Review sidebar has been removed.
 * Includes a dedicated "Price Changes" tab.
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
const ANALYSIS_PLATFORMS = ['AMAZON', 'EBAY', 'SHOPIFY', 'REVERB'];
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
const skuNormalizeCache = {};

// ---------------- UTILITY FUNCTIONS ----------------

function getBStockInfo(sku) {
  if (!sku || (typeof sku !== 'string' && typeof sku !== 'number')) return null;
  const upperSku = String(sku).toUpperCase();
  for (const [bStockType, multiplierOrSpecial] of Object.entries(BSTOCK_CATEGORIES)) {
    if (upperSku.includes('-' + bStockType) || upperSku.startsWith(bStockType + '-') || upperSku.endsWith('-' + bStockType) ) {
      let actualMultiplier; let isSpecial = false;
      if (typeof multiplierOrSpecial === 'number') actualMultiplier = multiplierOrSpecial;
      else { isSpecial = true; if (bStockType === 'NOACC') actualMultiplier = BSTOCK_CATEGORIES['BC']; else if (bStockType === 'AA') actualMultiplier = 0.98; else actualMultiplier = 0.85; }
      return { type: bStockType, multiplier: actualMultiplier, sku: sku, isSpecial: isSpecial };
    }
  }
  return null;
}
function getColumnIndices(headers, columnNames) { const indices = {}; columnNames.forEach(name => { indices[name] = headers.indexOf(name); }); return indices; }
function getPlatformVariations(platform) { const variations = [platform, platform.charAt(0) + platform.slice(1).toLowerCase(), platform.toLowerCase()]; if (platform === 'SELLERCLOUD') variations.push('ellerCloud'); if (platform === 'INFLOW') variations.push('inFlow'); return variations.filter(Boolean); }
function logError(error, functionName, ui, showAlert = true) { Logger.log(`Error in ${functionName}: ${error.toString()}\nStack: ${error.stack}`); if (showAlert && ui) ui.alert('Error', `An error occurred in ${functionName}: ${error.toString()}`, ui.ButtonSet.OK); }
function longestCommonSubstring(str1, str2) { if (!str1 || !str2) return ''; const s1 = [...str1], s2 = [...str2]; const matrix = Array(s1.length + 1).fill(null).map(() => Array(s2.length + 1).fill(0)); let maxLength = 0, endPosition = 0; for (let i = 1; i <= s1.length; i++) for (let j = 1; j <= s2.length; j++) if (s1[i - 1] === s2[j - 1]) { matrix[i][j] = matrix[i - 1][j - 1] + 1; if (matrix[i][j] > maxLength) { maxLength = matrix[i][j]; endPosition = i; } } return str1.substring(endPosition - maxLength, endPosition); }

// ---------------- ACCURACY-FOCUSED SKU MATCHING FUNCTIONS ----------------
function conservativeNormalizeSku(sku) { if (!sku || (typeof sku !== 'string' && typeof sku !== 'number')) return ''; const skuStr = String(sku); if (skuNormalizeCache[skuStr]) return skuNormalizeCache[skuStr]; let normalized = skuStr.toUpperCase().replace(/[^A-Z0-9\-]/g, '').replace(/\-+/g, '-').replace(/^-|-$/g, ''); skuNormalizeCache[skuStr] = normalized; return normalized; }
function conservativeExtractCore(sku) { if (!sku) return ''; let core = sku; const obviousPrefixes = [/^EMG-/, /^SHU-/, /^BOSS-/, /^MCK-/, /^GGACC-/, /^1SV-/, /^AA-/, /^360H-/, /^8SI-/]; for (const prefix of obviousPrefixes) if (prefix.test(core)) { core = core.replace(prefix, ''); break; } const obviousSuffixes = [/-FOL$/, /-FOLIOS?$/]; for (const suffix of obviousSuffixes) if (suffix.test(core)) { core = core.replace(suffix, ''); break; } return core.replace(/^-|-$/g, ''); }
function extractSkuAttributesAndCore(rawSku) { if (!rawSku || (typeof rawSku !== 'string' && typeof rawSku !== 'number')) return { originalSku: rawSku, normalizedSku: '', coreSku: '', bStock: null, color: null }; const originalSkuUpper = String(rawSku).toUpperCase(); const normalizedForProcessing = conservativeNormalizeSku(originalSkuUpper); const bStockInfo = getBStockInfo(originalSkuUpper); let tempSkuForCore = normalizedForProcessing; if (bStockInfo) { const bStockType = bStockInfo.type; const patternsToRemove = ['-' + bStockType + '-', '-' + bStockType, bStockType + '-']; for (const pattern of patternsToRemove) if (tempSkuForCore.includes(pattern)) tempSkuForCore = tempSkuForCore.replace(new RegExp(pattern.replace(/-/g, '\\-'), 'g'), pattern === `-${bStockType}-` ? '-' : ''); tempSkuForCore = tempSkuForCore.replace(/^-|-$/g, '').replace(/\-\-/g,'-'); } let foundColor = null; let skuAfterColorRemoval = tempSkuForCore; const sortedColorAbbrs = Object.keys(COLOR_ABBREVIATIONS).sort((a, b) => b.length - a.length); for (const abbr of sortedColorAbbrs) { const upperAbbr = abbr.toUpperCase(); let replaced = false; if (skuAfterColorRemoval.endsWith('-' + upperAbbr)) { skuAfterColorRemoval = skuAfterColorRemoval.substring(0, skuAfterColorRemoval.length - (upperAbbr.length + 1)); foundColor = COLOR_ABBREVIATIONS[abbr]; replaced = true; } if (!replaced && skuAfterColorRemoval.endsWith(upperAbbr)) { const precedingCharIndex = skuAfterColorRemoval.length - upperAbbr.length - 1; if (precedingCharIndex < 0 || !/[A-Z0-9]/.test(skuAfterColorRemoval.charAt(precedingCharIndex)) || upperAbbr.length > 2) { skuAfterColorRemoval = skuAfterColorRemoval.substring(0, skuAfterColorRemoval.length - upperAbbr.length); foundColor = COLOR_ABBREVIATIONS[abbr]; replaced = true; } } if (replaced) break; } skuAfterColorRemoval = skuAfterColorRemoval.replace(/-$/, ''); let finalCore = conservativeExtractCore(skuAfterColorRemoval); finalCore = finalCore.replace(/^-|-$/g, ''); return { originalSku: rawSku, normalizedSku: normalizedForProcessing, coreSku: finalCore, bStock: bStockInfo, color: foundColor }; }
function levenshteinDistance(a, b) { if (!a && !b) return 0; if (!a) return b.length; if (!b) return a.length; const matrix = Array(b.length + 1).fill(null).map(() => Array(a.length + 1).fill(0)); for (let i = 0; i <= a.length; i++) matrix[0][i] = i; for (let j = 0; j <= b.length; j++) matrix[j][0] = j; for (let j = 1; j <= b.length; j++) for (let i = 1; i <= a.length; i++) { const cost = a[i - 1] === b[j - 1] ? 0 : 1; matrix[j][i] = Math.min(matrix[j][i - 1] + 1, matrix[j - 1][i] + 1, matrix[j - 1][i - 1] + cost); } return matrix[b.length][a.length]; }
function strictSimilarityCheck(coreSku1, coreSku2) { if (!coreSku1 && !coreSku2) return { score: 5, reason: 'Both cores empty (attribute-only SKUs)'}; if (!coreSku1 || !coreSku2) return { score: 0, reason: 'Empty Core SKU' }; const len1 = coreSku1.length; const len2 = coreSku2.length; if (len1 < 2 || len2 < 2) { if (coreSku1 === coreSku2) return { score: 90, reason: "Exact short core match" }; const shortDist = levenshteinDistance(coreSku1, coreSku2); if (shortDist <=1 && Math.max(len1, len2) <=2) return {score: 70, reason: `Near exact short core (dist ${shortDist})`}; return { score: 0, reason: `Core too short for non-exact (c1:${len1}, c2:${len2})` }; } const lengthRatio = Math.min(len1, len2) / Math.max(len1, len2); if (lengthRatio < 0.45) return { score: 0, reason: `Core length ratio too low (${lengthRatio.toFixed(2)})` }; if (coreSku1 === coreSku2) return { score: 95, reason: `Exact core match (${coreSku1})` }; const coreDistance = levenshteinDistance(coreSku1, coreSku2); const coreMaxLength = Math.max(len1, len2); const coreSimilarity = ((coreMaxLength - coreDistance) / coreMaxLength); if (coreSimilarity >= 0.88) return { score: Math.min(90, 65 + Math.round(coreSimilarity * 30)), reason: `High core similarity: ${Math.round(coreSimilarity * 100)}% (dist ${coreDistance})` }; if (len1 >= 3 && coreSku2.includes(coreSku1)) return { score: 88, reason: `Core1 (${coreSku1}) in Core2 (${coreSku2})` }; if (len2 >= 3 && coreSku1.includes(coreSku2)) return { score: 88, reason: `Core2 (${coreSku2}) in Core1 (${coreSku1})` }; const commonSub = longestCommonSubstring(coreSku1, coreSku2); if (commonSub.length >= Math.max(2, Math.min(len1, len2) * 0.5)) { const overlapRatioCore = commonSub.length / Math.min(len1, len2); if (overlapRatioCore >= 0.60) return { score: Math.min(85, 55 + Math.round(overlapRatioCore * 35)), reason: `Strong core overlap: ${commonSub} (${Math.round(overlapRatioCore * 100)}%)` }; } if (coreSimilarity >= 0.75) return { score: Math.min(80, 50 + Math.round(coreSimilarity * 35)), reason: `Good core similarity: ${Math.round(coreSimilarity * 100)}% (dist ${coreDistance})` }; return { score: Math.max(0, Math.round(coreSimilarity * 60)), reason: `Low core similarity: ${Math.round(coreSimilarity*100)}% (dist ${coreDistance})` }; }
function conservativePlatformMatch(mfrSkuNormalized, platformSkuNormalized, platform) { if (!mfrSkuNormalized || !platformSkuNormalized) return { score: 0, reason: 'Empty SKU for platform match' }; let score = 0; let reason = ''; const mfrCorePlat = conservativeExtractCore(mfrSkuNormalized); const platCorePlat = conservativeExtractCore(platformSkuNormalized); switch (platform) { case 'AMAZON': if (platformSkuNormalized.endsWith('-FOL') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Amazon FOL suffix, base match'; } else if ((platformSkuNormalized.startsWith('360H-') || platformSkuNormalized.startsWith('8SI-')) && platformSkuNormalized.substring(platformSkuNormalized.indexOf('-')+1) === mfrSkuNormalized) { score = 96; reason = 'Amazon prefix, exact mfr SKU'; } break; case 'SELLERCLOUD': if (platformSkuNormalized.startsWith('EMG-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'SC EMG prefix, base match'; } else if (platformSkuNormalized.startsWith('EMG-') && platformSkuNormalized.substring(4) === mfrSkuNormalized) { score = 96; reason = 'SC EMG prefix, exact mfr SKU';} break; case 'REVERB': if (platformSkuNormalized.startsWith('SHU-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Reverb SHU prefix, base match'; } else if (platformSkuNormalized.startsWith('BOSS-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Reverb BOSS prefix, base match'; } else if (platformSkuNormalized.startsWith('MCK-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'Reverb MCK prefix, base match'; } break; case 'EBAY': if (platformSkuNormalized.startsWith('GGACC-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'eBay GGACC prefix, base match'; } break; case 'SHOPIFY': if (platformSkuNormalized.startsWith('AA-') && mfrCorePlat === platCorePlat && mfrCorePlat.length >= 4 && !getBStockInfo(platformSkuNormalized)) { score = 92; reason = 'Shopify AA prefix (non-BStock), base match (len>=4)'; } break; case 'INFLOW': if (platformSkuNormalized.startsWith('1SV-') && mfrCorePlat === platCorePlat) { score = 95; reason = 'inFlow 1SV prefix, base match'; } break; } const mfrBstockPlat = getBStockInfo(mfrSkuNormalized); const platBstockPlat = getBStockInfo(platformSkuNormalized); if (mfrBstockPlat && platBstockPlat && mfrBstockPlat.type === platBstockPlat.type) { let baseMfr = mfrSkuNormalized.replace('-'+mfrBstockPlat.type, '').replace(mfrBstockPlat.type+'-','').replace(/^-|-$/g, ''); let basePlat = platformSkuNormalized.replace('-'+platBstockPlat.type, '').replace(platBstockPlat.type+'-','').replace(/^-|-$/g, ''); if (baseMfr === basePlat && baseMfr.length > 0) { score = Math.max(score, 97); reason = (reason ? reason + "; " : "") + `Platform B-Stock (${mfrBstockPlat.type}) base match`; } } else if (!mfrBstockPlat && platBstockPlat) { let basePlat = platformSkuNormalized.replace('-'+platBstockPlat.type, '').replace(platBstockPlat.type+'-','').replace(/^-|-$/g, ''); if (mfrSkuNormalized === basePlat && mfrSkuNormalized.length > 0) { score = Math.max(score, 96); reason = (reason ? reason + "; " : "") + `Platform B-Stock (${platBstockPlat.type}) matches new MFR SKU base`; } } return { score: score, reason: reason }; }
function validateMatch_optimized(mfrSkuRaw, mfrAtt, platformSkuRaw, platAtt, platform) { if (!mfrSkuRaw || !platformSkuRaw) return { valid: false, confidence: 0, reason: 'Empty SKU provided to validateMatch' }; mfrAtt = mfrAtt || { originalSku: mfrSkuRaw, normalizedSku: '', coreSku: '', bStock: null, color: null }; platAtt = platAtt || { originalSku: platformSkuRaw, normalizedSku: '', coreSku: '', bStock: null, color: null }; if (String(mfrSkuRaw).toUpperCase() === String(platformSkuRaw).toUpperCase()) return { valid: true, confidence: 100, reason: 'Exact raw match (case-insensitive)' }; if (mfrAtt.normalizedSku && platAtt.normalizedSku && mfrAtt.normalizedSku === platAtt.normalizedSku) return { valid: true, confidence: 99, reason: `Exact normalized match (${mfrAtt.normalizedSku})` }; if ((mfrAtt.normalizedSku && mfrAtt.normalizedSku.length < 3) || (platAtt.normalizedSku && platAtt.normalizedSku.length < 3)) { const dist = levenshteinDistance(mfrAtt.normalizedSku, platAtt.normalizedSku); const maxL = Math.max(mfrAtt.normalizedSku.length, platAtt.normalizedSku.length); if (dist === 0 && maxL > 0) return { valid: true, confidence: 98, reason: `Short exact normalized (${mfrAtt.normalizedSku})`}; if (dist <= 1 && maxL <= 4 && maxL > 0) return { valid: true, confidence: 70 + (5-maxL)*2, reason: `Short SKU near match (dist ${dist}, len ${maxL})` }; const sim = maxL > 0 ? (maxL - dist) / maxL : 0; if (sim >= 0.60 && maxL > 0) return { valid: 'NEEDS_REVIEW', confidence: Math.max(50, Math.round(sim * 60)), reason: `Short SKU, mod. similarity ${Math.round(sim * 100)}%` }; return { valid: false, confidence: 0, reason: `SKU too short (MfrN:${mfrAtt.normalizedSku}, PlatN:${platAtt.normalizedSku})` }; } const platformMatchResult = conservativePlatformMatch(mfrAtt.normalizedSku, platAtt.normalizedSku, platform); const coreSimilarityResult = strictSimilarityCheck(mfrAtt.coreSku, platAtt.coreSku); let attributeBonus = 0; let attributeReason = ""; let coreIdenticalAndAttributesDiffer = false; if (mfrAtt.coreSku && platAtt.coreSku && mfrAtt.coreSku === platAtt.coreSku && mfrAtt.coreSku !== "") attributeBonus += 5; if (mfrAtt.bStock && platAtt.bStock) { if (mfrAtt.bStock.type === platAtt.bStock.type) { attributeBonus += 15; attributeReason += `B-Stock type match (${mfrAtt.bStock.type}). `; } else { attributeBonus -= 10; attributeReason += `B-Stock type mismatch (${mfrAtt.bStock.type} vs ${platAtt.bStock.type}). `; if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true; } } else if (mfrAtt.bStock || platAtt.bStock) { const bStockSource = mfrAtt.bStock ? `MFR (${mfrAtt.bStock.type})` : `Plat (${platAtt.bStock.type})`; if (mfrAtt.coreSku === platAtt.coreSku && mfrAtt.coreSku !== "") { attributeBonus += 8; attributeReason += `Expected B-Stock diff (${bStockSource}). `; } else { attributeBonus -= 5; attributeReason += `B-Stock presence diff (${bStockSource}). `; } if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true; } if (mfrAtt.color && platAtt.color) { if (mfrAtt.color === platAtt.color) { attributeBonus += 15; attributeReason += `Color match (${mfrAtt.color}). `; } else { attributeBonus -= 10; attributeReason += `Color mismatch (${mfrAtt.color} vs ${platAtt.color}). `; if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true; } } else if (mfrAtt.color || platAtt.color) { const colorSource = mfrAtt.color ? `MFR (${mfrAtt.color})` : `Plat (${platAtt.color})`; if (mfrAtt.coreSku === platAtt.coreSku && mfrAtt.coreSku !== "") { attributeBonus += 8; attributeReason += `Expected Color diff (${colorSource}). `; } else { attributeBonus -= 5; attributeReason += `Color presence diff (${colorSource}). `; } if (mfrAtt.coreSku === platAtt.coreSku) coreIdenticalAndAttributesDiffer = true; } attributeBonus = Math.max(-20, Math.min(attributeBonus, 30)); let finalScore = 0; let finalReason = ""; if (coreSimilarityResult.score > 0) { finalReason = `CoreSim(${coreSimilarityResult.score}%): ${coreSimilarityResult.reason}. `; if (attributeReason) finalReason += `Attr: ${attributeReason}(Bonus ${attributeBonus}). `; finalScore = coreSimilarityResult.score + attributeBonus; } else { const fullNormalizedDistance = levenshteinDistance(mfrAtt.normalizedSku, platAtt.normalizedSku); const fullNormalizedMaxLength = Math.max(mfrAtt.normalizedSku.length, platAtt.normalizedSku.length); const fullNormalizedSimilarity = fullNormalizedMaxLength > 0 ? ((fullNormalizedMaxLength - fullNormalizedDistance) / fullNormalizedMaxLength) : 0; let baseScore = Math.round(fullNormalizedSimilarity * 60); finalReason = `Low CoreSim. FullNormSim(${Math.round(fullNormalizedSimilarity*100)}%). `; if (attributeReason) finalReason += `Attr: ${attributeReason}(Bonus ${attributeBonus}). `; finalScore = baseScore + attributeBonus; } if (coreIdenticalAndAttributesDiffer && finalScore > 70) finalReason += "Cores matched but attributes differed. "; if (platformMatchResult.score > 0) { if (platformMatchResult.score > finalScore + 10 || (platformMatchResult.score >= 90 && finalScore < 90) ) { finalScore = platformMatchResult.score; finalReason = `Platform Specific: ${platformMatchResult.reason} (Core/Attr: ${finalScore > 0 ? finalScore : 'N/A'})`; } else if (finalScore < 50 && platformMatchResult.score > finalScore) { finalScore = platformMatchResult.score; finalReason = `Platform Hint: ${platformMatchResult.reason} (Low core/attr).`; } else { finalReason += ` PlatformNote: ${platformMatchResult.reason} (Score ${platformMatchResult.score}).`; finalScore = Math.max(finalScore, platformMatchResult.score); } } finalScore = Math.min(98, Math.max(0, Math.round(finalScore))); if (finalScore >= 85) return { valid: true, confidence: finalScore, reason: finalReason }; if (finalScore >= 70) return { valid: 'NEEDS_REVIEW', confidence: finalScore, reason: `Review: ${finalReason}` }; return { valid: false, confidence: finalScore, reason: `Low Confidence: ${finalReason}` }; }
function accurateFindBestMatch_optimized(rawMfrSku, mfrExtractedAttributes, platformItemsWithAttributes, platformSkuMap, cleanSkuMap, platform) { const normalizedMfrSkuForMapLookup = mfrExtractedAttributes.normalizedSku; if (platformSkuMap && platformSkuMap[normalizedMfrSkuForMapLookup]) { const platItemContainer = platformSkuMap[normalizedMfrSkuForMapLookup]; const validation = validateMatch_optimized(rawMfrSku, mfrExtractedAttributes, platItemContainer.sku, platItemContainer.extractedAttributes, platform); if (validation.confidence >= 98) return { platformSku: platItemContainer.sku, currentPrice: platItemContainer.price, currentCost: platItemContainer.cost, confidenceScore: validation.confidence, matchType: 'Exact Normalized (Validated)', matchReason: validation.reason }; } if (cleanSkuMap && cleanSkuMap[normalizedMfrSkuForMapLookup]) { const platItemContainer = cleanSkuMap[normalizedMfrSkuForMapLookup]; const platCleanAtt = platItemContainer.cleanExtractedAttributes || platItemContainer.extractedAttributes; const validation = validateMatch_optimized(rawMfrSku, mfrExtractedAttributes, platItemContainer.sku, platCleanAtt, platform); if (validation.confidence >= 97) return { platformSku: platItemContainer.sku, currentPrice: platItemContainer.price, currentCost: platItemContainer.cost, confidenceScore: validation.confidence, matchType: 'Clean Exact Normalized (Validated)', matchReason: validation.reason }; } let bestMatch = null; let highestConfidence = 0; for (const itemContainer of platformItemsWithAttributes) { if (!itemContainer.sku || !itemContainer.extractedAttributes) continue; const validation = validateMatch_optimized(rawMfrSku, mfrExtractedAttributes, itemContainer.sku, itemContainer.extractedAttributes, platform); if (validation.confidence > highestConfidence) { highestConfidence = validation.confidence; bestMatch = { platformSku: itemContainer.sku, currentPrice: itemContainer.price, currentCost: itemContainer.cost, confidenceScore: validation.confidence, matchType: validation.valid === true ? 'Validated-Strong' : (validation.valid === 'NEEDS_REVIEW' ? 'Validated-Needs-Review' : 'Validated-Low-Confidence'), matchReason: validation.reason }; } } if (bestMatch) { if (bestMatch.confidenceScore >= 85) return bestMatch; if (bestMatch.confidenceScore >= 70) { bestMatch.matchType = 'LOW-CONFIDENCE-REVIEW-REQUIRED'; bestMatch.matchReason = `REVIEW (Score ${bestMatch.confidenceScore}): ${bestMatch.matchReason}`; return bestMatch; } } return null; }

// ---------------- PRE-PROCESSING OF PLATFORM DATA ----------------
function preProcessAllPlatformData(platformDataRaw) {
  const platformSkuMaps = {};
  const cleanSkuMaps = {};
  const preProcessedPlatformDataWithAttributes = {};

  Logger.log("Starting pre-processing of all platform data...");
  for (const platform in platformDataRaw) {
    platformSkuMaps[platform] = {};
    cleanSkuMaps[platform] = {};
    preProcessedPlatformDataWithAttributes[platform] = platformDataRaw[platform].map(item => {
      if (!item.sku) return { ...item, extractedAttributes: null, cleanExtractedAttributes: null };

      const attributes = extractSkuAttributesAndCore(item.sku);
      const itemWithAttributes = { ...item, extractedAttributes: attributes, cleanExtractedAttributes: null };

      if (attributes && attributes.normalizedSku) {
        platformSkuMaps[platform][attributes.normalizedSku] = itemWithAttributes;
      }

      if (item.cleanSku) {
        const cleanAttributes = extractSkuAttributesAndCore(item.cleanSku);
        itemWithAttributes.cleanExtractedAttributes = cleanAttributes;
        if (cleanAttributes && cleanAttributes.normalizedSku) {
          cleanSkuMaps[platform][cleanAttributes.normalizedSku] = itemWithAttributes;
        }
      }
      return itemWithAttributes;
    });
    Logger.log(`Platform ${platform} (for pre-processing): Processed ${preProcessedPlatformDataWithAttributes[platform].length} items.`);
  }
  Logger.log("Full platform data pre-processing complete.");
  return { preProcessedPlatformDataWithAttributes, platformSkuMaps, cleanSkuMaps };
}


// ---------------- CORE MATCHING LOGIC (SINGLE RUN) ----------------
function performFullMatching(manufacturerData, preProcessedPlatformData, platformSkuMaps, cleanSkuMaps) {
  const matchResults = [];
  Object.keys(skuNormalizeCache).forEach(key => delete skuNormalizeCache[key]);

  manufacturerData.forEach((mfrItem, index) => {
    if (index > 0 && index % 100 === 0) { // Log progress every 100 MFR SKUs
        Logger.log(`Processing MFR SKU ${index + 1} of ${manufacturerData.length}: ${mfrItem.manufacturerSku}`);
        SpreadsheetApp.flush(); // Allow logs to appear and potentially prevent timeouts on very long loops
    }
    const rawMfrSku = mfrItem.manufacturerSku;
    const mfrExtractedAttributes = extractSkuAttributesAndCore(rawMfrSku);

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
          mfrExtractedAttributes,
          preProcessedPlatformData[platform],
          platformSkuMaps[platform],
          cleanSkuMaps[platform],
          platform
        );
      } catch (error) {
        logError(error, `performFullMatching for ${rawMfrSku} on ${platform}`, null, false);
        result.matches[platform] = null;
      }
    }
    matchResults.push(result);
  });
  return matchResults;
}


// ---------------- MENU FUNCTIONS ---------------- 
function onOpen() { const ui = SpreadsheetApp.getUi(); ui.createMenu('Price Management').addItem('Safe Setup System (Preserve Data)', 'safeSetupPriceManagementSystem').addSeparator().addItem('Run Accurate SKU Matching', 'runSkuMatching').addItem('Validate Match Quality', 'validateMatchQuality').addItem('Save Matches to History', 'saveMatches').addSeparator().addItem('Update Price Analysis', 'updatePriceAnalysis').addItem('Generate Export Files', 'generateExports').addSeparator().addItem('Generate CP Listings', 'generateCPListings').addItem('Identify Discontinued Items', 'identifyDiscontinuedItems').addToUi(); }

// ---------------- SETUP FUNCTIONS ---------------- 
function safeSetupPriceManagementSystem() { const ui = SpreadsheetApp.getUi(); const response = ui.alert('Safe Setup', 'Create missing tabs & preserve data?', ui.ButtonSet.YES_NO); if (response !== ui.Button.YES) return; const existingSheets = {}; SS.getSheets().forEach(sheet => { existingSheets[sheet.getName()] = true; }); if (!existingSheets['Manufacturer Price Sheet']) createManufacturerPriceSheet(); if (!existingSheets['SKU Matching Engine']) createSkuMatchingEngineTab(); if (!existingSheets['Price Analysis Dashboard']) createPriceAnalysisDashboard(); if (!existingSheets['Match History']) createMatchHistoryTab(); if (!existingSheets['Instructions']) createInstructionsTab(); if (!existingSheets['CP Listings']) createCPListingsTab(); if (!existingSheets['Discontinued']) createDiscontinuedTab(); if (!existingSheets['Price Changes']) createPriceChangesTab(); Object.keys(PLATFORM_TABS).forEach(platform => { const tabName = PLATFORM_TABS[platform].name; if (!existingSheets[tabName]) createExportTab(platform); }); ui.alert('Setup Complete', 'Missing tabs created. Data preserved.', ui.ButtonSet.OK); }
function createManufacturerPriceSheet() { const sheet = SS.insertSheet('Manufacturer Price Sheet'); const headers = ['Manufacturer SKU', 'UPC', 'MSRP', 'MAP', 'Dealer Price']; sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(2, 1, 1, headers.length).merge().setValue('Paste MFR price sheet data below.').setFontStyle('italic').setHorizontalAlignment('center'); sheet.getRange(3, 1, 100, headers.length).setBackground('#f3f3f3'); sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 150); sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 100); }
function createSkuMatchingEngineTab() { const sheet = SS.insertSheet('SKU Matching Engine'); const headers = ['Manufacturer SKU', 'MSRP', 'MAP', 'Dealer Price']; Object.keys(PLATFORM_TABS).forEach(platform => { headers.push(`${platform} SKU`, `${platform} Confidence`, `${platform} Current Price`); }); headers.push('Status', 'Notes'); sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(2, 1, 1, headers.length).merge().setValue('Matches shown here. Use "Run Accurate SKU Matching" to populate.').setFontStyle('italic').setHorizontalAlignment('center'); sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 80); sheet.setColumnWidth(4, 80); let colIndex = 5; Object.keys(PLATFORM_TABS).forEach(() => { sheet.setColumnWidth(colIndex, 200); sheet.setColumnWidth(colIndex + 1, 100); sheet.setColumnWidth(colIndex + 2, 100); colIndex += 3; }); sheet.setColumnWidth(colIndex, 100); sheet.setColumnWidth(colIndex + 1, 200); }
function createPriceAnalysisDashboard() { const sheet = SS.insertSheet('Price Analysis Dashboard'); sheet.getRange(1, 1, 1, 12).merge().setValue('PRICE ANALYSIS DASHBOARD').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center').setBackground('#e0e0e0'); sheet.getRange(2, 1, 1, 12).merge().setValue('SUMMARY METRICS').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f3f3f3'); sheet.getRange(3, 1).setValue('Total Products:'); sheet.getRange(3, 2).setValue('0'); sheet.getRange(3, 4).setValue('Price Increases:'); sheet.getRange(3, 5).setValue('0'); sheet.getRange(3, 7).setValue('Price Decreases:'); sheet.getRange(3, 8).setValue('0'); sheet.getRange(3, 10).setValue('Average Change:'); sheet.getRange(3, 11).setValue('0%'); sheet.getRange(4, 1).setValue('MAP Violations:'); sheet.getRange(4, 2).setValue('0'); sheet.getRange(4, 4).setValue('High Impact Changes:'); sheet.getRange(4, 5).setValue('0'); sheet.getRange(4, 7).setValue('Unmatched Products:'); sheet.getRange(4, 8).setValue('0'); sheet.getRange(4, 10).setValue('B-Stock Changes:'); sheet.getRange(4, 11).setValue('0'); sheet.getRange(3, 1, 2, 1).setFontWeight('bold'); sheet.getRange(3, 4, 2, 1).setFontWeight('bold'); sheet.getRange(3, 7, 2, 1).setFontWeight('bold'); sheet.getRange(3, 10, 2, 1).setFontWeight('bold'); const analysisHeaders = ['SKU', 'MAP', 'Dealer', 'B-Stock', 'Platform', 'Current', 'New', 'Change $', 'Change %', 'Status']; sheet.getRange(6, 1, 1, analysisHeaders.length).setValues([analysisHeaders]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(5, 1, 1, analysisHeaders.length).merge().setValue('Use "Update Price Analysis" to refresh.').setFontStyle('italic').setHorizontalAlignment('center'); sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 80); sheet.setColumnWidth(4, 80); sheet.setColumnWidth(5, 100); sheet.setColumnWidth(6, 100); sheet.setColumnWidth(7, 100); sheet.setColumnWidth(8, 100); sheet.setColumnWidth(9, 100); sheet.setColumnWidth(10, 120); sheet.getRange(7, 1, 100, analysisHeaders.length).setBackground('#f8f9fa'); }
function createMatchHistoryTab() { const sheet = SS.insertSheet('Match History'); const headers = ['Manufacturer SKU', 'Platform', 'Platform SKU', 'Match Type', 'Confidence', 'Date Added', 'User Notes']; sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(2, 1, 1, headers.length).merge().setValue('Confirmed SKU matches.').setFontStyle('italic').setHorizontalAlignment('center'); sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 100); sheet.setColumnWidth(3, 200); sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 100); sheet.setColumnWidth(6, 120); sheet.setColumnWidth(7, 200); sheet.getRange(3, 1, 100, headers.length).setBackground('#f8f9fa'); }
function createInstructionsTab() { const sheet = SS.insertSheet('Instructions'); const instructions = [['CROSS-PLATFORM PRICE MANAGEMENT SYSTEM'],['See script comments for detailed logic.']]; for (let i = 0; i < instructions.length; i++) sheet.getRange(i + 1, 1).setValue(instructions[i][0]); sheet.getRange(1, 1).setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center'); sheet.setColumnWidth(1, 600); }
function createCPListingsTab() { const sheet = SS.insertSheet('CP Listings'); const headers = ['Manufacturer SKU', 'UPC', 'MSRP', 'MAP', 'Dealer Price', 'Amazon', 'eBay', 'Shopify', 'Reverb', 'inFlow', 'SellerCloud', 'Action Needed']; sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(2, 1, 1, headers.length).merge().setValue('Items from MFR sheet needing listing.').setFontStyle('italic').setHorizontalAlignment('center'); sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 150); sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 100); for (let i = 6; i <= 11; i++) sheet.setColumnWidth(i, 100); sheet.setColumnWidth(12, 200); sheet.getRange(3, 1, 100, headers.length).setBackground('#f8f9fa'); }
function createDiscontinuedTab() { const sheet = SS.insertSheet('Discontinued'); const headers = ['Platform SKU', 'Platform', 'Brand', 'Current Price', 'Last Updated', 'Confidence', 'Status', 'Notes']; sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(2, 1, 1, headers.length).merge().setValue('Platform SKUs not in MFR sheet.').setFontStyle('italic').setHorizontalAlignment('center'); sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 100); sheet.setColumnWidth(3, 150); sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 120); sheet.setColumnWidth(6, 100); sheet.setColumnWidth(7, 150); sheet.setColumnWidth(8, 250); sheet.getRange(3, 1, 100, headers.length).setBackground('#f8f9fa'); return sheet; }
function createExportTab(platform) { const platformInfo = PLATFORM_TABS[platform]; const tabName = platformInfo.name; const headers = platformInfo.headers; const sheet = SS.insertSheet(tabName); sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(COLORS[platform]).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center'); sheet.getRange(2, 1, 1, headers.length).merge().setValue('NOT READY - Run "Generate Export Files"').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f4b400').setFontColor('#000000'); sheet.getRange(3, 1, 1, headers.length).merge().setValue('Data ready to export to ' + tabName.replace(' Export', '')).setFontStyle('italic').setHorizontalAlignment('center'); for (let i = 0; i < headers.length; i++) sheet.setColumnWidth(i + 1, 150); sheet.getRange(4, 1, 100, headers.length).setBackground('#f8f9fa'); }

function createPriceChangesTab() {
  const sheetName = "Price Changes";
  let sheet = SS.getSheetByName(sheetName);
  if (sheet) {
    Logger.log(`Sheet "${sheetName}" already exists. Content will be overwritten or appended.`);
  } else {
    sheet = SS.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" created.`);
  }

  sheet.getRange("A1").setValue("PRICE INCREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_UP);
  const headersUp = ["MFR SKU", "Platform", "Old Price", "New Price", "Change $", "Change %"];
  sheet.getRange(2, 1, 1, headersUp.length).setValues([headersUp]).setFontWeight("bold").setBackground("#f2f2f2");
  sheet.setColumnWidth(1, 200); 
  sheet.setColumnWidth(2, 100); 
  sheet.setColumnWidth(3, 100); 
  sheet.setColumnWidth(4, 100); 
  sheet.setColumnWidth(5, 100); 
  sheet.setColumnWidth(6, 100); 

  const startRowDecreases = Math.max(4, sheet.getLastRow() + 3); 

  sheet.getRange(startRowDecreases, 1).setValue("PRICE DECREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_DOWN);
  const headersDown = ["MFR SKU", "Platform", "Old Price", "New Price", "Change $", "Change %"];
  sheet.getRange(startRowDecreases + 1, 1, 1, headersDown.length).setValues([headersDown]).setFontWeight("bold").setBackground("#f2f2f2");
  
  return sheet;
}


// ---------------- SKU MATCHING EXECUTION (SINGLE RUN) ----------------
function runSkuMatching() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Run Accurate SKU Matching',
    'This will match all manufacturer SKUs to platform SKUs using accuracy-focused algorithms. This may take some time for large datasets. Previous results will be cleared. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const matchingSheet = SS.getSheetByName('SKU Matching Engine');
  if (!matchingSheet) {
    ui.alert('Error', 'SKU Matching Engine tab not found. Run Safe Setup System first.', ui.ButtonSet.OK);
    return;
  }

  // Clear previous results
  const lastRowContent = matchingSheet.getLastRow();
  if (lastRowContent >= 3) {
    matchingSheet.getRange(3, 1, lastRowContent - 2, matchingSheet.getLastColumn()).clearContent();
  }
   if (matchingSheet.getMaxRows() > 2) { // Clear any lingering notes or validations in data area
    matchingSheet.getRange(3, 1, matchingSheet.getMaxRows() - 2, matchingSheet.getLastColumn()).clearDataValidations().clearNote();
   }


  const statusCell = matchingSheet.getRange(2, 1, 1, matchingSheet.getLastColumn());
  statusCell.merge().setValue('PROCESSING - Running SKU matching...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center');
  SpreadsheetApp.flush();

  try {
    const startTime = new Date().getTime();
    Object.keys(skuNormalizeCache).forEach(key => delete skuNormalizeCache[key]); // Clear normalization cache

    statusCell.setValue('PROCESSING - Loading manufacturer data...'); SpreadsheetApp.flush();
    const manufacturerData = getManufacturerData();
    if (!manufacturerData.length) {
      statusCell.setValue('ERROR - No manufacturer data found in "Manufacturer Price Sheet".').setBackground('#f4cccc');
      return;
    }

    const mfrTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`PROCESSING - Loaded ${manufacturerData.length} MFR SKUs in ${mfrTime.toFixed(1)}s. Pre-processing platform data...`); SpreadsheetApp.flush();
    
    const rawPlatformData = getPlatformDataFromStructuredSheet();
    if (!Object.keys(rawPlatformData).some(p => rawPlatformData[p].length > 0)) {
      statusCell.setValue('ERROR - No platform data found in "Platform Databases" sheet.').setBackground('#f4cccc');
      return;
    }
    const { preProcessedPlatformDataWithAttributes, platformSkuMaps, cleanSkuMaps } = preProcessAllPlatformData(rawPlatformData);
    
    let totalPlatformSkus = 0; 
    for (const platform in preProcessedPlatformDataWithAttributes) {
        totalPlatformSkus += preProcessedPlatformDataWithAttributes[platform].length;
    }

    const dataLoadTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`PROCESSING - Loaded & pre-processed ${totalPlatformSkus} platform SKUs in ${dataLoadTime.toFixed(1)}s (total). Starting matching...`); SpreadsheetApp.flush();

    const matchResults = performFullMatching(manufacturerData, preProcessedPlatformDataWithAttributes, platformSkuMaps, cleanSkuMaps);

    let totalMatches = 0, exactMatches = 0, highConfidenceMatches = 0, mediumConfidenceMatches = 0, lowConfidenceMatches = 0, reviewRequiredMatches = 0;
    matchResults.forEach(result => {
      Object.values(result.matches).forEach(match => {
        if (match) {
          totalMatches++;
          const confidence = match.confidenceScore;
          if (confidence >= 95) exactMatches++;
          else if (confidence >= 85) highConfidenceMatches++;
          else if (confidence >= 70) mediumConfidenceMatches++;
          else lowConfidenceMatches++;
          if (match.matchType && (match.matchType.includes('REVIEW') || (confidence < 85 && confidence >= 70))) reviewRequiredMatches++;
        }
      });
    });

    const matchingTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`PROCESSING - Found ${totalMatches} potential matches in ${matchingTime.toFixed(1)}s. Updating sheet...`); SpreadsheetApp.flush();
    
    // For a single run, outputMatchResults_append with isFirstBatch = true will clear and write
    appendMatchResultsToSheet(matchResults, matchingSheet, true); 
    
    const finalRowCount = matchingSheet.getLastRow() - 2; // Recalculate after writing
    if (finalRowCount > 0) {
        addConditionalFormattingToMatchingSheet(matchingSheet, finalRowCount);
    }


    const totalTime = (new Date().getTime() - startTime) / 1000;
    statusCell.setValue(`MATCHING COMPLETE - ${matchResults.length} MFR SKUs processed in ${totalTime.toFixed(1)}s. ${totalMatches} matches. ${reviewRequiredMatches} for review.`).setBackground('#d9ead3');
    ui.alert('Accurate SKU Matching Complete',
             `${matchResults.length} MFR SKUs processed in ${totalTime.toFixed(1)}s.\n` +
             `Total matches: ${totalMatches}\n` +
             `Exact (95-100%): ${exactMatches}\nHigh (85-94%): ${highConfidenceMatches}\n` +
             `Medium (70-84%): ${mediumConfidenceMatches}\nLow (<70%): ${lowConfidenceMatches}\n` +
             `${reviewRequiredMatches} matches require review.`, ui.ButtonSet.OK);
  } catch (error) {
    logError(error, 'runSkuMatching', ui);
    statusCell.setValue('ERROR - ' + error.toString()).setBackground('#f4cccc');
  }
}


// ---------------- DATA FETCHING FUNCTIONS ----------------
function getManufacturerData() { const sheet = SS.getSheetByName('Manufacturer Price Sheet'); if (!sheet) throw new Error('Manufacturer Price Sheet tab not found'); const values = sheet.getDataRange().getValues(); if (values.length <= 1) return []; const headers = values[0]; const skuIndex = headers.indexOf('Manufacturer SKU'); const upcIndex = headers.indexOf('UPC'); const msrpIndex = headers.indexOf('MSRP'); const mapIndex = headers.indexOf('MAP'); const dealerPriceIndex = headers.indexOf('Dealer Price'); if (skuIndex === -1) throw new Error('Manufacturer SKU column not found in MFR Sheet'); const data = []; for (let i = 1; i < values.length; i++) { const row = values[i]; const sku = row[skuIndex]; if (!sku) continue; let dealerPrice = dealerPriceIndex >= 0 ? row[dealerPriceIndex] : null; let map = mapIndex >= 0 ? row[mapIndex] : null; if ((!map || map === 0 || map === "") && dealerPrice && dealerPrice !== 0 && dealerPrice !== "") map = parseFloat(dealerPrice) / 0.85; data.push({ manufacturerSku: sku, upc: upcIndex >= 0 ? row[upcIndex] : "", msrp: msrpIndex >= 0 ? row[msrpIndex] : null, map: map, dealerPrice: dealerPrice }); } return data; }
function getPlatformDataFromStructuredSheet() { const sheet = SS.getSheetByName('Platform Databases'); if (!sheet) throw new Error('Platform Databases tab not found. Please create it or ensure the name is exact.'); const values = sheet.getDataRange().getValues(); if (values.length <= 1) return {}; const platformData = { AMAZON: [], EBAY: [], INFLOW: [], REVERB: [], SHOPIFY: [], SELLERCLOUD: [] }; const platformHeaderIndices = {}; values[0].forEach((header, col) => { const upperHeader = header.toUpperCase(); if (upperHeader.includes('AMAZON')) platformHeaderIndices.AMAZON = col; else if (upperHeader.includes('EBAY')) platformHeaderIndices.EBAY = col; else if (upperHeader.includes('INFLOW')) platformHeaderIndices.INFLOW = col; else if (upperHeader.includes('REVERB')) platformHeaderIndices.REVERB = col; else if (upperHeader.includes('SHOPIFY')) platformHeaderIndices.SHOPIFY = col; else if (upperHeader.includes('SELLERCLOUD')) platformHeaderIndices.SELLERCLOUD = col; }); const secondaryHeaders = values[1] || []; for (const platform in platformHeaderIndices) { const startCol = platformHeaderIndices[platform]; if (startCol === undefined) { Logger.log(`Warning: Platform ${platform} main header not found in 'Platform Databases' row 1.`); continue; } let skuColOffset = -1, priceColOffset = -1, costColOffset = -1, cleanSkuColOffset = -1, conditionColOffset = -1; for (let i = startCol; i < secondaryHeaders.length; i++) { const currentCellPlatformHeader = values[0][i]; if (i > startCol && currentCellPlatformHeader && currentCellPlatformHeader.trim() !== "" && platformHeaderIndices[platform] !== i && Object.values(platformHeaderIndices).includes(i)) break; const colName = secondaryHeaders[i]; const offset = i - startCol; switch (platform) { case 'AMAZON': if (colName === 'seller-sku') skuColOffset = offset; if (colName === 'price') priceColOffset = offset; if (colName === 'Clean Sku') cleanSkuColOffset = offset; break; case 'EBAY': if (colName === 'Custom label (SKU)') skuColOffset = offset; if (colName === 'Start price') priceColOffset = offset; break; case 'INFLOW': if (colName === 'Name') skuColOffset = offset; if (colName === 'UnitPrice') priceColOffset = offset; if (colName === 'Cost') costColOffset = offset; if (colName === 'Clean Sku') cleanSkuColOffset = offset; break; case 'REVERB': if (colName === 'sku') skuColOffset = offset; if (colName === 'price') priceColOffset = offset; if (colName === 'condition') conditionColOffset = offset; break; case 'SHOPIFY': if (colName === 'Variant SKU') skuColOffset = offset; if (colName === 'Variant Price') priceColOffset = offset; if (colName === 'Variant Cost') costColOffset = offset; break; case 'SELLERCLOUD': if (colName === 'ProductID') skuColOffset = offset; if (colName === 'SitePrice') priceColOffset = offset; if (colName === 'SiteCost') costColOffset = offset; break; } } if (skuColOffset === -1) { Logger.log(`Warning: SKU column not found for platform ${platform} under its designated columns in 'Platform Databases' row 2.`); continue; } for (let rowIdx = 2; rowIdx < values.length; rowIdx++) { const sku = values[rowIdx][startCol + skuColOffset]; if (!sku) continue; platformData[platform].push({ platform: platform, sku: sku, price: priceColOffset !== -1 ? values[rowIdx][startCol + priceColOffset] : null, cost: costColOffset !== -1 ? values[rowIdx][startCol + costColOffset] : null, cleanSku: cleanSkuColOffset !== -1 ? values[rowIdx][startCol + cleanSkuColOffset] : null, condition: conditionColOffset !== -1 ? values[rowIdx][startCol + conditionColOffset] : null }); } } return platformData; }


// ---------------- OUTPUT & FORMATTING ---------------- 
function appendMatchResultsToSheet(batchMatchResults, sheet) {
  if (!batchMatchResults || batchMatchResults.length === 0) {
    Logger.log("No results in this batch to append.");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowsToAppend = [];

  batchMatchResults.forEach(result => {
    const row = Array(headers.length).fill('');
    headers.forEach((header, colIndex) => {
      if (header === 'Manufacturer SKU') row[colIndex] = result.manufacturerSku;
      else if (header === 'MSRP') row[colIndex] = result.msrp;
      else if (header === 'MAP') row[colIndex] = result.map;
      else if (header === 'Dealer Price') row[colIndex] = result.dealerPrice;
      else if (header === 'Status') {
        const hasAutoMatch = Object.values(result.matches).some(match => match && match.confidenceScore >= 85);
        const needsReview = Object.values(result.matches).some(match => match && match.matchType && (match.matchType.includes('REVIEW') || (match.confidenceScore < 85 && match.confidenceScore >= 70)));
        const hasAnyMatch = Object.values(result.matches).some(match => match);

        if (needsReview) row[colIndex] = 'Review Required';
        else if (hasAutoMatch) row[colIndex] = 'Auto-matched';
        else if (hasAnyMatch) row[colIndex] = 'Partial Matches';
        else row[colIndex] = 'No Match';
      }
      else if (header === 'Notes') row[colIndex] = ''; 
      else {
        const parts = header.split(' ');
        if (parts.length >= 2) {
          const platform = parts[0]; const field = parts.slice(1).join(' ');
          const match = result.matches[platform];
          if (match) {
            if (field === 'SKU') row[colIndex] = match.platformSku;
            else if (field === 'Confidence') row[colIndex] = match.confidenceScore;
            else if (field === 'Current Price') row[colIndex] = match.currentPrice;
          }
        }
      }
    });
    rowsToAppend.push(row);
  });

  if (rowsToAppend.length > 0) {
    const startRowForAppend = Math.max(3, sheet.getLastRow() + 1);
    sheet.getRange(startRowForAppend, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    Logger.log(`Appended ${rowsToAppend.length} rows to SKU Matching Engine.`);
  }
}

function addConditionalFormattingToMatchingSheet(sheet, rowCount) {
  if (rowCount <= 0) return;
  sheet.clearConditionalFormatRules();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rules = [];
  const dataStartRow = 3;

  headers.forEach((header, index) => {
    if (header.includes('Confidence')) {
      const range = sheet.getRange(dataStartRow, index + 1, rowCount, 1);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(95).setBackground('#d9ead3').setRanges([range]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(85, 94.99).setBackground('#cfe2f3').setRanges([range]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(70, 84.99).setBackground('#fff2cc').setRanges([range]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).whenNumberLessThan(70).setBackground('#f4cccc').setRanges([range]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#efefef').setRanges([range]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#ffffff').setRanges([range]).build());
    }
  });

  const statusColIndex = headers.indexOf('Status');
  if (statusColIndex >= 0) {
    const range = sheet.getRange(dataStartRow, statusColIndex + 1, rowCount, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Review Required').setBackground('#f4cccc').setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Partial Matches').setBackground('#fff2cc').setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Auto-matched').setBackground('#d9ead3').setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('No Match').setBackground('#efefef').setRanges([range]).build());
  }
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
    Logger.log("Applied conditional formatting to SKU Matching Engine.");
  }
}

// ---------------- MATCH REVIEW & UPDATE (SIDEBAR) - REMOVED ----------------
// The functions showMatchReviewSidebar, getMatchesForReviewAccurate, getConservativeSkuSuggestions_optimized, and updateMatch
// have been removed as this feature is no longer needed. Manual review will be done directly on the "SKU Matching Engine" sheet.

// ---------------- OTHER CORE FUNCTIONS ----------------
function validateMatchQuality() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Validate Match Quality', 'Analyze existing matches and flag potentially incorrect ones. Continue?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    if (!matchingSheet) { ui.alert('Error', 'SKU Matching Engine tab not found.', ui.ButtonSet.OK); return; }

    const matchingData = matchingSheet.getDataRange().getValues();
    const headers = matchingData[0];
    const mfrSkuIndex = headers.indexOf('Manufacturer SKU');
    const statusIndex = headers.indexOf('Status');
    const notesIndex = headers.indexOf('Notes');
    if (mfrSkuIndex === -1) { ui.alert('Error', 'Manufacturer SKU column not found.'); return; }

    const platformColumns = {}; const platformConfidenceColumns = {};
    headers.forEach((header, index) => {
      if (header.includes(' SKU') && !header.includes('Manufacturer')) platformColumns[header.split(' ')[0]] = index;
      if (header.includes(' Confidence')) platformConfidenceColumns[header.split(' ')[0]] = index;
    });

    let totalMatches = 0, suspiciousMatches = 0, goodMatches = 0;
    const suspiciousRowsInfo = [];

    for (let i = 2; i < matchingData.length; i++) {
      const row = matchingData[i];
      const mfrSkuRaw = row[mfrSkuIndex];
      if (!mfrSkuRaw) continue;
      const mfrAttributes = extractSkuAttributesAndCore(mfrSkuRaw);

      for (const platform in platformColumns) {
        const platformSkuRaw = row[platformColumns[platform]];
        const currentConfidence = platformConfidenceColumns[platform] !== undefined ? parseFloat(row[platformConfidenceColumns[platform]]) : 0;
        if (!platformSkuRaw || currentConfidence < 1) continue;

        totalMatches++;
        const platformAttributes = extractSkuAttributesAndCore(platformSkuRaw);
        const validation = validateMatch_optimized(mfrSkuRaw, mfrAttributes, platformSkuRaw, platformAttributes, platform);

        if (validation.confidence < 85 && currentConfidence >= 85) {
          suspiciousMatches++;
          suspiciousRowsInfo.push({
            rowNum: i + 1, platform: platform, mfrSku: mfrSkuRaw, platformSku: platformSkuRaw,
            originalConfidence: currentConfidence, newConfidence: validation.confidence, reason: validation.reason
          });
        } else if (validation.confidence >= 85) {
          goodMatches++;
        }
      }
    }

    suspiciousRowsInfo.forEach(info => {
      const currentNotes = matchingSheet.getRange(info.rowNum, notesIndex + 1).getValue().toString();
      const warningNote = `QUALITY CHECK: ${info.platform} match (${info.platformSku}) re-validated to ${info.newConfidence}% (was ${info.originalConfidence}%). Reason: ${info.reason}`;
      if (!currentNotes.includes('QUALITY CHECK: ' + info.platform)) {
        matchingSheet.getRange(info.rowNum, notesIndex + 1).setValue((currentNotes ? currentNotes + '; ' : '') + warningNote);
        if (statusIndex !== -1) matchingSheet.getRange(info.rowNum, statusIndex + 1).setValue('Quality Review Required');
      }
    });

    if (suspiciousRowsInfo.length > 0) applyQualityValidationFormatting(matchingSheet);
    let message = `Match Quality Validation Complete\nTotal: ${totalMatches}, Good: ${goodMatches}, Suspicious: ${suspiciousMatches}\n`;
    if (suspiciousMatches > 0) message += `Flagged matches updated in Notes. Please review.`;
    else message += `All existing high-confidence matches seem consistent with current logic.`;
    ui.alert('Validation Complete', message, ui.ButtonSet.OK);

  } catch (error) { logError(error, 'validateMatchQuality', ui); }
}

function applyQualityValidationFormatting(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const notesIndex = headers.indexOf('Notes');
  const statusIndex = headers.indexOf('Status');
  const dataStartRow = 3;
  const numDataRows = sheet.getLastRow() - dataStartRow + 1;
  if (numDataRows <= 0) return;

  const existingRules = sheet.getConditionalFormatRules();
  const newRules = [];
  if (notesIndex >= 0) {
    const notesRange = sheet.getRange(dataStartRow, notesIndex + 1, numDataRows, 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('QUALITY CHECK').setBackground('#fff2cc').setRanges([notesRange]).build());
  }
  if (statusIndex >= 0) {
    const statusRange = sheet.getRange(dataStartRow, statusIndex + 1, numDataRows, 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Quality Review').setBackground('#f4cccc').setRanges([statusRange]).build());
  }
  sheet.setConditionalFormatRules(existingRules.concat(newRules));
}

function saveMatches() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Save Matches', 'Save current matches (70%+) to Match History? Continue?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  try {
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    const historySheet = SS.getSheetByName('Match History');
    if (!matchingSheet || !historySheet) { ui.alert('Error', 'Required sheets not found.', ui.ButtonSet.OK); return; }

    const matchingData = matchingSheet.getDataRange().getValues();
    const matchHeaders = matchingData[0];
    const mfrSkuIndex = matchHeaders.indexOf('Manufacturer SKU');
    const statusIndex = matchHeaders.indexOf('Status');
    if (mfrSkuIndex === -1) { ui.alert('Error', 'MFR SKU column not found in Matching Engine.', ui.ButtonSet.OK); return; }

    const platformMeta = {};
     matchHeaders.forEach((header, index) => {
      const parts = header.split(' ');
      if (parts.length > 1) {
        const platform = parts[0];
        if (PLATFORM_TABS[platform]) {
          if (!platformMeta[platform]) platformMeta[platform] = {};
          if (parts[1] === 'SKU') platformMeta[platform].skuIndex = index;
          else if (parts[1] === 'Confidence') platformMeta[platform].confidenceIndex = index;
        }
      }
    });

    const historyHeaders = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
    const histMfrSkuIdx = historyHeaders.indexOf('Manufacturer SKU'); const histPlatformIdx = historyHeaders.indexOf('Platform');
    const histPlatSkuIdx = historyHeaders.indexOf('Platform SKU'); const histMatchTypeIdx = historyHeaders.indexOf('Match Type');
    const histConfidenceIdx = historyHeaders.indexOf('Confidence'); const histDateIdx = historyHeaders.indexOf('Date Added');
    
    const existingHistory = {};
    const historyData = historySheet.getDataRange().getValues();
    for (let i = 1; i < historyData.length; i++) { existingHistory[`${historyData[i][histMfrSkuIdx]}_${historyData[i][histPlatformIdx]}`] = true; }

    const rowsToAdd = []; const currentDate = new Date();
    for (let i = 2; i < matchingData.length; i++) {
      const row = matchingData[i]; const mfrSku = row[mfrSkuIndex]; if (!mfrSku) continue;
      const status = statusIndex >=0 ? row[statusIndex] : '';

      for (const platformKey in platformMeta) {
        const meta = platformMeta[platformKey];
        const platformSku = meta.skuIndex !== undefined ? row[meta.skuIndex] : '';
        if (!platformSku || existingHistory[`${mfrSku}_${platformKey}`]) continue;
        const confidence = meta.confidenceIndex !== undefined ? row[meta.confidenceIndex] : '';
        if (parseFloat(confidence) < 70) continue;

        let matchType = 'Auto';
        if (status.includes('Manual')) matchType = 'Manual';
        else if (confidence >= 95) matchType = 'Exact';
        else if (confidence >= 85) matchType = 'High Confidence';
        else if (confidence >= 70) matchType = 'Medium Confidence';
        
        const historyRow = Array(historyHeaders.length).fill('');
        historyRow[histMfrSkuIdx] = mfrSku; historyRow[histPlatformIdx] = platformKey; historyRow[histPlatSkuIdx] = platformSku;
        historyRow[histMatchTypeIdx] = matchType; historyRow[histConfidenceIdx] = confidence; historyRow[histDateIdx] = currentDate;
        rowsToAdd.push(historyRow);
      }
    }
    if (rowsToAdd.length > 0) historySheet.getRange(historySheet.getLastRow() + 1, 1, rowsToAdd.length, historyHeaders.length).setValues(rowsToAdd);
    ui.alert('Success', `${rowsToAdd.length} new matches (70%+) saved to Match History.`, ui.ButtonSet.OK);
  } catch (error) { logError(error, 'saveMatches', ui); }
}

function updatePriceAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Update Price Analysis', 'This will analyze price changes and update the "Price Analysis Dashboard" and the new "Price Changes" tab. Continue?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let dashboardSheet;
  let priceChangesSheet;
  try {
    const matchingSheet = SS.getSheetByName('SKU Matching Engine');
    dashboardSheet = SS.getSheetByName('Price Analysis Dashboard');
    priceChangesSheet = SS.getSheetByName('Price Changes'); // Get the new sheet
    const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet');

    if (!matchingSheet || !dashboardSheet || !mfrSheet) { ui.alert('Error', 'Required sheets (SKU Matching Engine, Price Analysis Dashboard, Manufacturer Price Sheet) not found.', ui.ButtonSet.OK); return; }
    if (!priceChangesSheet) {
        priceChangesSheet = createPriceChangesTab(); // Create if it doesn't exist
        if (!priceChangesSheet) { // Check again
             ui.alert('Error', 'Could not create or find "Price Changes" sheet.', ui.ButtonSet.OK); return;
        }
    }


    dashboardSheet.getRange(1, 1, 1, dashboardSheet.getLastColumn()).merge().setValue('PROCESSING - Analyzing prices...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center');
    priceChangesSheet.getRange("A1").setValue("PROCESSING - Updating Price Change Lists...").setFontWeight("bold").setBackground(COLORS.WARNING); // Status for new sheet
    SpreadsheetApp.flush();

    const matchingData = matchingSheet.getDataRange().getValues();
    const matchingHeaders = matchingData[0];
    const mfrSkuIndex = matchingHeaders.indexOf('Manufacturer SKU');
    const mapIndex = matchingHeaders.indexOf('MAP');
    const dealerPriceIndex = matchingHeaders.indexOf('Dealer Price');

    const platformMeta = {};
    matchingHeaders.forEach((header, index) => {
        const parts = header.split(' ');
        if (parts.length > 1) {
            const platform = parts[0];
            if (ANALYSIS_PLATFORMS.includes(platform) && PLATFORM_TABS[platform]) {
                if (!platformMeta[platform]) platformMeta[platform] = {};
                if (parts[1] === 'SKU') platformMeta[platform].skuIndex = index;
                else if (parts.slice(1).join(' ') === 'Current Price') platformMeta[platform].priceIndex = index;
                else if (parts[1] === 'Confidence') platformMeta[platform].confidenceIndex = index;
            }
        }
    });
    resetDashboardLayout(dashboardSheet);
    const dashboardHeaders = dashboardSheet.getRange(6, 1, 1, dashboardSheet.getLastColumn()).getValues()[0];
    const dashIndices = { sku: dashboardHeaders.indexOf('SKU'), map: dashboardHeaders.indexOf('MAP'), dealer: dashboardHeaders.indexOf('Dealer'), bStock: dashboardHeaders.indexOf('B-Stock'), platform: dashboardHeaders.indexOf('Platform'), current: dashboardHeaders.indexOf('Current'), newPrice: dashboardHeaders.indexOf('New'), changeAmt: dashboardHeaders.indexOf('Change $'), changePct: dashboardHeaders.indexOf('Change %'), status: dashboardHeaders.indexOf('Status') };

    const analysisData = [];
    const priceUpItems = [];
    const priceDownItems = [];

    let totalProducts = 0, priceIncreases = 0, priceDecreases = 0, mapViolations = 0, highImpactChanges = 0, unmatchedProductsThisRun = 0, totalPercentChange = 0, validProductCountForAvg = 0, bStockChanges = 0;

    for (let i = 2; i < matchingData.length; i++) {
      const row = matchingData[i];
      const mfrSku = row[mfrSkuIndex];
      if (!mfrSku) continue;
      totalProducts++;

      const newMapPrice = parseFloat(row[mapIndex]);
      const newDealerPrice = parseFloat(row[dealerPriceIndex]);
      const bStockInfo = getBStockInfo(mfrSku);
      const isBStock = bStockInfo !== null;

      let platformProcessedForMfrSku = false;
      for (const platformKey in platformMeta) {
        const meta = platformMeta[platformKey];
        const platformSku = meta.skuIndex !== undefined ? row[meta.skuIndex] : '';
        const currentPlatformPrice = meta.priceIndex !== undefined && row[meta.priceIndex] !== '' ? parseFloat(row[meta.priceIndex]) : NaN;
        const confidence = meta.confidenceIndex !== undefined && row[meta.confidenceIndex] !== '' ? parseFloat(row[meta.confidenceIndex]) : 0;

        if (!platformSku || confidence < 85) {
            if (!platformProcessedForMfrSku) {
                let anyPlatformMatched = false;
                for (const pkCheck in platformMeta) {
                    const m = platformMeta[pkCheck];
                    if (row[m.skuIndex] && parseFloat(row[m.confidenceIndex]) >= 85) {
                        anyPlatformMatched = true;
                        break;
                    }
                }
                if (!anyPlatformMatched) unmatchedProductsThisRun++;
                platformProcessedForMfrSku = true;
            }
            continue;
        }
        platformProcessedForMfrSku = true;

        let calculatedNewPrice;
        let priceStatus = '';
        if (isBStock) {
          calculatedNewPrice = newMapPrice * bStockInfo.multiplier;
          priceStatus = `${bStockInfo.type} ${bStockInfo.isSpecial ? "Special" : "B-Stock"}`;
        } else {
          calculatedNewPrice = newMapPrice;
        }
        calculatedNewPrice = Math.round(calculatedNewPrice * 100) / 100;

        let changeAmount = 0;
        let changePercent = 0;

        if (!isNaN(currentPlatformPrice) && !isNaN(calculatedNewPrice)) {
          changeAmount = calculatedNewPrice - currentPlatformPrice;
          changePercent = currentPlatformPrice !== 0 ? (changeAmount / currentPlatformPrice) * 100 : (calculatedNewPrice > 0 ? Infinity : 0);

          if (changeAmount > 0.001) {
            priceIncreases++;
            priceUpItems.push([mfrSku, platformKey, currentPlatformPrice, calculatedNewPrice, changeAmount, isFinite(changePercent) ? changePercent / 100 : 'N/A']);
          } else if (changeAmount < -0.001) {
            priceDecreases++;
            priceDownItems.push([mfrSku, platformKey, currentPlatformPrice, calculatedNewPrice, changeAmount, isFinite(changePercent) ? changePercent / 100 : 'N/A']);
          }

          if (Math.abs(changePercent) > 10 || Math.abs(changeAmount) > 10) { highImpactChanges++; priceStatus += (priceStatus ? '; ' : '') + 'High Impact'; }
          if (!isBStock && newMapPrice > 0 && currentPlatformPrice < newMapPrice - 0.001) { mapViolations++; priceStatus += (priceStatus ? '; ' : '') + 'MAP Violation'; }
          if (isBStock && Math.abs(changeAmount) > 0.01) bStockChanges++;
          if (isFinite(changePercent)) { totalPercentChange += changePercent; validProductCountForAvg++; }
        } else if (isNaN(currentPlatformPrice) && !isNaN(calculatedNewPrice)){
            priceStatus += (priceStatus ? '; ' : '') + 'New Listing/Price';
            priceIncreases++; 
            priceUpItems.push([mfrSku, platformKey, 'N/A', calculatedNewPrice, calculatedNewPrice, 'New Item']);
        }

        const analysisRow = Array(dashboardHeaders.length).fill('');
        if (dashIndices.sku !== -1) analysisRow[dashIndices.sku] = mfrSku;
        if (dashIndices.map !== -1) analysisRow[dashIndices.map] = isNaN(newMapPrice) ? '' : newMapPrice;
        if (dashIndices.dealer !== -1) analysisRow[dashIndices.dealer] = isNaN(newDealerPrice) ? '' : newDealerPrice;
        if (dashIndices.bStock !== -1) analysisRow[dashIndices.bStock] = isBStock ? bStockInfo.type : '';
        if (dashIndices.platform !== -1) analysisRow[dashIndices.platform] = platformKey;
        if (dashIndices.current !== -1) analysisRow[dashIndices.current] = isNaN(currentPlatformPrice) ? '' : currentPlatformPrice;
        if (dashIndices.newPrice !== -1) analysisRow[dashIndices.newPrice] = isNaN(calculatedNewPrice) ? '' : calculatedNewPrice;
        if (dashIndices.changeAmt !== -1) analysisRow[dashIndices.changeAmt] = changeAmount;
        if (dashIndices.changePct !== -1) analysisRow[dashIndices.changePct] = isFinite(changePercent) ? changePercent / 100 : (changePercent === Infinity ? 'New Item' : '');
        if (dashIndices.status !== -1) analysisRow[dashIndices.status] = priceStatus;
        analysisData.push(analysisRow);
      }
    }

    if (analysisData.length > 0) {
      const dataRange = dashboardSheet.getRange(7, 1, analysisData.length, dashboardHeaders.length);
      dataRange.setValues(analysisData); dataRange.clearFormat();
      const formatColumns = [ { header: 'MAP', format: '$#,##0.00' }, { header: 'Dealer', format: '$#,##0.00' }, { header: 'Current', format: '$#,##0.00' }, { header: 'New', format: '$#,##0.00' }, { header: 'Change $', format: '$#,##0.00' }, { header: 'Change %', format: '0.00%' }];
      formatColumns.forEach(colInfo => { const colIndex = dashboardHeaders.indexOf(colInfo.header); if (colIndex !== -1) dashboardSheet.getRange(7, colIndex + 1, analysisData.length, 1).setNumberFormat(colInfo.format); });
    }
    dashboardSheet.getRange(3, 2).setValue(totalProducts); dashboardSheet.getRange(3, 5).setValue(priceIncreases); dashboardSheet.getRange(3, 8).setValue(priceDecreases); dashboardSheet.getRange(3, 11).setValue(validProductCountForAvg > 0 ? (totalPercentChange / validProductCountForAvg / 100) : 0).setNumberFormat('0.0%');
    dashboardSheet.getRange(4, 2).setValue(mapViolations); dashboardSheet.getRange(4, 5).setValue(highImpactChanges); dashboardSheet.getRange(4, 8).setValue(unmatchedProductsThisRun); dashboardSheet.getRange(4, 11).setValue(bStockChanges);
    applyDashboardFormatting(dashboardSheet, analysisData.length);
    dashboardSheet.getRange(1, 1, 1, dashboardSheet.getLastColumn()).merge().setValue('PRICE ANALYSIS DASHBOARD - COMPLETE - Last Updated: ' + new Date().toLocaleString()).setBackground('#e0e0e0');

    // Populate Price Changes Tab
    populatePriceChangesTab(priceChangesSheet, priceUpItems, priceDownItems);

    ui.alert('Analysis Complete', `Price analysis updated. Products analyzed: ${totalProducts}. Price changes identified: ${priceIncreases + priceDecreases}.`, ui.ButtonSet.OK);

  } catch (error) {
    logError(error, 'updatePriceAnalysis', ui);
    if (dashboardSheet) dashboardSheet.getRange(1, 1, 1, dashboardSheet.getLastColumn()).merge().setValue('ERROR - ' + error.toString()).setBackground('#f4cccc');
    if (priceChangesSheet) priceChangesSheet.getRange("A1").setValue("ERROR during price analysis. Check logs.").setBackground(COLORS.NEGATIVE);
  }
}

function populatePriceChangesTab(sheet, priceUpItems, priceDownItems) {
  sheet.clearContents(); // Clear previous content
  sheet.getRange("A1").setValue("PRICE INCREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_UP);
  const headers = ["MFR SKU", "Platform", "Old Price", "New Price", "Change $", "Change %"];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f2f2f2"); // Corrected row from A2 to 2

  let currentRow = 3;
  if (priceUpItems.length > 0) {
    sheet.getRange(currentRow, 1, priceUpItems.length, headers.length).setValues(priceUpItems);
    sheet.getRange(currentRow, 3, priceUpItems.length, 3).setNumberFormat("$#,##0.00"); // Old, New, Change $
    sheet.getRange(currentRow, 6, priceUpItems.length, 1).setNumberFormat("0.00%"); // Change %
    currentRow += priceUpItems.length;
  } else {
    sheet.getRange(currentRow, 1).setValue("No high-confidence price increases found.").setFontStyle("italic");
    currentRow++;
  }

  currentRow += 2; // Add some space
  sheet.getRange(currentRow, 1).setValue("PRICE DECREASES (High Confidence Only - 85%+)").setFontWeight("bold").setBackground(COLORS.PRICE_DOWN);
  currentRow++;
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f2f2f2");
  currentRow++;

  if (priceDownItems.length > 0) {
    sheet.getRange(currentRow, 1, priceDownItems.length, headers.length).setValues(priceDownItems);
    sheet.getRange(currentRow, 3, priceDownItems.length, 3).setNumberFormat("$#,##0.00"); // Old, New, Change $
    sheet.getRange(currentRow, 6, priceDownItems.length, 1).setNumberFormat("0.00%"); // Change %
  } else {
    sheet.getRange(currentRow, 1).setValue("No high-confidence price decreases found.").setFontStyle("italic");
  }
  sheet.autoResizeColumns(1, headers.length);
  Logger.log("Price Changes tab populated.");
}


function resetDashboardLayout(sheet) { 
  if (sheet.getLastRow() >= 7) {
    sheet.getRange(7, 1, sheet.getLastRow() - 6, sheet.getMaxColumns()).clear({contentsOnly: true, formatOnly: true, validationsOnly: true, commentsOnly: true});
  }
  sheet.clearConditionalFormatRules();
  sheet.getRange(2, 1, 3, sheet.getMaxColumns()).clear({contentsOnly: true, formatOnly: true}); 

  sheet.getRange(2, 1, 1, 12).merge().setValue('SUMMARY METRICS').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f3f3f3');
  const summaryMetrics = [ ['Total Products:', '0', '', 'Price Increases:', '0', '', 'Price Decreases:', '0', '', 'Average Change:', '0%'], ['MAP Violations:', '0', '', 'High Impact Changes:', '0', '', 'Unmatched Products:', '0', '', 'B-Stock Changes:', '0'] ];
  sheet.getRange(3, 1, 2, 11).setValues(summaryMetrics); 
  sheet.getRange(3, 1, 2, 1).setFontWeight('bold'); sheet.getRange(3, 4, 2, 1).setFontWeight('bold'); sheet.getRange(3, 7, 2, 1).setFontWeight('bold'); sheet.getRange(3, 10, 2, 1).setFontWeight('bold');
  const analysisHeaders = ['SKU', 'MAP', 'Dealer', 'B-Stock', 'Platform', 'Current', 'New', 'Change $', 'Change %', 'Status'];
  sheet.getRange(6, 1, 1, analysisHeaders.length).setValues([analysisHeaders]).setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(5, 1, 1, analysisHeaders.length).merge().setValue('Use "Update Price Analysis" to refresh data below.').setFontStyle('italic').setHorizontalAlignment('center').setBackground(null); 
}
function applyDashboardFormatting(sheet, rowCount) { if (rowCount <= 0) return; sheet.clearConditionalFormatRules(); const headers = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues()[0]; const dataStartRow = 7; const changeAmtColIndex = headers.indexOf('Change $') + 1; const changePctColIndex = headers.indexOf('Change %') + 1; const statusColIndex = headers.indexOf('Status') + 1; const rules = []; if (changeAmtColIndex > 0) { const range = sheet.getRange(dataStartRow, changeAmtColIndex, rowCount, 1); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(COLORS.NEGATIVE).setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(COLORS.POSITIVE).setRanges([range]).build()); } if (changePctColIndex > 0) { const range = sheet.getRange(dataStartRow, changePctColIndex, rowCount, 1); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.10).setBackground('#fce5cd').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.05, 0.10).setBackground('#fff2cc').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.10).setBackground('#d9ead3').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.10, -0.05).setBackground('#e2f0d9').setRanges([range]).build()); } if (statusColIndex > 0) { const range = sheet.getRange(dataStartRow, statusColIndex, rowCount, 1); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('High Impact').setBackground(COLORS.WARNING).setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('MAP Violation').setBackground(COLORS.NEGATIVE).setFontColor('#FFFFFF').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('B-Stock').setBackground(COLORS.INFLOW).setFontColor('#FFFFFF').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Special').setBackground(COLORS.SHOPIFY).setFontColor('#FFFFFF').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('New Listing/Price').setBackground('#cfe2f3').setRanges([range]).build()); } if (rules.length > 0) sheet.setConditionalFormatRules(rules); }
function generateExports() { const ui = SpreadsheetApp.getUi(); const EXPORT_CONFIDENCE_THRESHOLD = 85; const response = ui.alert('Generate Export Files', `Generate exports for platforms with ${EXPORT_CONFIDENCE_THRESHOLD}%+ confidence matches? Continue?`, ui.ButtonSet.YES_NO); if (response !== ui.Button.YES) return; let dashboardSheet; try { dashboardSheet = SS.getSheetByName('Price Analysis Dashboard'); const matchingSheet = SS.getSheetByName('SKU Matching Engine'); const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet'); if (!matchingSheet || !mfrSheet || !dashboardSheet) { ui.alert('Error', 'Required sheets not found.', ui.ButtonSet.OK); return; } const titleCell = dashboardSheet.getRange(1, 1); const titleText = titleCell.getValue().toString(); if (!titleText.includes('COMPLETE')) { ui.alert('Price Analysis Required', 'Run "Update Price Analysis" first.', ui.ButtonSet.OK); return; } const analysisData = dashboardSheet.getRange(7, 1, Math.max(1, dashboardSheet.getLastRow()-6), dashboardSheet.getLastColumn()).getValues(); const analysisHeaders = dashboardSheet.getRange(6, 1, 1, dashboardSheet.getLastColumn()).getValues()[0]; const anSkuIdx = analysisHeaders.indexOf('SKU'); const anPlatformIdx = analysisHeaders.indexOf('Platform'); const anNewPriceIdx = analysisHeaders.indexOf('New'); const anMapIdx = analysisHeaders.indexOf('MAP'); const anDealerIdx = analysisHeaders.indexOf('Dealer'); const anBstockTypeIdx = analysisHeaders.indexOf('B-Stock'); const matchingData = matchingSheet.getDataRange().getValues(); const matchingHeaders = matchingData[0]; const matchMfrSkuIdx = matchingHeaders.indexOf('Manufacturer SKU'); const mfrSkuToMatchRowIndex = {}; for(let i=2; i < matchingData.length; i++) if(matchingData[i][matchMfrSkuIdx]) mfrSkuToMatchRowIndex[matchingData[i][matchMfrSkuIdx]] = i; const mfrPriceSheetData = mfrSheet.getDataRange().getValues(); const mfrPriceSheetHeaders = mfrPriceSheetData[0]; const mfrSkuIndexUPC = mfrPriceSheetHeaders.indexOf('Manufacturer SKU'); const upcIndexUPC = mfrPriceSheetHeaders.indexOf('UPC'); const upcMap = {}; if (mfrSkuIndexUPC !== -1 && upcIndexUPC !== -1) for (let i = 1; i < mfrPriceSheetData.length; i++) if (mfrPriceSheetData[i][mfrSkuIndexUPC]) upcMap[mfrPriceSheetData[i][mfrSkuIndexUPC]] = mfrPriceSheetData[i][upcIndexUPC] || ''; for (const platformKey in PLATFORM_TABS) { const platformSkuColNameOnMatchingSheet = platformKey + ' SKU'; const platformConfidenceColNameOnMatchingSheet = platformKey + ' Confidence'; const platformSkuIndexOnMatching = matchingHeaders.indexOf(platformSkuColNameOnMatchingSheet); const confidenceIndexOnMatching = matchingHeaders.indexOf(platformConfidenceColNameOnMatchingSheet); if (platformSkuIndexOnMatching === -1) { Logger.log(`Skipping ${platformKey} - SKU column not found on Matching Sheet.`); continue; } const exportSheetName = PLATFORM_TABS[platformKey].name; let exportSheet = SS.getSheetByName(exportSheetName); if (!exportSheet) { createExportTab(platformKey); exportSheet = SS.getSheetByName(exportSheetName); if(!exportSheet) { Logger.log(`Error: Could not create/find export sheet for ${platformKey}`); continue; }} const exportHeaders = PLATFORM_TABS[platformKey].headers; if (exportSheet.getLastRow() > 3) exportSheet.getRange(4, 1, exportSheet.getLastRow() - 3, exportHeaders.length).clearContent(); const exportRows = []; for (const analysisRow of analysisData) { const mfrSku = analysisRow[anSkuIdx]; const platformAnalyzed = analysisRow[anPlatformIdx]; if (!mfrSku || platformAnalyzed !== platformKey) continue; const matchRowIndex = mfrSkuToMatchRowIndex[mfrSku]; if (matchRowIndex === undefined) continue; const platformSkuOnMatch = matchingData[matchRowIndex][platformSkuIndexOnMatching]; const confidence = confidenceIndexOnMatching !== -1 ? parseFloat(matchingData[matchRowIndex][confidenceIndexOnMatching]) : 0; if (!platformSkuOnMatch || confidence < EXPORT_CONFIDENCE_THRESHOLD) continue; const newPrice = parseFloat(analysisRow[anNewPriceIdx]); const mapPrice = parseFloat(analysisRow[anMapIdx]); const dealerPrice = parseFloat(analysisRow[anDealerIdx]); const bStockTypeFromAnalysis = analysisRow[anBstockTypeIdx]; const isBAsku = bStockTypeFromAnalysis && bStockTypeFromAnalysis !== ''; const bStockInfoForExport = isBAsku ? getBStockInfo(mfrSku) : null; const exportRowData = createPlatformExportRow(platformKey, exportHeaders, platformSkuOnMatch, newPrice, 0, mapPrice, dealerPrice, upcMap[mfrSku] || '', isBAsku, bStockInfoForExport); if (exportRowData) exportRows.push(exportRowData); } if (exportRows.length > 0) { exportSheet.getRange(4, 1, exportRows.length, exportHeaders.length).setValues(exportRows); exportSheet.getRange(2, 1, 1, exportHeaders.length).merge().setValue(`READY FOR EXPORT - ${exportRows.length} items (${EXPORT_CONFIDENCE_THRESHOLD}%+) - ${new Date().toLocaleString()}`).setBackground('#d9ead3').setFontColor('#000000'); } else exportSheet.getRange(2, 1, 1, exportHeaders.length).merge().setValue(`NO DATA (${EXPORT_CONFIDENCE_THRESHOLD}%+ matches not found for ${platformKey})`).setBackground('#f4cccc').setFontColor('#000000'); } ui.alert('Exports Generated', `Export files updated with ${EXPORT_CONFIDENCE_THRESHOLD}%+ confidence matches using prices from the Analysis Dashboard.`, ui.ButtonSet.OK); } catch (error) { logError(error, 'generateExports', ui); } }
function createPlatformExportRow(platform, headers, platformSku, newPrice, msrp, map, dealerPrice, upc, isBAsku, bStockInfo) { const row = Array(headers.length).fill(''); const formattedPrice = !isNaN(newPrice) ? Math.round(newPrice * 100) / 100 : ''; const formattedMap = !isNaN(map) ? Math.round(map * 100) / 100 : ''; const formattedDealer = !isNaN(dealerPrice) ? Math.round(dealerPrice * 100) / 100 : ''; const idx = (name) => headers.indexOf(name); switch (platform) { case 'AMAZON': if (idx('seller-sku') !== -1) row[idx('seller-sku')] = platformSku; if (idx('price') !== -1) row[idx('price')] = formattedPrice; break; case 'EBAY': if (idx('Action') !== -1) row[idx('Action')] = 'Revise'; if (idx('Custom label (SKU)') !== -1) row[idx('Custom label (SKU)')] = platformSku; if (idx('Start price') !== -1) row[idx('Start price')] = formattedPrice; break; case 'SHOPIFY': if (idx('Variant SKU') !== -1) row[idx('Variant SKU')] = platformSku; if (idx('Variant Price') !== -1) row[idx('Variant Price')] = formattedPrice; if (idx('Variant Compare At Price') !== -1) row[idx('Variant Compare At Price')] = isBAsku ? formattedMap : ''; if (idx('Variant Cost') !== -1) row[idx('Variant Cost')] = formattedDealer; break; case 'INFLOW': if (idx('Name') !== -1) row[idx('Name')] = platformSku; if (idx('UnitPrice') !== -1) row[idx('UnitPrice')] = formattedPrice; if (idx('Cost') !== -1) row[idx('Cost')] = formattedDealer; break; case 'SELLERCLOUD': if (idx('ProductID') !== -1) row[idx('ProductID')] = platformSku; if (idx('MAPPrice') !== -1) row[idx('MAPPrice')] = formattedMap; if (idx('SitePrice') !== -1) row[idx('SitePrice')] = formattedPrice; if (idx('SiteCost') !== -1) row[idx('SiteCost')] = formattedDealer; break; case 'REVERB': if (idx('sku') !== -1) row[idx('sku')] = platformSku; if (idx('price') !== -1) row[idx('price')] = formattedPrice; if (idx('condition') !== -1) { if (isBAsku && bStockInfo) { switch (bStockInfo.type) { case 'AA': row[idx('condition')] = 'Excellent'; break; case 'BA': case 'BB': row[idx('condition')] = 'Very Good'; break; case 'BC': case 'BD': case 'NOACC': row[idx('condition')] = 'Good'; break; default: row[idx('condition')] = 'Good'; } } else row[idx('condition')] = 'Brand New'; } break; default: Logger.log("Unknown platform for export: " + platform); return null; } return row; }
function generateCPListings() { const ui = SpreadsheetApp.getUi(); const response = ui.alert('Generate CP Listings', 'Identify items from Manufacturer sheet that are not listed (or have low confidence matches) on some platforms. Continue?', ui.ButtonSet.YES_NO); if (response !== ui.Button.YES) return; let cpSheet; try { const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet'); const matchingSheet = SS.getSheetByName('SKU Matching Engine'); cpSheet = SS.getSheetByName('CP Listings'); if (!mfrSheet || !matchingSheet || !cpSheet) { ui.alert('Error', 'Required sheets not found.', ui.ButtonSet.OK); return; } cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge().setValue('PROCESSING - Generating CP listings...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center'); SpreadsheetApp.flush(); const mfrData = mfrSheet.getDataRange().getValues(); const mfrHeaders = mfrData[0]; const mfrSkuCol = mfrHeaders.indexOf('Manufacturer SKU'); const upcCol = mfrHeaders.indexOf('UPC'); const msrpCol = mfrHeaders.indexOf('MSRP'); const mapCol = mfrHeaders.indexOf('MAP'); const dealerPriceCol = mfrHeaders.indexOf('Dealer Price'); if (mfrSkuCol === -1) throw new Error('MFR SKU column not found in Manufacturer Price Sheet.'); const matchingData = matchingSheet.getDataRange().getValues(); const matchingHeaders = matchingData[0]; const matchingMfrSkuCol = matchingHeaders.indexOf('Manufacturer SKU'); const platformPresence = {}; const platformColsInMatching = {}; Object.keys(PLATFORM_TABS).forEach(pKey => { platformColsInMatching[pKey] = { sku: matchingHeaders.indexOf(pKey + ' SKU'), confidence: matchingHeaders.indexOf(pKey + ' Confidence') }; }); for (let i = 2; i < matchingData.length; i++) { const row = matchingData[i]; const mfrSku = row[matchingMfrSkuCol]; if (!mfrSku) continue; if (!platformPresence[mfrSku]) platformPresence[mfrSku] = {}; Object.keys(PLATFORM_TABS).forEach(pKey => { const skuIdx = platformColsInMatching[pKey].sku; const confIdx = platformColsInMatching[pKey].confidence; if (skuIdx !== -1 && row[skuIdx] && confIdx !== -1 && parseFloat(row[confIdx]) >= 85) platformPresence[mfrSku][pKey] = true; else platformPresence[mfrSku][pKey] = false; }); } const cpDataRows = []; const cpHeaders = cpSheet.getRange(1, 1, 1, cpSheet.getLastColumn()).getValues()[0]; const cpHeaderMap = {}; cpHeaders.forEach((h,i) => cpHeaderMap[h] = i); for (let i = 1; i < mfrData.length; i++) { const mfrRow = mfrData[i]; const mfrSku = mfrRow[mfrSkuCol]; if (!mfrSku) continue; const missingOnPlatforms = []; let listedOnAtLeastOne = false; Object.keys(PLATFORM_TABS).forEach(pKey => { if (platformPresence[mfrSku] && platformPresence[mfrSku][pKey]) listedOnAtLeastOne = true; else missingOnPlatforms.push(pKey); }); if (missingOnPlatforms.length > 0) { const cpRow = Array(cpHeaders.length).fill(''); cpRow[cpHeaderMap['Manufacturer SKU']] = mfrSku; if(cpHeaderMap['UPC'] !== undefined && upcCol !== -1) cpRow[cpHeaderMap['UPC']] = mfrRow[upcCol] || ''; if(cpHeaderMap['MSRP'] !== undefined && msrpCol !== -1) cpRow[cpHeaderMap['MSRP']] = mfrRow[msrpCol]; if(cpHeaderMap['MAP'] !== undefined && mapCol !== -1) cpRow[cpHeaderMap['MAP']] = mfrRow[mapCol]; if(cpHeaderMap['Dealer Price'] !== undefined && dealerPriceCol !== -1) cpRow[cpHeaderMap['Dealer Price']] = mfrRow[dealerPriceCol]; Object.keys(PLATFORM_TABS).forEach(pKey => { if(cpHeaderMap[pKey] !== undefined) cpRow[cpHeaderMap[pKey]] = (platformPresence[mfrSku] && platformPresence[mfrSku][pKey]) ? 'Yes' : 'No'; }); if(cpHeaderMap['Action Needed'] !== undefined) cpRow[cpHeaderMap['Action Needed']] = `List on: ${missingOnPlatforms.join(', ')}`; cpDataRows.push(cpRow); } } if (cpSheet.getLastRow() > 2) cpSheet.getRange(3, 1, cpSheet.getLastRow() - 2, cpHeaders.length).clearContent(); if (cpDataRows.length > 0) { cpSheet.getRange(3, 1, cpDataRows.length, cpHeaders.length).setValues(cpDataRows); applyCPListingsFormatting(cpSheet, cpDataRows.length); cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge().setValue(`CP LISTINGS - ${cpDataRows.length} items identified needing action - ${new Date().toLocaleString()}`).setBackground('#d9ead3'); } else cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge().setValue('All items from Manufacturer sheet appear to be listed with high confidence across all platforms.').setBackground('#d9ead3'); ui.alert('CP Listings Generated', `${cpDataRows.length} items identified for cross-platform listing review.`, ui.ButtonSet.OK); } catch (error) { logError(error, 'generateCPListings', ui); if (cpSheet) cpSheet.getRange(2, 1, 1, cpSheet.getLastColumn()).merge().setValue('ERROR - ' + error.toString()).setBackground('#f4cccc'); } }
function applyCPListingsFormatting(sheet, rowCount) { if (rowCount <= 0) return; const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; const dataStartRow = 3; const rules = sheet.getConditionalFormatRules().slice(); Object.keys(PLATFORM_TABS).forEach(pKey => { const colIndex = headers.indexOf(pKey); if (colIndex !== -1) { const range = sheet.getRange(dataStartRow, colIndex + 1, rowCount, 1); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#d9ead3').setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#f4cccc').setRanges([range]).build()); } }); sheet.setConditionalFormatRules(rules); ['MSRP', 'MAP', 'Dealer Price'].forEach(colName => { const colIndex = headers.indexOf(colName); if (colIndex !== -1) sheet.getRange(dataStartRow, colIndex + 1, rowCount, 1).setNumberFormat('$#,##0.00'); }); }
function identifyDiscontinuedItems() { const ui = SpreadsheetApp.getUi(); const response = ui.alert('Identify Discontinued Items', 'Analyze platform SKUs that are not found in the Manufacturer Price Sheet. This may indicate discontinued items. Continue?', ui.ButtonSet.YES_NO); if (response !== ui.Button.YES) return; let discontinuedSheet; try { const mfrSheet = SS.getSheetByName('Manufacturer Price Sheet'); const platformDbSheet = SS.getSheetByName('Platform Databases'); discontinuedSheet = SS.getSheetByName('Discontinued'); if (!discontinuedSheet) discontinuedSheet = createDiscontinuedTab(); if (!mfrSheet || !platformDbSheet) { ui.alert('Error', 'Required sheets (MFR Price Sheet, Platform Databases) not found.', ui.ButtonSet.OK); return; } discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge().setValue('PROCESSING - Identifying discontinued items...').setBackground('#f9cb9c').setFontWeight('bold').setHorizontalAlignment('center'); SpreadsheetApp.flush(); const mfrData = mfrSheet.getDataRange().getValues(); const mfrHeaders = mfrData[0]; const mfrSkuCol = mfrHeaders.indexOf('Manufacturer SKU'); if (mfrSkuCol === -1) throw new Error('Manufacturer SKU column not found in Manufacturer Price Sheet.'); const mfrSkuSet = new Set(); const mfrCoreSkuSet = new Set(); const mfrOriginalSkuSet = new Set(); for (let i = 1; i < mfrData.length; i++) { const rawMfrSku = mfrData[i][mfrSkuCol]; if (rawMfrSku) { mfrOriginalSkuSet.add(String(rawMfrSku).toUpperCase()); const normMfr = conservativeNormalizeSku(rawMfrSku); mfrSkuSet.add(normMfr); const mfrAttributes = extractSkuAttributesAndCore(rawMfrSku); if (mfrAttributes.coreSku) mfrCoreSkuSet.add(mfrAttributes.coreSku); } } const platformData = getPlatformDataFromStructuredSheet(); const discontinuedDataRows = []; const discHeaders = discontinuedSheet.getRange(1,1,1, discontinuedSheet.getLastColumn()).getValues()[0]; const discHeaderMap = {}; discHeaders.forEach((h,i) => discHeaderMap[h] = i); for (const platformKey in platformData) platformData[platformKey].forEach(item => { if (!item.sku) return; const platOriginalUpper = String(item.sku).toUpperCase(); const platAttributes = extractSkuAttributesAndCore(item.sku); let foundInMfr = false; if (mfrOriginalSkuSet.has(platOriginalUpper)) foundInMfr = true; else if (mfrSkuSet.has(platAttributes.normalizedSku)) foundInMfr = true; else if (platAttributes.coreSku && mfrCoreSkuSet.has(platAttributes.coreSku)) foundInMfr = true; if (!foundInMfr) { const row = Array(discHeaders.length).fill(''); row[discHeaderMap['Platform SKU']] = item.sku; row[discHeaderMap['Platform']] = platformKey; row[discHeaderMap['Brand']] = extractBrandFromSku(item.sku, platformKey); row[discHeaderMap['Current Price']] = item.price || ''; row[discHeaderMap['Last Updated']] = new Date(); row[discHeaderMap['Status']] = 'Potential Discontinued'; discontinuedDataRows.push(row); } }); if (discontinuedSheet.getLastRow() > 2) discontinuedSheet.getRange(3, 1, discontinuedSheet.getLastRow() - 2, discHeaders.length).clearContent(); if (discontinuedDataRows.length > 0) { discontinuedSheet.getRange(3, 1, discontinuedDataRows.length, discHeaders.length).setValues(discontinuedDataRows); applyDiscontinuedFormatting(discontinuedSheet, discontinuedDataRows.length); discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge().setValue(`${discontinuedDataRows.length} potential discontinued items found - ${new Date().toLocaleString()}`).setBackground('#d9ead3'); } else discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge().setValue('No potential discontinued items found based on current criteria.').setBackground('#d9ead3'); ui.alert('Discontinued Item Analysis Complete', `${discontinuedDataRows.length} items found on platforms but not in the Manufacturer Price Sheet have been listed in the "Discontinued" tab.`, ui.ButtonSet.OK); } catch (error) { logError(error, 'identifyDiscontinuedItems', ui); if (discontinuedSheet) discontinuedSheet.getRange(2, 1, 1, discontinuedSheet.getLastColumn()).merge().setValue('ERROR - ' + error.toString()).setBackground('#f4cccc'); } }
function extractBrandFromSku(sku, platform) { if (!sku) return 'Unknown'; const upperSku = String(sku).toUpperCase(); const brandPrefixes = { 'FENDER': 'Fender', 'GIBSON': 'Gibson', 'SQUIER': 'Squier', 'EPIPHONE': 'Epiphone', 'IBANEZ': 'Ibanez', 'YAMAHA': 'Yamaha', 'PRS': 'PRS', 'MARTIN': 'Martin', 'TAYLOR': 'Taylor', 'BOSS': 'Boss', 'ROLAND': 'Roland', 'KORG': 'Korg', 'MOOG': 'Moog', 'SHURE': 'Shure', 'SENNHEISER': 'Sennheiser', 'AKG': 'AKG', 'RODE': 'Rode', 'MACKIE': 'Mackie', 'BEHRINGER': 'Behringer', 'PRESONUS': 'Presonus', 'FOCUSRITE': 'Focusrite', 'EMG': 'EMG', 'DIMARZIO': 'DiMarzio', 'SEYMOUR DUNCAN': 'Seymour Duncan', 'GODIN': 'Godin', 'G&L': 'G&L', 'GRETSCH': 'Gretsch', 'JACKSON': 'Jackson', 'CHARVEL': 'Charvel', 'EVH': 'EVH', 'PEAVEY': 'Peavey', 'ORANGE': 'Orange Amps', 'MARSHALL': 'Marshall', 'VOX': 'Vox', 'LINE 6': 'Line 6', 'DUNLOP': 'Dunlop', 'MXR': 'MXR', 'ELECTRO-HARMONIX': 'Electro-Harmonix', 'EHX': 'Electro-Harmonix', 'STRYMON': 'Strymon', 'KEELEY': 'Keeley', 'WALRUS AUDIO': 'Walrus Audio', 'JHS': 'JHS Pedals', 'ZILDJIAN': 'Zildjian', 'SABIAN': 'Sabian', 'PAISTE': 'Paiste', 'MEINL': 'Meinl', 'PEARL': 'Pearl Drums', 'TAMA': 'Tama', 'DW': 'DW Drums', 'LUDWIG': 'Ludwig', 'SHU-': 'Shure', 'MCK-': 'Mackie', 'GGC-': 'Godin', 'FEND': 'Fender', 'GIB': 'Gibson', 'IBZ': 'Ibanez', 'YMH': 'Yamaha', 'MART': 'Martin', 'TAYL': 'Taylor', 'RLND': 'Roland', 'SENNH': 'Sennheiser', 'PRESO': 'Presonus', 'FOCU': 'Focusrite', 'SEYM': 'Seymour Duncan' }; const parts = upperSku.split(/[-_ ]/); if (parts.length > 0) { const firstPart = parts[0]; if (brandPrefixes[firstPart]) return brandPrefixes[firstPart]; for (const pfxKey of Object.keys(brandPrefixes).sort((a,b) => b.length - a.length)) if (firstPart.startsWith(pfxKey)) return brandPrefixes[pfxKey]; } for (const brandKey of Object.keys(brandPrefixes).sort((a,b) => b.length - a.length)) if (upperSku.includes(brandKey)) return brandPrefixes[brandKey]; return 'Unknown'; }
function applyDiscontinuedFormatting(sheet, rowCount) { if (rowCount <= 0) return; const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; const dataStartRow = 3; const rules = sheet.getConditionalFormatRules().slice(); const platformCol = headers.indexOf('Platform'); if (platformCol !== -1) { const range = sheet.getRange(dataStartRow, platformCol + 1, rowCount, 1); Object.keys(PLATFORM_TABS).forEach(pKey => { if (COLORS[pKey]) { const bgColor = COLORS[pKey]; let fontColor = '#FFFFFF'; if (bgColor.match(/^#[0-9A-F]{6}$/i)) { const r = parseInt(bgColor.substring(1,3), 16); const g = parseInt(bgColor.substring(3,5), 16); const b = parseInt(bgColor.substring(5,7), 16); if ((r*0.299 + g*0.587 + b*0.114) > 150) fontColor = '#000000'; } rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(pKey).setBackground(bgColor).setFontColor(fontColor).setRanges([range]).build()); } }); } sheet.setConditionalFormatRules(rules); const priceCol = headers.indexOf('Current Price'); if (priceCol !== -1) sheet.getRange(dataStartRow, priceCol + 1, rowCount, 1).setNumberFormat('$#,##0.00'); const dateCol = headers.indexOf('Last Updated'); if (dateCol !== -1) sheet.getRange(dataStartRow, dateCol + 1, rowCount, 1).setNumberFormat('M/d/yyyy h:mm:ss');
}
