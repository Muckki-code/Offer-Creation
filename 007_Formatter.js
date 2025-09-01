/**
 * @file This file contains shared helper functions for formatting data,
 * such as dates, numbers, and terms, for display purposes.
 */

/**
 * Formats a date string or object into a localized date string.
 * @param {string|Date} dateStringYYYYMMDD The date to format.
 * @param {string} language The target language ('german' or 'english').
 * @returns {string} The formatted date string.
 */
function formatDateForLocale(dateStringYYYYMMDD, language) {
  const sourceFile = "Formatter_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_start' });
  Log[sourceFile](`[${sourceFile} - formatDateForLocale] Start: date='${dateStringYYYYMMDD}', language='${language}'`);
  if (!dateStringYYYYMMDD) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_noDate' });
    return "";
  }
  let dateObj;
  if (dateStringYYYYMMDD instanceof Date && !isNaN(dateStringYYYYMMDD)) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_isDateObject' });
    dateObj = dateStringYYYYMMDD;
  } else if (typeof dateStringYYYYMMDD === 'string' && dateStringYYYYMMDD.match(/^\d{4}-\d{2}-\d{2}$/)) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_isString' });
    const parts = dateStringYYYYMMDD.split('-');
    if (parts.length === 3) {
      Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_stringHas3Parts' });
      dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
    }
  }
  if (!dateObj || isNaN(dateObj.getTime())) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_invalidDate' });
    return String(dateStringYYYYMMDD);
  }
  if (language === "german") {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_german' });
    return dateObj.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }
  // Default to YYYY-MM-DD for English or other languages
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatDateForLocale_english' });
  return dateObj.getFullYear() + '-' + ('0' + (dateObj.getMonth() + 1)).slice(-2) + '-' + ('0' + dateObj.getDate()).slice(-2);
}

/**
 * Formats a number into a localized currency string.
 * @param {number} number The number to format.
 * @param {string} language The target language ('german' or 'english').
 * @param {boolean} includeCurrencySymbol Whether to include the '€' symbol.
 * @returns {string} The formatted currency string.
 */
function formatNumberForLocale(number, language, includeCurrencySymbol) {
  const sourceFile = "Formatter_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatNumberForLocale_start' });
  Log[sourceFile](`[${sourceFile} - formatNumberForLocale] Start: number='${number}', language='${language}'`);
  if (number === null || number === undefined || isNaN(parseFloat(number))) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatNumberForLocale_invalidNumber' });
    return "";
  }
  const num = parseFloat(number);
  const options = { minimumFractionDigits: 2, maximumFractionDigits: 2 };
  const locale = language === "german" ? (Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatNumberForLocale_german' }), 'de-DE') : (Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatNumberForLocale_english' }), 'en-US');
  const formattedNumber = num.toLocaleString(locale, options);
  return includeCurrencySymbol ? (Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatNumberForLocale_withCurrency' }), `€ ${formattedNumber}`) : (Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatNumberForLocale_withoutCurrency' }), formattedNumber);
}

/**
 * Formats a term number into a localized string with units.
 * @param {number|string} termValue The term in months.
 * @param {string} language The target language ('german' or 'english').
 * @returns {string} The formatted term string (e.g., "24 Monate").
 */
function formatTermForLocale(termValue, language) {
  const sourceFile = "Formatter_gs";
  Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatTermForLocale_start' });
  Log[sourceFile](`[${sourceFile} - formatTermForLocale] Start: term='${termValue}', language='${language}'`);
  if (termValue === null || termValue === undefined) {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatTermForLocale_invalidTerm' });
    return "";
  }
  const cleanedTerm = String(termValue).replace(/[^0-9]/g, "").trim();
  if (cleanedTerm === "") {
    Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatTermForLocale_emptyTerm' });
    return String(termValue);
  }
  const suffix = language === "german" ? (Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatTermForLocale_german' }), " Monate") : (Log.TestCoverage_gs({ file: sourceFile, coverage: 'formatTermForLocale_english' }), " months");
  return cleanedTerm + suffix;
}