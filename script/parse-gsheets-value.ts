/**
 * Parses a raw string value from a GSheets CSV export into a typed value.
 *
 * GSheets exports financial function results with currency and percent
 * formatting (e.g. `$6.25`, `-$1,848.51`, `9%`). These must be converted
 * to numeric values for accurate comparison with HyperFormula output.
 *
 * @param raw - Raw string from the CSV cell
 * @returns Typed value: number, boolean, null, or string
 */
export function parseGSheetsValue(raw: string): string | number | boolean | null {
  const trimmed = raw.trim();

  // Empty → null
  if (trimmed === "") return null;

  // Booleans
  if (trimmed === "TRUE") return true;
  if (trimmed === "FALSE") return false;

  // Error strings (keep as-is)
  if (trimmed.startsWith("#")) return trimmed;

  // Currency: -$N,NNN.NN or $N,NNN.NN
  const currencyMatch = trimmed.match(/^(-?)\$([0-9,]+(\.[0-9]+)?)$/);
  if (currencyMatch) {
    const sign = currencyMatch[1] === "-" ? -1 : 1;
    const digits = currencyMatch[2]!.replace(/,/g, "");
    const num = Number(digits);
    if (!isNaN(num)) return sign * num;
  }

  // Percentage: N% or N.NN%
  const percentMatch = trimmed.match(/^(-?[0-9]+(\.[0-9]+)?)%$/);
  if (percentMatch) {
    const num = Number(percentMatch[1]);
    if (!isNaN(num)) return num / 100;
  }

  // Preserve leading-zero strings as strings: they represent base-conversion
  // output (e.g. DEC2BIN → "01100100", OCT2HEX → "0000001F"). Number("010011")
  // would strip the leading zeros and produce 10011, causing a mismatch.
  if (trimmed.length > 1 && trimmed.startsWith("0")) return trimmed;

  // Try plain numeric
  const num = Number(trimmed);
  if (!isNaN(num) && trimmed !== "") return num;

  // Default: string
  return trimmed;
}
