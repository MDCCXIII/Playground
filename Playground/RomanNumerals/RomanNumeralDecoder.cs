using System.Collections.Generic;
using System.Linq;

namespace Playground {
    class RomanNumeralDecoder {
        public static int RomanNumeralToNumber(string romanNumeral)
        {
            int result = 0;
            if (romanNumeral.Length > 0 && romanNumeral.All(c => "MmDdCcLlXxVvI".Contains(c))) {
                foreach (KeyValuePair<string, string> kvp in RomanNumeralValues.ExtendedShortHands) {
                    if (romanNumeral.Contains(kvp.Value)) {
                        romanNumeral = romanNumeral.Replace(kvp.Value, kvp.Key);
                    }
                }
                foreach (char c in romanNumeral.ToCharArray()) {
                    result += RomanNumeralValues.ExtendedRomanNumerals[c];
                }
            }
            return result;
        }
    }
}
