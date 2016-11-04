using System.Collections.Generic;
using System.Linq;

namespace Playground {
    class RomanNumeralDecoder {
        public static int RomanNumeralToNumber(string romanNumeral)
        {
            int result = 0;
            if (romanNumeral.Length > 0 && romanNumeral.All(c => RomanNumeralValues.ValidRomanNumerals.Contains(c))) {
                romanNumeral = ConvertToLongHand(romanNumeral);
                foreach (char c in romanNumeral.ToCharArray()) {
                    result += RomanNumeralValues.RomanNumerals[c];
                }
            }
            return result;
        }

        private static string ConvertToLongHand(string romanNumeral)
        {
            foreach (KeyValuePair<string, string> kvp in RomanNumeralValues.ShortHands) {
                if (romanNumeral.Contains(kvp.Value)) {
                    romanNumeral = romanNumeral.Replace(kvp.Value, kvp.Key);
                }
            }

            return romanNumeral;
        }
    }
}
