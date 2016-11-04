using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Playground
{
    class Depricated
    {
        public static int RomanNumeralToNumber(string roman)
        {
            int result = 0;
            if (roman.All(c => "MmDdCcLlXxVvIi".Contains(c))) {
                if (roman.Length > 0) {
                    char[] values = roman.ToArray();
                    if (values.Length > 1) {
                        for (int i = 1; i < values.Length; i++) {
                            if (RomanNumeralValues.ExtendedRomanNumerals[values[i - 1]] >= RomanNumeralValues.ExtendedRomanNumerals[values[i]]
                                || (values[i - 1].Equals('D') && values[i].Equals('M'))
                                || (values[i - 1].Equals('d') && values[i].Equals('m'))) {
                                result += RomanNumeralValues.ExtendedRomanNumerals[values[i - 1]];
                                if (i == values.Length - 1) {
                                    result += RomanNumeralValues.ExtendedRomanNumerals[values[i]];
                                }
                            } else {
                                result += (RomanNumeralValues.ExtendedRomanNumerals[values[i]] - RomanNumeralValues.ExtendedRomanNumerals[values[i - 1]]);
                                i++;
                                if (i + 1 == values.Length) {
                                    result += RomanNumeralValues.ExtendedRomanNumerals[values[i]];
                                }
                            }
                        }
                    } else {
                        result += RomanNumeralValues.ExtendedRomanNumerals[values[0]];
                    }
                }
            }
            return result;
        }
    }
}
