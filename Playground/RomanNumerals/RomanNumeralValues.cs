using System.Collections.Generic;

namespace Playground
{
    public class RomanNumeralValues
    {
        public static Dictionary<char, int> ExtendedRomanNumerals = new Dictionary<char, int>() {
            { 'm', 1000000 },
            { 'd', 500000 },
            { 'c', 100000 },
            { 'l', 50000 },
            { 'x', 10000 },
            { 'v', 5000 },
            { 'M', 1000 },
            { 'D', 500 },
            { 'C', 100 },
            { 'L', 50 },
            { 'X', 10 },
            { 'V', 5 },
            { 'I', 1 }
        };

        public static Dictionary<string, string> ExtendedShortHands = new Dictionary<string, string>() {
            { "dcccc", "cm"},
            { "cccc", "cd" },
            { "lxxxx", "xc"},
            { "xxxx", "xl" },
            { "vMMMM", "Mx"},
            { "MMMM", "Mv" },
            { "DCCCC", "CM"},
            { "CCCC", "CD" },
            { "LXXXX", "XC"},
            { "XXXX", "XL" },
            { "VIIII", "IX"},
            { "IIII", "IV" }
        };

        public static Dictionary<char, int> RomanNumerals = new Dictionary<char, int>() {
            { 'M', 1000 },
            { 'D', 500 },
            { 'C', 100 },
            { 'L', 50 },
            { 'X', 10 },
            { 'V', 5 },
            { 'I', 1 }
        };

        public static Dictionary<string, string> ShortHands = new Dictionary<string, string>() {
            { "DCCCC", "CM"},
            { "CCCC", "CD" },
            { "LXXXX", "XC"},
            { "XXXX", "XL" },
            { "VIIII", "IX"},
            { "IIII", "IV" }
        };
    }
}
