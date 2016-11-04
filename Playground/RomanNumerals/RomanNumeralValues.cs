using System;
using System.Collections.Generic;

namespace Playground
{
    public class RomanNumeralValues
    {
        private const string StandardMessage = "Using standard Roman numerals.";
        private const string ExtendedMessage = "Using extended Roman numerals.";

        private const string StandardRomanCharacters = "MDCLXVI";
        private const string ExtendedRomanCharacters = "MmDdCcLlXxVvI";

        private static Dictionary<char, int> ExtendedRomanNumerals = new Dictionary<char, int>() {
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

        private static Dictionary<string, string> ExtendedShortHands = new Dictionary<string, string>() {
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

        private static Dictionary<char, int> StandardRomanNumerals = new Dictionary<char, int>() {
            { 'M', 1000 },
            { 'D', 500 },
            { 'C', 100 },
            { 'L', 50 },
            { 'X', 10 },
            { 'V', 5 },
            { 'I', 1 }
        };

        private static Dictionary<string, string> StandardShortHands = new Dictionary<string, string>() {
            { "DCCCC", "CM"},
            { "CCCC", "CD" },
            { "LXXXX", "XC"},
            { "XXXX", "XL" },
            { "VIIII", "IX"},
            { "IIII", "IV" }
        };

        private static Dictionary<T, V> AddRange<T, V>(Dictionary<T, V> target, Dictionary<T, V> source)
        {
            if (target == null)
                target = new Dictionary<T, V>();
            if (source == null)
                throw new ArgumentNullException("source");
            foreach (var element in source)
                target.Add(element.Key, element.Value);
            return target;
        }

        public static Dictionary<char, int> RomanNumerals = AddRange(RomanNumerals, StandardRomanNumerals);
        public static Dictionary<string, string> ShortHands = AddRange(ShortHands, StandardShortHands);
        public static string ValidRomanNumerals = StandardRomanCharacters;
        public static string UsingMessage = StandardMessage;

        public static string SetRomanNumerals(bool isExtended) {
            string result = "";
            if (isExtended) {
                RomanNumerals = ExtendedRomanNumerals;
                ShortHands = ExtendedShortHands;
                ValidRomanNumerals = ExtendedRomanCharacters;
                result = ExtendedMessage;
            } else {
                RomanNumerals = StandardRomanNumerals;
                ShortHands = StandardShortHands;
                ValidRomanNumerals = StandardRomanCharacters;
                result = StandardMessage;
            }
            return result;
        }
    }
}
