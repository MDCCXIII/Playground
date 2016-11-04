﻿using System;
using System.Collections.Generic;

namespace Playground
{
    class RomanNumeralEncoder
    {
        public static String NumberToRomanNumeral(int n)
        {
            string result = "";
            int remainder = n;
            foreach(KeyValuePair<char, int> kvp in RomanNumeralValues.RomanNumerals) {
                result += new string(kvp.Key, remainder / kvp.Value);
                remainder = n % kvp.Value;
                if(remainder == 0) {
                    break;
                }
            }

            return ConvertToShortHand(result);
        }

        private static string ConvertToShortHand(string result)
        {
            foreach(KeyValuePair<string, string> kvp in RomanNumeralValues.ShortHands) {
                if (result.Contains(kvp.Key)) {
                    result = result.Replace(kvp.Key, kvp.Value);
                }
            }
            return result;
        }
    }
}
