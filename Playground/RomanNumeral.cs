using System;
using System.Collections.Generic;
using System.Linq;

namespace Playground {
    class RomanNumeral {

        static Dictionary<char, int> RN = new Dictionary<char, int>() {
            { 'M',1000 },
            { 'D', 500 },
            { 'C', 100 },
            { 'L',  50 },
            { 'X',  10 },
            { 'V',   5 },
            { 'I',   1 },
            { 'm',1000000 },
            { 'd', 500000 },
            { 'c', 100000 },
            { 'l',  50000 },
            { 'x',  10000 },
            { 'v',   5000 },
            { 'i',   1000 }
        };

        public static void solution() {
            string lastCommand = "";
            Console.WriteLine("Enter a Roman number. ");
            while (!lastCommand.ToLower().Equals("exit")) {
                lastCommand = ExecuteUserInterface();
            }
        }

        private static string ExecuteUserInterface() {
            string lastCommand;
            Console.Write("-> ");
            lastCommand = Console.ReadLine();
            if (lastCommand.StartsWith("/help") || lastCommand.StartsWith("/?")) {
                NavOptions(lastCommand);
            } else if (lastCommand.StartsWith("/") || lastCommand.ToLower().Contains("help") || lastCommand.ToLower().Contains("?")) {
                Console.WriteLine("Proper usage is '/help [command]' or '/? [command]'.");
                Console.WriteLine("For a reference list of Roman numeral characters and thier values type '/help' or '/?'.");
                Console.WriteLine("For a list of commands type '/help commands' or '/? commands'.");
            } else {
                Console.WriteLine(RomanNumeralToNumber(lastCommand));
            }
            Console.WriteLine();
            return lastCommand;
        }

        private static void NavOptions(string command) {
            command = command.Replace("/help", "").Replace("/?", "").TrimStart();
            switch (command) {
                case "":
                    PrintRomanNumeralChart();
                    break;
                case "commands":
                    Console.WriteLine("Sorry, at this time there are no valid help options.");
                    Console.WriteLine("A reference list of Roman numeral characters and thier values may be found by typing '/help' or '/?'.");
                    break;
                default:
                    Console.WriteLine("Error: '" + command + "' is not a command.");
                    Console.WriteLine("For a list of commands type '/help commands' or '/? commands'.");
                    break;
            }
        }

        private static void PrintRomanNumeralChart() {
            Console.WriteLine("Valid Roman numerals and their equivalent values are as follows: ");
            foreach (KeyValuePair<char, int> kvp in RN.OrderBy(pair => pair.Value)) {
                Console.WriteLine(kvp.Key + " = " + kvp.Value);
            }
        }

        private static int RomanNumeralToNumber(string roman) {
            int result = 0;
            if (roman.All(c => "MmDdCcLlXxVvIi".Contains(c))) {
                if (roman.Length > 0) {
                    char[] values = roman.ToArray();
                    if (values.Length > 1) {
                        for (int i = 1; i < values.Length; i++) {
                            if (RN[values[i - 1]] >= RN[values[i]] ) {
                                result += RN[values[i - 1]];
                                if (i == values.Length - 1) {
                                    result += RN[values[i]];
                                }
                            } else {
                                result += (RN[values[i]] - RN[values[i - 1]]);
                                i++;
                                if (i + 1 == values.Length) {
                                    result += RN[values[i]];
                                }
                            }
                        }
                    } else {
                        result += RN[values[0]];
                    }
                }
            }
            return result;
        }
    }
}
