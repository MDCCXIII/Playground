using System;
using System.Collections.Generic;
using System.Linq;

namespace Playground
{
    public class RomanNumeralConsole
    {
        public RomanNumeralConsole()
        {
            string lastCommand = "";
            Console.WriteLine("Enter a Roman number or a number. ");
            while (!lastCommand.ToLower().Equals("exit")) {
                lastCommand = ExecuteUserInterface();
            }
        }

        private static string ExecuteUserInterface()
        {
            string lastCommand;
            Console.Write("-> ");
            lastCommand = Console.ReadLine();
            ParseCommand(lastCommand);
            Console.WriteLine();
            return lastCommand;
        }

        private static void ParseCommand(string lastCommand)
        {
            if (lastCommand.StartsWith("/help") || lastCommand.StartsWith("/?")) {
                NavOptions(lastCommand);
            } else if (lastCommand.StartsWith("/") || lastCommand.ToLower().Contains("help") || lastCommand.ToLower().Contains("?")) {
                Console.WriteLine("Proper usage is '/help [command]' or '/? [command]'.");
                Console.WriteLine("For a reference list of Roman numeral characters and thier values type '/help' or '/?'.");
                Console.WriteLine("For a list of commands type '/help commands' or '/? commands'.");
            } else {
                int number = 0;
                if (Int32.TryParse(lastCommand, out number)) {
                    Console.WriteLine(RomanNumeralEncoder.NumberToRomanNumeral(number));
                } else {
                    Console.WriteLine(RomanNumeralDecoder.remake(lastCommand));
                }

            }
        }

        private static void NavOptions(string command)
        {
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

        private static void PrintRomanNumeralChart()
        {
            Console.WriteLine("Valid Roman numerals and their equivalent values are as follows: ");
            foreach (KeyValuePair<char, int> kvp in RomanNumeralValues.ExtendedRomanNumerals.OrderBy(pair => pair.Value)) {
                Console.WriteLine(kvp.Key + " = " + kvp.Value);
            }
        }
    }
}
