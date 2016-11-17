using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Playground
{
    public class NextNumbers
    {
        public NextNumbers(long n)
        {
            Console.WriteLine(NextBiggerNumber(n));
            Console.ReadLine();
        }

        public static long NextBiggerNumber(long n)
        {
            char[] c = n.ToString().ToCharArray();
            long result = -1;

            for (int i = c.Length - 2; i > -1; i--) {
                if(c[i + 1] > c[i]) {
                    string validPart = n.ToString().Substring(0, i);
                    string changed = "";
                    if(nextClosestNumber(c, i)) {
                        changed = getNextNumber(c, i, changed);
                    }
                    
                    result = long.Parse(validPart + changed);
                    break;
                }
            }

            return result;
        }

        private static bool nextClosestNumber(char[] c, int i)
        {
            for (int j = i; j > -1; j--) {
                if (c[j] > c[i] && c[j] < c[j + 1]) {
                    return false;
                }
            }
            return true;
        }

        private static string getNextNumber(char[] c, int i, string changed)
        {
            string changing = new string(c, i, c.Length - i);
            int distinctCharacterCount = changing.Distinct().Count();
            char max = ' ';
            for (int j = 0; j < distinctCharacterCount; j++) {
                try {
                    if (changed.Length == 0) {
                        max = changing.Where((l, m) => l > m && l > changing.ToArray()[0]).OrderBy(l => l).ToArray().First();
                    } else {
                        max = changing.Where((l, m) => l > m).OrderBy(l => l).ToArray().First();
                    }
                }
                catch (Exception) {
                    max = changing.ToArray()[0];
                }

                int count = changing.Select(l => l).Where(l => l.Equals(max)).Count();

                changed += new string(max, count);
                changing = changing.Replace(max.ToString(), "");
            }

            return changed;
        }
    }
}
