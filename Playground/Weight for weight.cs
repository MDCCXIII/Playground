using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Playground
{
    class Weight_for_weight
    {
        public static string orderWeight(string strng)
        {
            Dictionary<string, int> weights = getWieghts(strng);
            CalcWeightsInNumbers(weights);

        }

        private static void CalcWeightsInNumbers(Dictionary<string, int> weights)
        {
            foreach (string weight in weights.Keys) {
                int weightInNumbers = 0;
                foreach (int val in weight) {
                    weightInNumbers += val;
                }
                weights[weight] = weightInNumbers;
            }
        }

        private static Dictionary<string, int> getWieghts(string strng)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
              string[] weights = strng.Split(' ');
            foreach (string weight in weights) {
                result.Add(weight, 0);
            }
            return result;
        }
    }
}
