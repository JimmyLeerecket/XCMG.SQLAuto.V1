using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XCMG.SQLAuto.V1.Study
{
    public class MethodSyntax
    {
        public void MainMethodSyntax()
        {
            int[] array = [1, 2, 3, 4, 5, 6, 7, 8, 9];
            var _queryMany = array.SelectMany(i => array
                                    .Where(j => j <= i)
                                    .Select(j => $"{(j == 1 ? "\n" : "")}{j}*{i}={i * j}\t")
                                );

            Console.WriteLine(string.Concat(_queryMany));

            IEnumerable<string> emoji = ["🤓", "💯", "🔥", "🎉", "👀", "⭐", "💜", "✔"];

            string[] words = ["the", "quick", "brown", "fox", "jumped", "over", "the", "lazy", "dog"];
            foreach (string word in words.DistinctBy(p => p.Length))
            {
                Console.WriteLine(word);
            }


            string[] words1 = ["the", "quick", "brown", "fox"];
            string[] words2 = ["jumped", "over", "the", "lazy", "dog"];
            IEnumerable<string> query = from word in words1.Except(words2)
                                        select word;


        }

        
        
    }

}
