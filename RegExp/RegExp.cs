using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RegExp
{
    class RegExp
    {
        public static List<List<string>> parseAddress(List<string> input)
        {
            List<List<string>> output = new List<List<string>>(); // в списке output будет находиться 10 списков строчных значений: 
                                                                  // индекс - страна - регион - город - улица - дом - корпус - строение - квартира\помещение - служебный код

            for (int i = 0; i < 10; i ++)
            {
                output.Add(new List<string>()); // создание списка на 10 списков
            }



            return output;
        }

        private static Tuple<int, string> parseUnknown(string unknownToken, string cell)
        {
            Tuple<int, string> output;

        }
    }
}
