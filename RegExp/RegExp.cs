using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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


            for (int i = 0; i < input.Count; i ++)
            {
                string[] tmp = input[i].Split(',');

                
            }



            return output;
        }

        private Tuple<int, string> chooser(string input)
        {
            Tuple<int, string> output;

            Regex index = new Regex(@"+[0-9]");
            Regex country = new Regex(@"Россия|РФ|Российская Феерация|Украина");
            Regex region = new Regex(@"обл|область|край|респ|республика");
            Regex cities = new Regex(@"г|город|село|с|пос|поселок|посёлок");
            Regex streetPrefix = new Regex(@"ул|мкр|проспект|пр|пер|пр-кт|бульвар|б-р|промзона|проезд|пр-т|наб|набережная|пл|площадь|улица|бул");
            Regex streetPostfix = new Regex(@"пер|переулок||||||||||||||||||")

            return output;
        }

        
    }
}
