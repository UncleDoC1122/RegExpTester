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

            Regex index = new Regex(@"[0-9]*", RegexOptions.IgnoreCase);
            Regex country = new Regex(@"Россия|РФ|Российская Федерация|Украина", RegexOptions.IgnoreCase);
            Regex region = new Regex(@"обл|область|край|респ|республика", RegexOptions.IgnoreCase);
            Regex cities = new Regex(@"г|город|село|с|пос|поселок|посёлок", RegexOptions.IgnoreCase);
            Regex streetPrefix = new Regex(@"ул|мкр|проспект|пр|пер|пр-кт|бульвар|б-р|промзона|проезд|пр-т|наб|набережная|пл|площадь|улица|бул", RegexOptions.IgnoreCase);
            Regex streetPostfix = new Regex(@"пер|переулок|тракт|наб|набережная|пр-д|проезд|тупик|ш|шоссе|бульвар|б|пл|проспект|пр-кт|площадь|переулок|пер|пр-т|пгт", RegexOptions.IgnoreCase);
            Regex house = new Regex(@"д|[0-9]*", RegexOptions.IgnoreCase);
            Regex apartment = new Regex(@"кв|оф|офис|пом|помещение|квартира", RegexOptions.IgnoreCase);
            Regex housing = new Regex(@"к|корп|корпус", RegexOptions.IgnoreCase);
            Regex building = new Regex(@"с|строение|стр|литера", RegexOptions.IgnoreCase);

            List<Regex> tokens = new List<Regex>();

            tokens.Add(index);
            tokens.Add(country);
            tokens.Add(region);
            tokens.Add(cities);
            tokens.Add(streetPrefix);
            tokens.Add(streetPostfix);
            tokens.Add(house);
            tokens.Add(apartment);
            tokens.Add(housing);
            tokens.Add(building);

            for (int i = 0; i < 10; i ++)
            {
                output.Add(new List<string>()); // создание списка на 10 списков
            }



            for (int i = 0; i < input.Count; i ++) // перебор всех входных элементов
            {

                List<string> addable = new List<string>(); // список вывода
                string[] tmp = input[i].Split(','); // разделить строку по запятым

                for (int j = 0; j < tmp.Length; j ++) // перебор всех токенов строки
                {

                    bool flag = false; // флаг для обозначения элемента, который не определился
                    for (int k = 0; k < tokens.Count; k ++) // перебор по списку токенов
                    {
                        Match match = tokens[k].Match(tmp[j]);
                        if (match.Length > 0)
                        {
                            output[i].Add(tmp[j]);
                            flag = true;
                        }
                        if (flag == true) break;
                       
                    }

                }
                
                
            }



            return output;
        }

       

        
    }
}
