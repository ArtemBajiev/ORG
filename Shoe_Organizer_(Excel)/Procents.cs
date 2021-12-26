using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shoe_Organizer__Excel_
{
    public class Procents
    {
        public decimal Prc(decimal x, decimal y)
        {
            return Math.Round(x * ((y / 100) + 1), 2); //Исправление ошибки
        }
    }
}
