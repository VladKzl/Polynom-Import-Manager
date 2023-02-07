using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TCS_Polynom_data_actualiser
{
    public static class CommonCode
    {
        public static void GetPercent(Int32 max, Int32 min, string args = null)
        {
            Int32 percent = (Int32)(min / (max / 100M));
            Console.Write($"\r{args} {percent}%"); 
            if(percent == 100)
                Console.WriteLine();
        }
        public static bool UserValidation()
        {
            bool userActionBool = true;
            do
            {
                Console.WriteLine("+/-");
                string userAction = Console.ReadLine();
                switch (userAction)
                {
                    case "+":
                        userActionBool = true;
                        break;
                    case "-":
                        userActionBool = false;
                        break;
                    default:
                        Console.WriteLine("Ошибка. Введите + или -.");
                        break;
                }
                if (userAction == "+")
                {
                    break;
                }
                if (userAction == "-")
                {
                    break;
                }
            }
            while (true);
            return userActionBool;
        }
    }
}
