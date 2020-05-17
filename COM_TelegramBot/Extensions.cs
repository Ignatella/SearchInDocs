using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace COM_TelegramBot
{
    public static class Extensions
    {
        public static void Foreach<T>(this IEnumerable<T> source, Action<T> action)
        {
            if (source == null)
                throw new ArgumentNullException();

            foreach (var item in source)
            {
                action(item);
            }
        }
    }
}
