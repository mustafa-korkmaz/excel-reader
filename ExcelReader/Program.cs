using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.cs
{
    class Program
    {
        static void Main(string[] args)
        {
            var reader = new Reader();
            reader.ReadPlayers();
            reader.ReadTeams();
        }
    }
}
