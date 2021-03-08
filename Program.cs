using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace AbbyOCRProcessOutputConversion
{
    class Program
    {
        static void Main(string[] args)
        {

            var myExcelService = new ExcelService();

            var myFilePaths = Directory.GetFiles(args[0], "*.xlsx", SearchOption.AllDirectories);
            var myOrderFilePaths = myFilePaths.OrderBy(x => x);

            foreach (var myFilePath in myOrderFilePaths)
            {
                string myOutputFileName = myFilePath.Replace(".xlsx", ".csv");

                // remove date time file string
                //myOutputFileName = Regex.Replace(myOutputFileName, @"_\d{4}_\d{2}_\d{2}_\d{2}_\d{2}_\d{2}", "");
                if (args.Count() > 2)
                {
                    myOutputFileName = Regex.Replace(myOutputFileName, args[1], "");
                }

                if (File.Exists(myOutputFileName))
                {
                    var myNum = 1;

                    while (File.Exists(myOutputFileName))
                    {
                        myOutputFileName = myOutputFileName.Replace(".csv", "-" + myNum.ToString() + ".csv");
                        myNum++;
                    }
                }

                myExcelService.ConvertExcelToCsv(myFilePath, myOutputFileName);
            }
        }

        

    }
}
