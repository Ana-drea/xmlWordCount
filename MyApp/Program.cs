using System;
using System.Collections;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml;

namespace MyApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            System.Console.WriteLine("Please enter the source folder path:");
            String path = System.Console.ReadLine();
            List<string> filelist = new List<string>{};
            //get a list of all the mxliff files in the folder
            GetMxliff(path, filelist);
            //create an ExcelPackage
            ExcelPackage excelPkg = new ExcelPackage();
            //add a sheet
            ExcelWorksheet ews = excelPkg.Workbook.Worksheets.Add("Sheet1");
            //add table heads
            ews.Cells[1,1].Value="file name";
            ews.Cells[1,2].Value="word count";
            //Enumerate the mxliff files, get file name and word count of each file
            for (int i = 0; i < filelist.Count; i++)
            {
                ArrayList result = GetSourceString(filelist[i]);
                //write file name and word count into excel
                ews.Cells[i+2,1].Value=result[0];
                ews.Cells[i+2,2].Value=result[1];

            }
            //create the sum formula
            String Formula = String.Format("SUM({0}:{1})", "B2", "B"+(1+filelist.Count));
            //get the total word count in the last cell
            ews.Cells[filelist.Count+2,2].Formula=Formula;
            //save this excel in the source folder
            excelPkg.SaveAs(new FileInfo(Path.Combine(path,"WordCount.xlsx")));
        }

        public static void GetMxliff(String path, List<string> filelist){
            if(Directory.GetDirectories(path)!=null){
                foreach (String subdirectory in Directory.GetDirectories(path))
                {
                    GetMxliff(subdirectory, filelist);
                }
            }
            foreach (string file in Directory.EnumerateFiles(path, "*.mxliff"))
            {
                filelist.Add(file);
            }
        }

        public static ArrayList GetSourceString(String path){
            ArrayList nameandcount = new ArrayList();
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(path);
            XmlNodeList list = xmldoc.GetElementsByTagName("trans-unit");
            int wcount = 0;
            foreach(XmlNode node in list){
                if(node.Attributes["m:confirmed"].Value=="1"){
                    String source = node.ChildNodes[1].InnerText;
                    wcount+=CountWords(source);
                }
            }
            nameandcount.Add(Path.GetFileName(path));
            nameandcount.Add(wcount);
            return nameandcount;
        }

        public static int CountWords(string s)
        {
	        MatchCollection collection = Regex.Matches(s, @"[\S]+");
	        return collection.Count;
        }

    }
}