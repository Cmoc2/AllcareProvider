using System;
using System.IO;
using Microsoft.VisualBasic.FileIO;
//using ClosedXML.Excel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

string directoryPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\xlsInput\");
//output directory path
Console.WriteLine($"Monitoring directory: {directoryPath}");
//set output directory path
string outputDirectoryPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\xlsOutput\");
Console.WriteLine($"Output directory: {outputDirectoryPath}");

//if directories do not exist, create it:
//input Directory
if (!Directory.Exists(directoryPath))
{
    Directory.CreateDirectory(directoryPath);
}

//output Directory
if (!Directory.Exists(outputDirectoryPath))
{
    Directory.CreateDirectory(outputDirectoryPath);
}

//check if there are xls files in the input directory beginning with 'Visit Report -', ignore everything else
if (Directory.GetFiles(directoryPath, "Visit Report -*.xls").Length > 0)
{
    //process the files in the directory
    string[] files = Directory.GetFiles(directoryPath, "Visit Report -*.xls");
    foreach (var file in files)
    {
        //Only the filename, not the full path
        Console.WriteLine($"Found file: {Path.GetFileName(file)}");

        //Open the xls file and save it as xlsx in the output directory.
        //Send the original xls file to recycle bin after conversion.
        string outputFilePath = Path.Combine(outputDirectoryPath, Path.GetFileNameWithoutExtension(file) + ".xlsx");
        using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read))
        {
            var hssfWorkbook = new HSSFWorkbook(fileStream);
            var xssfWorkbook = new XSSFWorkbook();
            for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
            {
                var sheet = hssfWorkbook.GetSheetAt(i);
                var newSheet = xssfWorkbook.CreateSheet(sheet.SheetName);
                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    var oldRow = sheet.GetRow(row);
                    var newRow = newSheet.CreateRow(row);
                    if (oldRow != null)
                    {
                        for (int col = 0; col < oldRow.LastCellNum; col++)
                        {
                            var oldCell = oldRow.GetCell(col);
                            var newCell = newRow.CreateCell(col);
                            if (oldCell != null)
                            {
                                newCell.SetCellValue(oldCell.ToString());
                            }
                        }
                    }
                }
            }
            using (var outputStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
            {
                xssfWorkbook.Write(outputStream);
            }
        }
        Console.WriteLine($"Converted and saved file: {Path.GetFileName(outputFilePath)}");
        //Delete the original xls file after conversion.
        FileSystem.DeleteFile(file, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
    }
}
else
{
    Console.WriteLine("No file found. Exiting program.\nPress any key to exit.");
    Console.ReadKey();
    Environment.Exit(0);
}
