using Aspose.Cells;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Excel_API;

string path = @"C:\Users\Redmi\Documents\Робот_КП\KPI_list.xlsx";
//Загрузить файл Excel
Workbook workbook = new Workbook(path);

//Получить все рабочие листы
WorksheetCollection worksheets = workbook.Worksheets;  

//Перебрать все рабочие листы
for (int workSheetIndex = 0; workSheetIndex < worksheets.Count; workSheetIndex++)
{
    //Поучить рабочий лист, используя его индекс
    Worksheet worksheet = worksheets[workSheetIndex];

    //Печать имени рабочего листа
    Console.WriteLine("Worksheet: " +  worksheet.Name);

    //Поучить количество строк и столбцов
    int rows = worksheet.Cells.MaxDataRow;
    int cols = worksheet.Cells.MaxDataColumn;

    //Цикл по строкам
    for(int i = 0; i < rows; i++)
    {
        //Перебрать каждый столбец в выбранной строке
        for(int j = 0; j < cols; j++)
        {
            //Печать значение ячейки
            Console.WriteLine(worksheet.Cells[i, j].Value + " | ");
        }
        //Печать разрыв строки
        Console.WriteLine();
    }
}
string filePath = @"C:\Users\Redmi\Documents\Робот_КП\file.docx";
var helper = new WordHelper(filePath);
var item = new Dictionary<string, string> 
{ 

};

