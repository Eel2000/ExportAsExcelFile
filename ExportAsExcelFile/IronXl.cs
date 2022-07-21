using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExportAsExcelFile
{
    internal class IronXl
    {
        //        var excelBook = WorkBook.Create(ExcelFileFormat.XLSX);
        //        var sheet = excelBook.CreateWorkSheet(Assembly.GetExecutingAssembly().GetName().Name);

        //        sheet.Header.Center = $"Wekly Timesheet ({DateTime.Now.ToString("d")} au {DateTime.Now.AddDays(7).ToString("d")})";

        //sheet.Merge("A1:W1");

        //sheet["A1:W1"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
        //sheet["A1:W1"].Style.Font.Bold = false;
        //sheet["A1:W1"].Style.Font.Height = 25;
        //sheet["A1:W1"].Value = $"Wekly Timesheet ({DateTime.Now.ToString("d")} au {DateTime.Now.AddDays(7).ToString("d")})";


        //sheet["A5"].StringValue = "Employee Id";
        //sheet["A5"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.CenterSelection;
        //sheet["A5"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Bottom;
        //sheet["A5"].Style.Font.Bold = true;

        ////day col
        //sheet.Merge("C4:E4");
        //sheet["C4:E4"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
        //sheet["C4:E4"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
        //sheet["C4:E4"].Style.Font.Bold = true;  
        //sheet["C4:E4"].Style.Font.Height = 16;
        //sheet["C4:E4"].Style.SetBackgroundColor(Color.FromArgb(255, 229, 153));
        //sheet["C4:E4"].StringValue = DateTime.Now.DayOfWeek.ToString() + $" ({DateTime.Now.ToString("d")})";

        //sheet["B5"].StringValue = "Fullname";
        //sheet["B5"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
        //sheet["B5"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
        //sheet["B5"].Style.Font.Bold = true;
        //sheet["B5"].Style.SetBackgroundColor(Color.Gray);
        //sheet["B5"].Style.Font.SetColor(color: Color.DarkOrange);
        //sheet["B5"].Style.BottomBorder.SetColor(color: Color.Black);
        //sheet["B5"].Style.LeftBorder.SetColor(color: Color.Black);
        //sheet["B5"].Style.TopBorder.SetColor(color: Color.Black);
        //sheet["B5"].Style.RightBorder.SetColor(color: Color.Black);
        //sheet["B5"].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["B5"].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["B5"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["B5"].Style.RightBorder.Type = IronXL.Styles.BorderType.Thin;


        //sheet["C5"].StringValue = "Entry Time";
        //sheet["C5"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
        //sheet["C5"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
        //sheet["C5"].Style.Font.Bold = true;
        //sheet["C5"].Style.SetBackgroundColor(Color.Gray);
        //sheet["C5"].Style.Font.SetColor(color: Color.DarkOrange);
        //sheet["C5"].Style.BottomBorder.SetColor(color: Color.Black);
        //sheet["C5"].Style.LeftBorder.SetColor(color: Color.Black);
        //sheet["C5"].Style.TopBorder.SetColor(color: Color.Black);
        //sheet["C5"].Style.RightBorder.SetColor(color: Color.Black);
        //sheet["C5"].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["C5"].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["C5"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["C5"].Style.RightBorder.Type = IronXL.Styles.BorderType.Thin;


        //sheet["D5"].StringValue = "Exit Time";
        //sheet["D5"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
        //sheet["D5"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
        //sheet["D5"].Style.Font.Bold = true;
        //sheet["D5"].Style.SetBackgroundColor(Color.Gray);
        //sheet["D5"].Style.Font.SetColor(color: Color.DarkOrange);
        //sheet["D5"].Style.BottomBorder.SetColor(color: Color.Black);
        //sheet["D5"].Style.LeftBorder.SetColor(color: Color.Black);
        //sheet["D5"].Style.TopBorder.SetColor(color: Color.Black);
        //sheet["D5"].Style.RightBorder.SetColor(color: Color.Black);
        //sheet["D5"].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["D5"].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["D5"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["D5"].Style.RightBorder.Type = IronXL.Styles.BorderType.Thin;


        //sheet["E5"].StringValue = "Duration";
        //sheet["E5"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.CenterSelection;
        //sheet["E5"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
        //sheet["E5"].Style.Font.Bold = true;
        //sheet["E5"].Style.SetBackgroundColor(Color.Gray);
        //sheet["E5"].Style.Font.SetColor(color: Color.DarkOrange);
        //sheet["E5"].Style.BottomBorder.SetColor(color: Color.Black);
        //sheet["E5"].Style.LeftBorder.SetColor(color: Color.Black);
        //sheet["E5"].Style.TopBorder.SetColor(color: Color.Black);
        //sheet["E5"].Style.RightBorder.SetColor(color: Color.Black);
        //sheet["E5"].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["E5"].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["E5"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;
        //sheet["E5"].Style.RightBorder.Type = IronXL.Styles.BorderType.Thin;


        ////Adding data in cells
        //var groupedEmpList = employees.GroupBy(x => x.EntryTime.DayOfWeek, (dayOfWeek, employees) => new
        //{
        //    DayOfWeek = dayOfWeek,
        //    Employees = employees
        //});

        //        groupedEmpList = groupedEmpList.OrderByDescending(x => x.DayOfWeek);

        ////int celTitleStartIndex = 4;
        ////int cellValueStartIndex = 6;

        ////foreach (var item in groupedEmpList)
        ////{
        ////    //style
        ////    var rangeToMerge = sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"];
        ////    var lastCell = rangeToMerge.Last().Address;//get the location of the last cell to the range

        ////    //var getLastCol = lastCell.

        ////    sheet.Merge($"C{celTitleStartIndex}:E{celTitleStartIndex}");
        ////    sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
        ////    sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
        ////    sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"].Style.Font.Bold = true;
        ////    sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"].Style.Font.Height = 16;
        ////    sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"].Style.SetBackgroundColor(Color.FromArgb(255, 229, 153));

        ////    sheet[$"C{celTitleStartIndex}:E{celTitleStartIndex}"].StringValue =
        ////        item.DayOfWeek.ToString() + $" ({item.Employees.First().EntryTime.ToString("d")})";
        ////    celTitleStartIndex++;//next day
        ////}


        //excelBook.SaveAs(@"C:\Users\CTRL TECH\Documents\SideProject\Sheets\" +
        //    Assembly.GetExecutingAssembly().GetName().Name + "-" +
        //    DateTime.Now.ToString("d") + ".xlsx");

    }
}
