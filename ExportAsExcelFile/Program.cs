using ExportAsExcelFile;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;

Console.WriteLine("Starting to generate file....");
const string fileLocation = @"C:\Users\CTRL TECH\Documents\SideProject\Sheets\";
var employees = new List<Employee>()
{
    new Employee(1,"MBINDA", "Jean",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(1),DateTime.Now.AddHours(1)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(1),DateTime.Now.AddHours(1)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(1),DateTime.Now.AddHours(1)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(2),DateTime.Now.AddHours(4)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(2),DateTime.Now.AddHours(4)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(2),DateTime.Now.AddHours(4)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(3),DateTime.Now.AddHours(8)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(3),DateTime.Now.AddHours(8)),
    new Employee(1,"MBINDA", "Jean",DateTime.Now.AddDays(3),DateTime.Now.AddHours(8)),
    ////next one
    new Employee(2,"KASONGO", "Eliel",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(1),DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(1),DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(1),DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(2),DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(2),DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(2),DateTime.Now.AddHours(8)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(3),DateTime.Now.AddHours(2)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(3),DateTime.Now.AddHours(2)),
    new Employee(2,"KASONGO", "Eliel",DateTime.Now.AddDays(3),DateTime.Now.AddHours(2)),
    ////next one
    new Employee(3,"NSENDA", "Claude",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now,DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(1),DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(1),DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(1),DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(2),DateTime.Now.AddHours(12)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(2),DateTime.Now.AddHours(12)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(2),DateTime.Now.AddHours(12)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(3),DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(3),DateTime.Now.AddHours(8)),
    new Employee(3,"NSENDA", "Claude",DateTime.Now.AddDays(3),DateTime.Now.AddHours(8)),
};

//ExcelPackage excel = new ExcelPackage();
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;//set license to no Commercial
using (var excel = new ExcelPackage())
{
    var sheet = excel.Workbook.Worksheets.Add(Assembly.GetExecutingAssembly().GetName().Name);

    sheet.Row(1).Height = 50;//set the height of the whole row
    sheet.Cells[1, 1, 2, 23].AutoFitColumns();
    sheet.Cells[1, 1, 2, 23].Style.Font.Size = 25;
    sheet.Cells[1, 1, 2, 23].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    sheet.Cells[1, 1, 2, 23].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    sheet.Cells[1, 1, 2, 23].Merge = true;
    sheet.Cells[1, 1, 2, 23].Value = $"Weekly Timesheet ({DateTime.Now.ToString("d")} au {DateTime.Now.AddDays(7).ToString("d")})";



    var groupedEmployeesList = employees.GroupBy(x => x.EntryTime.DayOfWeek, (dayOfWeek, employeesList) => new
    {
        Day = dayOfWeek,
        EmployeeList = employeesList
    });

    var added = new List<int>();

    int fromRowIndex = 3, fromColIndex = 3,
           toRowIndex = 3, toColIndex = 5, empIdcol = 1,
           empIdRow = 5, empStarInsertAtRow = 5,
           colInsertValueFullNameAtIndex = 2, colInsertValuesForInfo = 2;

    var currentName = string.Empty;
    var prevIndex = 0;

    foreach (var dayData in groupedEmployeesList)
    {
        //generate first the header
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].AutoFitColumns();
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Merge = true;
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Style.Font.Bold = true;
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Style.Font.Size = 16;
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Style.Fill.SetBackground(Color.FromArgb(255, 231, 163));
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        sheet.Cells[fromRowIndex, fromColIndex, toRowIndex, toColIndex].Value = $"{dayData.Day} ({dayData.EmployeeList.First().EntryTime.ToString("d")})";

        //generate entry time col
        sheet.Column(fromColIndex).Width = 20;//set the entryTime col width to 20px
        sheet.Cells[4, fromColIndex].Value = "Entry Time";
        sheet.Cells[4, fromColIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        sheet.Cells[4, fromColIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        sheet.Cells[4, fromColIndex].Style.Font.Bold = true;
        sheet.Cells[4, fromColIndex].Style.Font.Size = 16;
        sheet.Cells[4, fromColIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        sheet.Cells[4, fromColIndex].Style.Font.Color.SetColor(Color.DarkOrange);
        sheet.Cells[4, fromColIndex].Style.Fill.SetBackground(Color.FromArgb(214, 214, 214));

        //generate exit time col
        sheet.Column(fromColIndex + 1).Width = 20;//set the exitTime col width to 20 px
        sheet.Cells[4, fromColIndex + 1].Value = "Exit Time";
        sheet.Cells[4, fromColIndex + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        sheet.Cells[4, fromColIndex + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        sheet.Cells[4, fromColIndex + 1].Style.Font.Bold = true;
        sheet.Cells[4, fromColIndex + 1].Style.Font.Size = 16;
        sheet.Cells[4, fromColIndex + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        sheet.Cells[4, fromColIndex + 1].Style.Font.Color.SetColor(Color.DarkOrange);
        sheet.Cells[4, fromColIndex + 1].Style.Fill.SetBackground(Color.FromArgb(214, 214, 214));

        //generate duration col
        sheet.Column(fromColIndex + 2).Width = 20;//set the duration col width to 20px
        sheet.Cells[4, fromColIndex + 2].Value = "Duration";
        sheet.Cells[4, fromColIndex + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        sheet.Cells[4, fromColIndex + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        sheet.Cells[4, fromColIndex + 2].Style.Font.Bold = true;
        sheet.Cells[4, fromColIndex + 2].Style.Font.Size = 16;
        sheet.Cells[4, fromColIndex + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        sheet.Cells[4, fromColIndex + 2].Style.Font.Color.SetColor(Color.DarkOrange);
        sheet.Cells[4, fromColIndex + 2].Style.Fill.SetBackground(Color.FromArgb(214, 214, 214));


        var groupedById = dayData.EmployeeList.GroupBy(x => x.ID, (empId, data) => new
        {
            EmpId = empId,
            Data = data
        });


        foreach (var emp in groupedById)
        {
            foreach (var item in emp.Data)
            {
                //write his ID
                sheet.Cells[empIdRow, empIdcol].Value = item.ID;
                sheet.Cells[empIdRow, empIdcol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[empIdRow, empIdcol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[empIdRow, empIdcol].Style.Font.Size = 16;

                //write his fullname
                sheet.Cells[empStarInsertAtRow, colInsertValueFullNameAtIndex].Value = $"{item.Name} {item.FirstName}";
                sheet.Cells[empStarInsertAtRow, colInsertValueFullNameAtIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValueFullNameAtIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValueFullNameAtIndex].Style.Font.Size = 16;
                sheet.Cells[empStarInsertAtRow, colInsertValueFullNameAtIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //write his entry time
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 1].Value = item.EntryTime.ToString("t");
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 1].Style.Font.Size = 16;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //write his exit time
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 2].Value = item.ExitTime.ToString("t");
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 2].Style.Font.Size = 16;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //write his exit time
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 3].Value = (item.ExitTime.Hour - item.EntryTime.Hour) + (item.ExitTime.Minute - item.EntryTime.Minute);
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 3].Style.Font.Size = 16;
                sheet.Cells[empStarInsertAtRow, colInsertValuesForInfo + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
            empIdRow++;
            empStarInsertAtRow++;
        }

        fromColIndex += 3;//skip 3 col to go to next day scope
        toColIndex += 3;//skip 3 col to go to next day scope
        colInsertValuesForInfo = fromColIndex - 1;//ajust the line to the previous one 
        empIdRow = 5;//re-initialize the index to the first row means 5 each time we switch the day
        empStarInsertAtRow = 5; //re-initialize the index to the first row means 5 each time we switch the day
    }


    sheet.Row(4).Height = 40;
    sheet.Column(1).Width = 20;
    sheet.Cells[4, 1].Value = "Employee Id";
    sheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    sheet.Cells[4, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
    sheet.Cells[4, 1].Style.Font.Size = 16;

    ////To be generated
    sheet.Row(3).Height = 40;//set the height of the whole row

    sheet.Column(2).Width = 20;
    //sheet.Cells[4, 2].AutoFitColumns();
    sheet.Cells[4, 2].Value = "Fullname";
    sheet.Cells[4, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    sheet.Cells[4, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    sheet.Cells[4, 2].Style.Font.Bold = true;
    sheet.Cells[4, 2].Style.Font.Size = 16;
    sheet.Cells[4, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
    sheet.Cells[4, 2].Style.Font.Color.SetColor(Color.DarkOrange);
    sheet.Cells[4, 2].Style.Fill.SetBackground(Color.FromArgb(214, 214, 214));




    var name = fileLocation + Assembly.GetExecutingAssembly().GetName().Name + "-" +
                DateTime.Now.ToString("d") + ".xlsx";
;
    var p = excel.GetAsByteArray();
    File.WriteAllBytes(name, p);

    var proce = new Process();
    proce.StartInfo = new ProcessStartInfo
    {
        UseShellExecute = true,
        FileName = name
    };
    proce.Start();
    Console.WriteLine("Operation finished");
}