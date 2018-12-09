# C-Sharp-Data-Exporting-Excel-Word

List<Person> people = new List<Person>();

for(int i = 1; i < 100; ++i) people.Add(new Person { Name = $"Person Lastname {i}" });

List<Foo> foos = new List<Foo>()
{
    new Foo
    {
        Name = "Marcus",
        Marks = 94.6f
    },
    new Foo
    {
        Name = "Steve",
        Marks = 87f
    },
    new Foo
    {
        Name = "Paul",
        Marks = 92.10f
    },
    new Foo
    {
        Name = "Anna",
        Marks = 94.87f
    }
};

//using (ExcelWriter excelWriter = ExcelWriter.LoadExcel(@"C:\Users\User\Documents\Sample.xls"))
//{
//    excelWriter
//        .NewRow()
//        .AppendText("Edited")
//        .AppendText("Edited 2")
//        .SkipRows(10)
//        .ResetColumn()
//        .AppendText("safasfasffsa")
//        .Write();
//}
DateTime cache = DateTime.Now;

List<dynamic> fooDynamic = new List<dynamic>()
{
    new { Name = "assfasfasfsgaga", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "x", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "t", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "vul", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "ism", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "tel", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "cosco", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "jfc", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "mrc", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "xc ncv", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "pgold", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "smdc", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "smc", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "assfasfasfsgaga", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "assfasfasfsgaga", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
    new { Name = "assfasfasfsgaga", SomeProp = SomeEnum.SomeValue1, Birthdate = cache },
};
    
using (ExcelWriter excelWriter = new ExcelWriter())
{
    string path = @"Sample.xls"; //Default saves to My Documents

    excelWriter.DocumentRowOffset = 2; //Skips 2 rows in the document
    excelWriter.DocumentColumnOffset = 3; //Skips 3 columns in the document

    excelWriter.CellStyle = new ExcelWriter.PredefinedWriterStyle("Check Cell");

    ExcelWriter.TableStyle tableStyle = new ExcelWriter.TableStyle();
    tableStyle.TitleStyle = new ExcelWriter.BasicWriterStyle(fontName: string.Empty, fontSize: 18, bold: true, halign: HorizontalAlignment.Middle, cellHeight: 45);
    tableStyle.FieldNameStyle = new ExcelWriter.BasicWriterStyle(string.Empty, 12, true, halign: HorizontalAlignment.Left);
    //tableStyle.BodyStyle = new ExcelWriter.BasicWriterStyle(
    //    "Arial",
    //    8,
    //    true,
    //    true,
    //    true,
    //    true,
    //    0,
    //    0,
    //    HorizontalAlignment.Left,
    //    VerticalAlignment.Top,
    //    Color.Chartreuse,
    //    Color.DarkBlue);

                

    tableStyle.BodyStyle = new ExcelWriter.StripedRowWriterStyle(new Color[] { Color.Red, Color.RoyalBlue, Color.SeaGreen, Color.Snow });

    excelWriter.DataTableStyle = tableStyle;

    excelWriter
        .AppendTable(fooDynamic)
        .AppendTable(people, "People of 2018")
        .SkipRows(3)
        .AppendText("Count: " + people.Count)
        .NewRow()
        .AppendText("End")
        .AppendText("Column2")
        .NewRow()
        .SkipColumns(4)
        .AppendTable(foos, "2nd Table")
        .SetCurrentSheetName("Sample Tables")
        .NewSheet("The new sheet")
        .AppendText("New sheet Column 1")
        .AppendText("New sheet Column 2")
        .ResetColumn()
        .NewRow()
        .AppendChart(
            new Dictionary<string, double>
            {
                { "January", 200 },
                { "February", 324.88 },
                { "March", 624 }
            },
            250, 130, "Sales")
        .AppendText("Sales for 1st Quarter")
        .NewSheet("Grades")
        .Write(path);
}
