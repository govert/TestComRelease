using ExcelDna.Integration;

namespace Test190
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function (from 190)")]
        public static string SayHello(string name)
        {
            return "Hello from 190 " + name;
        }

        public static void Macro1()
        {
            XlCall.Excel(XlCall.xlcAlert, "XLL Message 1", 2);
        }

        [ExcelCommand(MenuName = "Macros", MenuText = "Message")]
        public static void Macro2()
        {
            XlCall.Excel(XlCall.xlcAlert, "XLL Message 2", 2);
        }
    }
}
