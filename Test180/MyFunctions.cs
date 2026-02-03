using ExcelDna.Integration;

namespace Test180
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function (from 180)")]
        public static string SayHello(string name)
        {
            return "Hello from 180 " + name;
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
