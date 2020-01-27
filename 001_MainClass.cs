
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;


// MY ADDIN als Referenz für VBA
public class MyAddin : IExcelAddIn
{
    public static Excel.Application App = (Excel.Application)ExcelDnaUtil.Application;
    public void AutoOpen()
    {
        ExcelDna.ComInterop.ComServer.DllRegisterServer(); //this is needed to expose exceldna to vba
    }
    public void AutoClose()
    {
        ExcelDna.ComInterop.ComServer.DllUnregisterServer(); //this is needed to expose exceldna to vba
    }
};


// Global Variables
public static class Globals
{
    public static string Nt_Account = Environment.UserName;
}