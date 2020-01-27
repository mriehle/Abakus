using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration.CustomUI;



[System.Runtime.InteropServices.ComVisible(true)]
public class EFRRibbon : ExcelDna.Integration.CustomUI.ExcelRibbon
{

    public override string GetCustomUI(string uiName)
    {
        string strXML = "";
        strXML += "<customUI xmlns=\'http://schemas.microsoft.com/office/2006/01/customui\' loadImage=\'GetImage\'>";
        strXML += "<ribbon>";
        strXML += "<tabs>";
        strXML += "<tab id=\'XLDataBrowser_Tab1\' label=\'Abakus BI\'>";
        strXML += "<group id=\'XLDataBrowser_Grp1\'>";
        strXML += "<button id=\'btnInsert\' label=\'User Info\' image=\'Databrowser\' size=\'large\' onAction=\'OnShowUserInfo\'" + "/>";
        strXML += "<button id=\'btnData\' label=\'Menü\' image=\'Databrowser\' size=\'large\' onAction=\'OnShowDataBrowser\'" + "/>";
        strXML += "</group>";
        strXML += "</tab>";
        strXML += "</tabs>";
        strXML += "</ribbon>";
        strXML += "</customUI>";
        return strXML;
    }


    public void OnShowUserInfo(ExcelDna.Integration.CustomUI.IRibbonControl control)
    {
        MessageBox.Show("Angemeldeter User: " + Environment.UserName);

    }


    public void OnShowDataBrowser(ExcelDna.Integration.CustomUI.IRibbonControl control)
    {
        CTPManager.ShowCTP();
    }

};

