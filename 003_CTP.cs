using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration;
using System;
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

[ComVisible(true)]
public class XLDNA_CTP_Control : UserControl
{
    private DateTimePicker dateTimePicker1;
    private Label label_Header;


    public XLDNA_CTP_Control()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(3, 132);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 22);
            this.dateTimePicker1.TabIndex = 0;
            // 
            // XLDNA_CTP_Control
            // 
            this.Controls.Add(this.dateTimePicker1);
            this.Name = "XLDNA_CTP_Control";
            this.Size = new System.Drawing.Size(212, 669);
            this.Load += new System.EventHandler(this.XLDNA_CTP_Control_Load);
            this.ResumeLayout(false);

    }

    private void XLDNA_CTP_Control_Load(object sender, EventArgs e)
    {

    }
}


/////////////// Helper class to manage CTP ///////////////////////////
internal static class CTPManager
{
    static CustomTaskPane ctp;

    public static void ShowCTP()
    {
        if (ctp == null)
        {
            // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
            ctp = ExcelDna.Integration.CustomUI.CustomTaskPaneFactory.CreateCustomTaskPane(typeof(XLDNA_CTP_Control), "Abakus");
            ctp.Visible = false;
            ctp.DockPosition = ExcelDna.Integration.CustomUI.MsoCTPDockPosition.msoCTPDockPositionLeft;
            ctp.VisibleStateChange += ctp_VisibleStateChange;
            ctp.Width = 250;
            ctp.Visible = true;

        }
        else
        {
            //Delete when open
            DeleteCTP();

        }
    }


    public static void DeleteCTP()
    {
        if (ctp != null)
        {
            // Could hide instead, by calling ctp.Visible = false;
            ctp.Delete();
            ctp = null;
        }
    }

    static void ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
    {

        // Wenn Sichtbar auf False dann lösche die CTP

        //MessageBox.Show("Visibility changed to " + CustomTaskPaneInst.Visible);
        bool openCTP = CustomTaskPaneInst.Visible;

        if (openCTP == false)
        {
            ctp.Delete();
        }
    }

}