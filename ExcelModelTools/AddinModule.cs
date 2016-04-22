using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;
using AddinExpress.MSO;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelModelTools
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("EFE35A0C-6AF1-4419-BFB3-9D591403706A"), ProgId("ExcelModelTools.AddinModule")]
    public partial class AddinModule : AddinExpress.MSO.ADXAddinModule
    {

        public FontDialog inputFontDialog = new FontDialog();
        public static Color inputDefaultFontColor = Color.FromArgb(1, 1, 1, 255);
        public static Color inputDefaultFillColor = Color.FromArgb(1, 255, 255, 204);
        public static Color assumptionDefaultFontColor = Color.FromArgb(1, 0, 176, 80);
        public int numRowsToInsert;
        public int numColsToInsert;
        public int selectPrecedentsBttnCounter;
        public double rowResizeValue;
        public double colResizeValue;
        public Excel.Range rootCell;
        public Excel.Range[] precidentArray;
        public Excel.Range rangeStore1;   //this gets used to store a cell's format for the autohighlight function
        public Excel.Range rangeStore2;
        public Excel.Range[] cellArray1;
        public Excel.Range[] cellArray2;
        public Color colorStore1;
        public Color colorStore2;
        public Color[] colorStoreArray1;
        public Color[] colorStoreArray2;
        public double tintStore1;
        public Excel.Borders borderStore1;
        public bool autoFormHighlightBool;  //true when auto-highlighting has been turned on, false otherwise
        public bool autoFormHighlightFirstSelectionBool; //use this to keep track whether auto-highlighting was just pressed.  this is important for keeping track of a cell's formatting with the auto-highlight feature





        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler

            inputFontDialog.ShowApply = true;
            inputFontDialog.ShowColor = true;
            inputFontDialog.ShowEffects = true;

           

        }
 
        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance 
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

        private void inputFontButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;   //create range object
            Excel.Borders selectedBorders = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;   //set range object to current selection
                selectedBorders = selectedRange.Borders;

                selectedRange.Font.Color = inputDefaultFontColor;
                selectedRange.Interior.Color = inputDefaultFillColor;

                selectedBorders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                selectedBorders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                selectedBorders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                selectedBorders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                selectedBorders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                selectedBorders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;

                selectedBorders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedBorders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedBorders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedBorders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedBorders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedBorders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);     //unlink selected Range object from the Excel application
                if (selectedBorders != null) Marshal.ReleaseComObject(selectedBorders);
            }
        }

        private void inputFontPickerButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            inputFontDialog.ShowDialog();
        }

        private void hardCodeButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;   //create range object

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;   //set range object to current selection
                selectedRange.Font.Color = inputDefaultFontColor;
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);     //unlink selected Range object from the Excel application
            }

        }

        private void assumptionButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;   //create range object

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;   //set range object to current selection
                selectedRange.Font.Color = assumptionDefaultFontColor;
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);     //unlink selected Range object from the Excel application
            }
        }

        private void insertRowsComboBox_OnChange(object sender, IRibbonControl Control, string text)
        {
            if (Int32.TryParse(text, out numRowsToInsert))
            {
                numRowsToInsert = Int32.Parse(text);
            }
            else
            {
                numRowsToInsert = 0;
            }
        }

        private void insertRowsButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;   //create range object
            Excel.Range endRange = null;
            Excel.Range insertRange = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;   //set range object to current selection
                endRange = selectedRange.Offset[numRowsToInsert-1, 0];  //we need to subtract 1 from num rows to insert
                insertRange = ExcelApp.Range[selectedRange, endRange];
                insertRange.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);     //unlink selected Range object from the Excel application
                if (endRange != null) Marshal.ReleaseComObject(endRange);
                if (insertRange != null) Marshal.ReleaseComObject(insertRange);
            }
        }

        private void insertColsComboBox_OnChange(object sender, IRibbonControl Control, string text)
        {
            if (Int32.TryParse(text, out numColsToInsert))
            {
                numColsToInsert = Int32.Parse(text);
            }
            else
            {
                numColsToInsert = 0;
            }
        }

        private void insertColsButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;   //create range object
            Excel.Range endRange = null;
            Excel.Range insertRange = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;   //set range object to current selection
                endRange = selectedRange.Offset[0, numColsToInsert-1];  //we need to subtract 1 from num rows to insert
                insertRange = ExcelApp.Range[selectedRange, endRange];
                insertRange.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);     //unlink selected Range object from the Excel application
                if (endRange != null) Marshal.ReleaseComObject(endRange);
                if (insertRange != null) Marshal.ReleaseComObject(insertRange);
            }
        }

        private void rowResizeButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;
                selectedRange.EntireRow.RowHeight = rowResizeValue;
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
            }
        }

        private void colResizeButton_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;
                selectedRange.EntireColumn.ColumnWidth = colResizeValue;
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
            }
        }

        private void resizeRowComboBox_OnChange(object sender, IRibbonControl Control, string text)
        {
            if(Double.TryParse(text, out rowResizeValue))
            {
                rowResizeValue = Double.Parse(text);
            }
            else
            {
                rowResizeValue = 0;
            }
        }

        private void resizeColComboBox_OnChange(object sender, IRibbonControl Control, string text)
        {
            if (Double.TryParse(text, out colResizeValue))
            {
                colResizeValue = Double.Parse(text);
            }
            else
            {
                colResizeValue = 0;
            }
        }

        private void pasteValueKeyboardShortcut_Action(object sender)
        {
            Excel.Range selectedRange = null;
            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;
                if (Clipboard.ContainsText())
                { 
                    selectedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                }
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
            }
        }

        private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
        {
            numRowsToInsert = 1;    //needs to match Item[0] of the combobox
            numColsToInsert = 1;
            rowResizeValue = 1;
            colResizeValue = 1;
            this.insertRowsComboBox.Text = this.insertRowsComboBox.Items[0].AsRibbonItem.Caption;
            this.insertColsComboBox.Text = this.insertColsComboBox.Items[0].AsRibbonItem.Caption;
            this.resizeRowComboBox.Text = this.resizeRowComboBox.Items[0].AsRibbonItem.Caption;
            this.resizeColComboBox.Text = this.resizeColComboBox.Items[0].AsRibbonItem.Caption;
            this.rootCell = null;
            selectPrecedentsBttnCounter = 1;
            autoFormHighlightBool = false;
            autoFormHighlightFirstSelectionBool = false;
            this.rangeStore1 = null;
            this.rangeStore2 = null;
        }

        private void selectPrecedentsBttn_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;
            Excel.Range precedentsRange = null;

            //first we try to set the precedents range equal to the rootCell.Precendents, otherwise, we set to null
            //NOTE: WE CANNOT SIMPLY TEST FOR NULL VALUE, IF NO PRECEDENTS EXIST THE rootCell.Precedents PROPERTY CONTAINS
            //AN EXCEPTION, NOT A NULL VALUE.

            try
            {
               precedentsRange = rootCell.DirectPrecedents ?? null;
            }
            catch
            {
                precedentsRange = null;
            }
                       

            //If rootCell exists, and it contains precendents, then cycle through the cells in the precedent array.
            try
            {
                    if (rootCell != null && precedentsRange != null && Convert.ToBoolean(precedentsRange.HasArray) == false)
                    {
                        //precedentsRange = rootCell.DirectPrecedents;
                        selectedRange = precidentArray[selectPrecedentsBttnCounter-1];
                        selectedRange.Select();

                        int test = selectPrecedentsBttnCounter + 1;

                    if (test > precedentsRange.Count)
                    {
                        selectPrecedentsBttnCounter = 1;
                    }
                    else
                    {
                        selectPrecedentsBttnCounter = test;
                    }
                }
            }
            
            finally
            {
                //if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);   //not sure why needs to be commented out
                if (precedentsRange != null) Marshal.ReleaseComObject(precedentsRange);
            }
        }

        private void setRootCellBttn_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Excel.Range selectedRange = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;
                if (selectedRange.Count == 1)
                {
                    rootCell = selectedRange;
                    rootCellLabel.Caption = rootCell.Address;
                    selectPrecedentsBttnCounter = 1;

                    if (rootCell.DirectPrecedents != null && rootCell.DirectPrecedents.Count > 0)
                    {

                        precidentArray = new Excel.Range[rootCell.DirectPrecedents.Count];
                        int i = 0;
                        foreach (Excel.Range rng in rootCell.DirectPrecedents.Cells)
                        {
                            precidentArray[i] = rng;
                            i++;
                            Debug.WriteLine(rng.Address);
                        }
                    }
                }
                
            }
            catch
            {

            }
        }

        private void rootCellLabel_PropertyChanging(object sender, ADXRibbonPropertyChangingEventArgs e)
        {
           
        }

        private void returnToRootCellBttn_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (rootCell != null)
            {
                rootCell.Select();
                selectPrecedentsBttnCounter = 1;
            }
        }

        private void formulaHighlightBttn_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (formulaHighlightBttn.Caption == "Turn On Formula Highlighting")
            {
                autoFormHighlightBool = true;
                autoFormHighlightFirstSelectionBool = true;
                formulaHighlightBttn.Caption = "Turn Off Formula Highlighting";
            }
            else
            {
                autoFormHighlightBool = false;
                formulaHighlightBttn.Caption = "Turn On Formula Highlighting";
            }
        }

        private void adxExcelAppEvents1_SheetBeforeRightClick(object sender, ADXExcelSheetBeforeEventArgs e)
        {
   
        
        }

        private void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)
        {
            Excel.Range selectedRange = null;
            Excel.Range precedentsRange = null;
            //Excel.Range prevSelectedRange = null;

            if (autoFormHighlightBool == true)
            {

                autoFormHighlightBool = false;

                try
                {
                    selectedRange = range as Excel.Range;
                    Debug.Print("SELECTED RANGE:");
                    Debug.Print(selectedRange.Address);
                    precedentsRange = selectedRange.DirectPrecedents;
                    rangeStore2 = selectedRange.DirectPrecedents;
                    cellArray1 = new Excel.Range[selectedRange.DirectPrecedents.Count];
                    int i=0;
                    foreach (Excel.Range rng in selectedRange.DirectPrecedents.Cells)
                    {
                        cellArray1[i] = rng;
                        i++;
                        Debug.WriteLine(rng.Address);
                    }

                    //selectedRange = ExcelApp.Selection as Excel.Range;
                    //precedentsRange = rootCell.DirectPrecedents ?? null;
                }
                catch
                {
                    precedentsRange = null;
                    rangeStore2 = null;
                }

                try
                {

                    if (selectedRange != null && precedentsRange != null)
                    {
                        if (autoFormHighlightFirstSelectionBool == true)
                        {
                            autoFormHighlightFirstSelectionBool = false;    //this bool should be true only for the first selection
                        }
                        else
                        {
                            try
                            {
                                for (int i=0; i<cellArray2.Length; i++)
                                {
                                    if (colorStoreArray1[i] == Color.White)
                                    {
                                        cellArray2[i].Interior.Color = -4142;
                                    }
                                    else
                                    {
                                        cellArray2[i].Interior.Color = colorStoreArray1[i];
                                    }
                                }
                            }
                            catch
                            {

                            }

                        }

                        colorStoreArray1 = new Color[cellArray1.Length];
                        for (int i = 0; i < cellArray1.Length; i++)
                        {
                            colorStoreArray1[i] = ColorTranslator.FromOle((int)((double)cellArray1[i].Interior.Color));
                        }

                        colorStore1 = ColorTranslator.FromOle((int)((double)precedentsRange.Interior.Color)); //interior color first cast as double, then int, then converted to System.Color                                                                       
                        colorStore2 = ColorTranslator.FromOle((int)((double)precedentsRange.Borders.Color));
                        tintStore1 = Convert.ToDouble(precedentsRange.Interior.TintAndShade);
                        borderStore1 = precedentsRange.Borders;
                        rangeStore1 = rangeStore2;
                        cellArray2 = new Excel.Range[cellArray1.Length];
                        Array.Copy(cellArray1, cellArray2, cellArray1.Length);
                        precedentsRange.Interior.Color = Color.FromArgb(1, 1, 1, 255);
                        precedentsRange.Interior.TintAndShade = 0.7;

                    }
                    else
                    {
                        for (int i = 0; i < cellArray2.Length; i++)
                        {
                            if (colorStoreArray1[i] == Color.White)
                            {
                                cellArray2[i].Interior.Color = -4142;
                            }
                            else
                            {
                                cellArray2[i].Interior.Color = colorStoreArray1[i];
                            }
                        }
                    }

                }
                catch
                {

                }
                finally
                {
                    if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);   //not sure why needs to be commented out
                    if (precedentsRange != null) Marshal.ReleaseComObject(precedentsRange);
                }

                autoFormHighlightBool = true;
            }
            else
            {
                if(cellArray2 != null && cellArray2.Length > 0)
                {
                    for (int i = 0; i < cellArray2.Length; i++)
                    {
                        if (colorStoreArray1[i] == Color.White)
                        {
                            cellArray2[i].Interior.Color = -4142;
                        }
                        else
                        {
                            cellArray2[i].Interior.Color = colorStoreArray1[i];
                        }
                    }
                }
            }
        }

        private void cellFormatDropDown_OnAction(object sender, IRibbonControl Control, string selectedId, int selectedIndex)
        {
            Excel.Range selectedRange = null;

            try
            {
                selectedRange = ExcelApp.Selection as Excel.Range;
                //selectedRange.Value = "TEST";

                switch (selectedIndex)
                {
                    case 0:
                        //cents
                        selectedRange.NumberFormat = "_(#,##0.00_);_((#,##0.00);_(\"-\"??_);_(@_)";
                        break;
                    case 1:
                        //dollars
                        selectedRange.NumberFormat = "_(#,##0_);_((#,##0);_(\"-\"??_);_(@_)";
                        break;
                    case 2:
                        //thousands
                        selectedRange.NumberFormat = "_(#,##0,_);_((#,##0,);_(\"-\"??_);_(@_)";
                        break;
                    case 3:
                        //millions
                        selectedRange.NumberFormat = "_(#,##0,,_);_((#,##0,,);_(\"-\"??_);_(@_)";
                        break;
                    case 4:
                        //billions
                        selectedRange.NumberFormat = "_(#,##0,,,_);_((#,##0,,,);_(\"-\"??_);_(@_)";
                        break;
                    default:
                        //default is cents
                        selectedRange.NumberFormat = "_(#,##0.00_);_((#,##0.00);_(\"-\"??_);_(@_)";
                        break;
                         
                }
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
            }
        }
    }
}

