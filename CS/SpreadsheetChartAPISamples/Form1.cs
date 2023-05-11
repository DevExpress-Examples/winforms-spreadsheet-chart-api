using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraRichEdit;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraTab;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;


namespace SpreadsheetChartAPISamples
{
    public class Form1 : DevExpress.XtraEditors.XtraForm
    {
        CultureInfo defaultCulture = new CultureInfo("en-US");
        //IContainer components = null;

        #region Controls
        private SpreadsheetControl spreadsheet;
        private TreeList treeList1;
        private XtraTabControl xtraTabControl1;
        private XtraTabPage xtraTabPage1;
        private RichEditControl richEditControlCS;
        private XtraTabPage xtraTabPage2;
        private RichEditControl richEditControlVB;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.SimpleLabelItem codeExampleNameLbl;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.SplitterItem splitterItem1;
        private DevExpress.XtraLayout.SplitterItem splitterItem2;
        #endregion
        ExampleCodeEditor codeEditor;
        ExampleEvaluatorByTimer evaluator;
        List<CodeExampleGroup> examples;
        bool treeListRootNodeLoading = true;

        public Form1()
        {
            InitializeComponent();
            string examplePath = CodeExampleDemoUtils.GetExamplePath("CodeExamples");

            //string examplePath = "D:\\VB\\CS\\SpreadsheetMainDemo\\CodeExamples";
            Dictionary<string, FileInfo> examplesCS = CodeExampleDemoUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp);
            Dictionary<string, FileInfo> examplesVB = CodeExampleDemoUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.VB);
            DisableTabs(examplesCS.Count, examplesVB.Count);
            this.examples = CodeExampleDemoUtils.FindExamples(examplePath, examplesCS, examplesVB);
            RearrangeExamples();
            MergeGroups();
            ShowExamplesInTreeList(treeList1, examples);

            this.codeEditor = new ExampleCodeEditor(richEditControlCS, richEditControlVB);
            CurrentExampleLanguage = CodeExampleDemoUtils.DetectExampleLanguage("SpreadsheetChartAPISamples");
            this.evaluator = new SpreadsheetExampleEvaluatorByTimer(); //this.components

            this.evaluator.QueryEvaluate += OnExampleEvaluatorQueryEvaluate;
            this.evaluator.OnBeforeCompile += evaluator_OnBeforeCompile;
            this.evaluator.OnAfterCompile += evaluator_OnAfterCompile;

            ShowFirstExample();
            this.xtraTabControl1.SelectedPageChanged += new TabPageChangedEventHandler(this.xtraTabControl1_SelectedPageChanged);
        }

        private void MergeGroups()
        {
            var uniqueNameGroup = new Dictionary<string, CodeExampleGroup>();
            foreach (CodeExampleGroup n in examples)
                if (uniqueNameGroup.ContainsKey(n.Name))
                {
                    uniqueNameGroup[n.Name].Merge(n);
                }
                else
                {
                    uniqueNameGroup[n.Name] = n;
                }

            examples.Clear();
            foreach (var value in uniqueNameGroup.Values)
                examples.Add(value);
        }

        void RearrangeExamples()
        {
            for (int i = 0; i < examples.Count; i++)
            {
                CodeExampleGroup group = examples[i];
                if (group.Name == "Charts")
                {
                    examples.RemoveAt(i);
                    examples.Insert(0, group);
                    break;
                }
            }
            for (int i = 0; i < examples.Count; i++)
            {
                CodeExampleGroup group = examples[i];
                if (group.Name.StartsWith("Creation"))
                {
                    examples.RemoveAt(i);
                    examples.Insert(1, group);
                    break;
                }
            }
        }

        void evaluator_OnAfterCompile(object sender, OnAfterCompileEventArgs args)
        {
            IWorkbook workbook = spreadsheet.Document;
            foreach (Worksheet sheet in workbook.Worksheets)
                sheet.PrintOptions.PrintGridlines = true;

            Worksheet active = workbook.Worksheets.ActiveWorksheet;
            if (active != null) {
                CellRange usedRange =active.GetUsedRange();
                active.SelectedCell = usedRange[usedRange.RowCount * usedRange.ColumnCount - 1].Offset(1, 1);
            }
            codeEditor.AfterCompile(args.Result);
            spreadsheet.EndUpdate();
        }

        void evaluator_OnBeforeCompile(object sender, EventArgs e)
        {
            spreadsheet.BeginUpdate();
            codeEditor.BeforeCompile();

            IWorkbook workbook = spreadsheet.Document;
            workbook.Options.Culture = defaultCulture;
            bool loaded = workbook.LoadDocument("Document.xlsx");
            Debug.Assert(loaded);
        }
        ExampleLanguage CurrentExampleLanguage
        {
            get { return (ExampleLanguage)xtraTabControl1.SelectedTabPageIndex; }
            set
            {
                this.codeEditor.CurrentExampleLanguage = value;
                xtraTabControl1.SelectedTabPageIndex = (value == ExampleLanguage.Csharp) ? 0 : 1;
            }
        }
        void ShowExamplesInTreeList(TreeList treeList, List<CodeExampleGroup> examples)
        {
            #region InitializeTreeList
            treeList.OptionsPrint.UsePrintStyles = true;
            treeList.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.OnNewExampleSelected);
            treeList.OptionsView.ShowColumns = false;
            treeList.OptionsView.ShowIndicator = false;


            treeList.VirtualTreeGetChildNodes += treeList_VirtualTreeGetChildNodes;
            treeList.VirtualTreeGetCellValue += treeList_VirtualTreeGetCellValue;
            #endregion
            TreeListColumn col1 = new TreeListColumn();
            col1.VisibleIndex = 0;
            col1.OptionsColumn.AllowEdit = false;
            col1.OptionsColumn.AllowMove = false;
            col1.OptionsColumn.ReadOnly = true;
            treeList.Columns.AddRange(new TreeListColumn[] { col1 });

            treeList.DataSource = new Object();
            treeList.ExpandAll();
        }

        void treeList_VirtualTreeGetCellValue(object sender, VirtualTreeGetCellValueInfo args)
        {
            CodeExampleGroup group = args.Node as CodeExampleGroup;
            if (group != null)
                args.CellData = group.Name;

            CodeExample example = args.Node as CodeExample;
            if (example != null)
                args.CellData = example.RegionName;
        }

        void treeList_VirtualTreeGetChildNodes(object sender, VirtualTreeGetChildNodesInfo args)
        {
            if (treeListRootNodeLoading)
            {
                args.Children = examples;
                treeListRootNodeLoading = false;
            }
            else
            {
                if (args.Node == null)
                    return;
                CodeExampleGroup group = args.Node as CodeExampleGroup;
                if (group != null)
                    args.Children = group.Examples;
            }
        }
        void ShowFirstExample()
        {
            treeList1.ExpandAll();
            if (treeList1.Nodes.Count > 0)
                treeList1.FocusedNode = treeList1.MoveFirst().FirstNode;
        }
        void OnNewExampleSelected(object sender, FocusedNodeChangedEventArgs e)
        {
            CodeExample newExample = (sender as TreeList).GetDataRecordByNode(e.Node) as CodeExample;
            CodeExample oldExample = (sender as TreeList).GetDataRecordByNode(e.OldNode) as CodeExample;

            if (newExample == null)
                return;

            string exampleCode = codeEditor.ShowExample(oldExample, newExample);
            codeExampleNameLbl.Text = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(newExample.RegionName) + " example";
            CodeEvaluationEventArgs args = new CodeEvaluationEventArgs();
            InitializeCodeEvaluationEventArgs(args, newExample.RegionName);
            evaluator.ForceCompile(args);
        }
        void InitializeCodeEvaluationEventArgs(CodeEvaluationEventArgs e, string regionName)
        {
            e.Result = true;
            e.Code = codeEditor.CurrentCodeEditor.Text;
            e.Language = CurrentExampleLanguage;
            e.EvaluationParameter = spreadsheet.Document;
            e.RegionName = regionName;
        }
        void OnExampleEvaluatorQueryEvaluate(object sender, CodeEvaluationEventArgs e)
        {
            e.Result = false;
            if (codeEditor.RichEditTextChanged)
            {// && compileComplete) {
                TimeSpan span = DateTime.Now - codeEditor.LastExampleCodeModifiedTime;

                if (span < TimeSpan.FromMilliseconds(1000))
                {//CompileTimeIntervalInMilliseconds  1900
                    codeEditor.ResetLastExampleModifiedTime();
                    return;
                }
                //e.Result = true;
                InitializeCodeEvaluationEventArgs(e, e.RegionName);
            }
        }
        #region InitializeComponent
        private void InitializeComponent()
        {
            this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
            this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
            this.richEditControlCS = new DevExpress.XtraRichEdit.RichEditControl();
            this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
            this.richEditControlVB = new DevExpress.XtraRichEdit.RichEditControl();
            this.spreadsheet = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.treeList1 = new DevExpress.XtraTreeList.TreeList();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.codeExampleNameLbl = new DevExpress.XtraLayout.SimpleLabelItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.splitterItem1 = new DevExpress.XtraLayout.SplitterItem();
            this.splitterItem2 = new DevExpress.XtraLayout.SplitterItem();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
            this.xtraTabControl1.SuspendLayout();
            this.xtraTabPage1.SuspendLayout();
            this.xtraTabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.codeExampleNameLbl)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitterItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitterItem2)).BeginInit();
            this.SuspendLayout();
            // 
            // xtraTabControl1
            // 
            this.xtraTabControl1.AppearancePage.PageClient.BackColor = System.Drawing.Color.Transparent;
            this.xtraTabControl1.AppearancePage.PageClient.BackColor2 = System.Drawing.Color.Transparent;
            this.xtraTabControl1.AppearancePage.PageClient.BorderColor = System.Drawing.Color.Transparent;
            this.xtraTabControl1.AppearancePage.PageClient.Options.UseBackColor = true;
            this.xtraTabControl1.AppearancePage.PageClient.Options.UseBorderColor = true;
            this.xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.True;
            this.xtraTabControl1.Location = new System.Drawing.Point(12, 48);
            this.xtraTabControl1.Name = "xtraTabControl1";
            this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
            this.xtraTabControl1.Size = new System.Drawing.Size(911, 326);
            this.xtraTabControl1.TabIndex = 11;
            this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtraTabPage1,
            this.xtraTabPage2});
            // 
            // xtraTabPage1
            // 
            this.xtraTabPage1.Appearance.HeaderActive.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.xtraTabPage1.Appearance.HeaderActive.Options.UseFont = true;
            this.xtraTabPage1.Controls.Add(this.richEditControlCS);
            this.xtraTabPage1.Name = "xtraTabPage1";
            this.xtraTabPage1.Size = new System.Drawing.Size(909, 301);
            this.xtraTabPage1.Text = "C#";
            // 
            // richEditControlCS
            // 
            this.richEditControlCS.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Draft;
            this.richEditControlCS.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControlCS.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel;
            this.richEditControlCS.Location = new System.Drawing.Point(0, 0);
            this.richEditControlCS.Name = "richEditControlCS";
            this.richEditControlCS.Options.Annotations.ShowAllAuthors = false;
            this.richEditControlCS.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden;
            this.richEditControlCS.Size = new System.Drawing.Size(909, 301);
            this.richEditControlCS.TabIndex = 14;
            // 
            // xtraTabPage2
            // 
            this.xtraTabPage2.Appearance.HeaderActive.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.xtraTabPage2.Appearance.HeaderActive.Options.UseFont = true;
            this.xtraTabPage2.Controls.Add(this.richEditControlVB);
            this.xtraTabPage2.Name = "xtraTabPage2";
            this.xtraTabPage2.Size = new System.Drawing.Size(938, 236);
            this.xtraTabPage2.Text = "VB";
            // 
            // richEditControlVB
            // 
            this.richEditControlVB.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Draft;
            this.richEditControlVB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControlVB.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel;
            this.richEditControlVB.Location = new System.Drawing.Point(0, 0);
            this.richEditControlVB.Name = "richEditControlVB";
            this.richEditControlVB.Options.Annotations.ShowAllAuthors = false;
            this.richEditControlVB.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden;
            this.richEditControlVB.Size = new System.Drawing.Size(938, 236);
            this.richEditControlVB.TabIndex = 15;
            // 
            // spreadsheet
            // 
            this.spreadsheet.Location = new System.Drawing.Point(12, 388);
            this.spreadsheet.Name = "spreadsheet";
            this.spreadsheet.Options.Culture = new System.Globalization.CultureInfo("ru-RU");
            this.spreadsheet.Options.Export.Csv.Culture = new System.Globalization.CultureInfo("");
            this.spreadsheet.Options.Export.Txt.Culture = new System.Globalization.CultureInfo("");
            this.spreadsheet.Options.Import.Csv.AutoDetectDelimiter = false;
            this.spreadsheet.Options.Import.Csv.Culture = new System.Globalization.CultureInfo("");
            this.spreadsheet.Options.Import.Csv.Delimiter = ',';
            this.spreadsheet.Options.Import.Txt.AutoDetectDelimiter = false;
            this.spreadsheet.Options.Import.Txt.Culture = new System.Globalization.CultureInfo("");
            this.spreadsheet.Options.Import.Txt.Delimiter = ',';
            this.spreadsheet.Options.View.Charts.Antialiasing = DevExpress.XtraSpreadsheet.DocumentCapability.Enabled;
            this.spreadsheet.Size = new System.Drawing.Size(911, 255);
            this.spreadsheet.TabIndex = 5;
            // 
            // treeList1
            // 
            this.treeList1.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline);
            this.treeList1.Appearance.FocusedCell.Options.UseFont = true;
            this.treeList1.Location = new System.Drawing.Point(937, 12);
            this.treeList1.Name = "treeList1";
            this.treeList1.Size = new System.Drawing.Size(263, 631);
            this.treeList1.TabIndex = 11;
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.treeList1);
            this.layoutControl1.Controls.Add(this.spreadsheet);
            this.layoutControl1.Controls.Add(this.xtraTabControl1);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(269, 0, 650, 400);
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1212, 655);
            this.layoutControl1.TabIndex = 15;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.codeExampleNameLbl,
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.splitterItem1,
            this.splitterItem2});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1212, 655);
            this.Root.TextVisible = false;
            // 
            // codeExampleNameLbl
            // 
            this.codeExampleNameLbl.AllowHotTrack = false;
            this.codeExampleNameLbl.AppearanceItemCaption.Font = new System.Drawing.Font("Arial", 20.25F);
            this.codeExampleNameLbl.AppearanceItemCaption.Options.UseFont = true;
            this.codeExampleNameLbl.Location = new System.Drawing.Point(0, 0);
            this.codeExampleNameLbl.Name = "codeExampleNameLbl";
            this.codeExampleNameLbl.Size = new System.Drawing.Size(915, 36);
            this.codeExampleNameLbl.TextSize = new System.Drawing.Size(335, 32);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.xtraTabControl1;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 36);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(915, 330);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.spreadsheet;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 376);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(915, 259);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.treeList1;
            this.layoutControlItem3.Location = new System.Drawing.Point(925, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(267, 635);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // splitterItem1
            // 
            this.splitterItem1.AllowHotTrack = true;
            this.splitterItem1.Location = new System.Drawing.Point(915, 0);
            this.splitterItem1.Name = "splitterItem1";
            this.splitterItem1.Size = new System.Drawing.Size(10, 635);
            // 
            // splitterItem2
            // 
            this.splitterItem2.AllowHotTrack = true;
            this.splitterItem2.Location = new System.Drawing.Point(0, 366);
            this.splitterItem2.Name = "splitterItem2";
            this.splitterItem2.Size = new System.Drawing.Size(915, 10);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1212, 655);
            this.Controls.Add(this.layoutControl1);
            this.Name = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
            this.xtraTabControl1.ResumeLayout(false);
            this.xtraTabPage1.ResumeLayout(false);
            this.xtraTabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.codeExampleNameLbl)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitterItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitterItem2)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion

        void xtraTabControl1_SelectedPageChanged(object sender, TabPageChangedEventArgs e)
        {
            ExampleLanguage value = (ExampleLanguage)(xtraTabControl1.SelectedTabPageIndex);
            if (codeEditor != null)
                this.codeEditor.CurrentExampleLanguage = value;
        }
        void ChartAPIModule_Disposed(object sender, EventArgs e)
        {
            evaluator.Dispose();
        }
        void DisableTabs(int examplesCSCount, int examplesVBCount)
        {
            if (examplesCSCount == 0)
                xtraTabControl1.TabPages[(int)ExampleLanguage.Csharp].PageEnabled = false;
            if (examplesVBCount == 0)
                xtraTabControl1.TabPages[(int)ExampleLanguage.VB].PageEnabled = false;
        }
    }
}
