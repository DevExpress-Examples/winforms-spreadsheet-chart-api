Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraTab
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Columns
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Namespace SpreadsheetChartAPISamples

    Public Class Form1
        Inherits Form

        Private defaultCulture As CultureInfo = New CultureInfo("en-US")

        'PrintableComponentLinkBase link;
        Private horizontalSplitContainerControl1 As SplitContainerControl

        Private verticalSplitContainerControl1 As SplitContainerControl

        'IContainer components = null;
'#Region "Controls"
        Private spreadsheet As SpreadsheetControl

        Private treeList1 As TreeList

        Private xtraTabControl1 As XtraTabControl

        Private xtraTabPage1 As XtraTabPage

        Private richEditControlCS As RichEditControl

        Private xtraTabPage2 As XtraTabPage

        Private richEditControlVB As RichEditControl

'#End Region
        Private codeExampleNameLbl As LabelControl

        Private codeEditor As ExampleCodeEditor

        Private evaluator As ExampleEvaluatorByTimer

        Private examples As List(Of CodeExampleGroup)

        Private treeListRootNodeLoading As Boolean = True

        Public Sub New()
            InitializeComponent()
            Dim examplePath As String = GetExamplePath("CodeExamples")
            'string examplePath = "D:\\VB\\CS\\SpreadsheetMainDemo\\CodeExamples";
            Dim examplesCS As Dictionary(Of String, FileInfo) = GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp)
            Dim examplesVB As Dictionary(Of String, FileInfo) = GatherExamplesFromProject(examplePath, ExampleLanguage.VB)
            DisableTabs(examplesCS.Count, examplesVB.Count)
            examples = FindExamples(examplePath, examplesCS, examplesVB)
            RearrangeExamples()
            MergeGroups()
            ShowExamplesInTreeList(treeList1, examples)
            codeEditor = New ExampleCodeEditor(richEditControlCS, richEditControlVB)
            CurrentExampleLanguage = DetectExampleLanguage("SpreadsheetChartAPISamples")
            evaluator = New SpreadsheetExampleEvaluatorByTimer() 'this.components
            AddHandler evaluator.QueryEvaluate, AddressOf OnExampleEvaluatorQueryEvaluate
            AddHandler evaluator.OnBeforeCompile, AddressOf evaluator_OnBeforeCompile
            AddHandler evaluator.OnAfterCompile, AddressOf evaluator_OnAfterCompile
            ShowFirstExample()
            AddHandler xtraTabControl1.SelectedPageChanged, New TabPageChangedEventHandler(AddressOf xtraTabControl1_SelectedPageChanged)
        End Sub

        Private Sub MergeGroups()
            Dim uniqueNameGroup = New Dictionary(Of String, CodeExampleGroup)()
            For Each n As CodeExampleGroup In examples
                If uniqueNameGroup.ContainsKey(n.Name) Then
                    uniqueNameGroup(n.Name).Merge(n)
                Else
                    uniqueNameGroup(n.Name) = n
                End If
            Next

            examples.Clear()
            For Each value In uniqueNameGroup.Values
                examples.Add(value)
            Next
        End Sub

        Private Sub RearrangeExamples()
            For i As Integer = 0 To examples.Count - 1
                Dim group As CodeExampleGroup = examples(i)
                If Equals(group.Name, "Charts") Then
                    examples.RemoveAt(i)
                    examples.Insert(0, group)
                    Exit For
                End If
            Next

            For i As Integer = 0 To examples.Count - 1
                Dim group As CodeExampleGroup = examples(i)
                If group.Name.StartsWith("Creation") Then
                    examples.RemoveAt(i)
                    examples.Insert(1, group)
                    Exit For
                End If
            Next
        End Sub

        Private Sub evaluator_OnAfterCompile(ByVal sender As Object, ByVal args As OnAfterCompileEventArgs)
            Dim workbook As IWorkbook = spreadsheet.Document
            For Each sheet As Worksheet In workbook.Worksheets
                sheet.PrintOptions.PrintGridlines = True
            Next

            Dim active As Worksheet = workbook.Worksheets.ActiveWorksheet
            Dim usedRange As Range = active.GetUsedRange()
            active.SelectedCell = usedRange(usedRange.RowCount * usedRange.ColumnCount - 1).Offset(1, 1)
            codeEditor.AfterCompile(args.Result)
            spreadsheet.EndUpdate()
        End Sub

        Private Sub evaluator_OnBeforeCompile(ByVal sender As Object, ByVal e As EventArgs)
            spreadsheet.BeginUpdate()
            codeEditor.BeforeCompile()
            Dim workbook As IWorkbook = spreadsheet.Document
            workbook.Options.Culture = defaultCulture
            Dim loaded As Boolean = workbook.LoadDocument("Document.xlsx")
            Call Debug.Assert(loaded)
        End Sub

        Private Property CurrentExampleLanguage As ExampleLanguage
            Get
                Return CType(xtraTabControl1.SelectedTabPageIndex, ExampleLanguage)
            End Get

            Set(ByVal value As ExampleLanguage)
                codeEditor.CurrentExampleLanguage = value
                xtraTabControl1.SelectedTabPageIndex = If(value = ExampleLanguage.Csharp, 0, 1)
            End Set
        End Property

        Private Sub ShowExamplesInTreeList(ByVal treeList As TreeList, ByVal examples As List(Of CodeExampleGroup))
'#Region "InitializeTreeList"
            treeList.OptionsPrint.UsePrintStyles = True
            AddHandler treeList.FocusedNodeChanged, New FocusedNodeChangedEventHandler(AddressOf OnNewExampleSelected)
            treeList.OptionsView.ShowColumns = False
            treeList.OptionsView.ShowIndicator = False
            AddHandler treeList.VirtualTreeGetChildNodes, AddressOf treeList_VirtualTreeGetChildNodes
            AddHandler treeList.VirtualTreeGetCellValue, AddressOf treeList_VirtualTreeGetCellValue
'#End Region
            Dim col1 As TreeListColumn = New TreeListColumn()
            col1.VisibleIndex = 0
            col1.OptionsColumn.AllowEdit = False
            col1.OptionsColumn.AllowMove = False
            col1.OptionsColumn.ReadOnly = True
            treeList.Columns.AddRange(New TreeListColumn() {col1})
            treeList.DataSource = New [Object]()
            treeList.ExpandAll()
        End Sub

        Private Sub treeList_VirtualTreeGetCellValue(ByVal sender As Object, ByVal args As VirtualTreeGetCellValueInfo)
            Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
            If group IsNot Nothing Then args.CellData = group.Name
            Dim example As CodeExample = TryCast(args.Node, CodeExample)
            If example IsNot Nothing Then args.CellData = example.RegionName
        End Sub

        Private Sub treeList_VirtualTreeGetChildNodes(ByVal sender As Object, ByVal args As VirtualTreeGetChildNodesInfo)
            If treeListRootNodeLoading Then
                args.Children = examples
                treeListRootNodeLoading = False
            Else
                If args.Node Is Nothing Then Return
                Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
                If group IsNot Nothing Then args.Children = group.Examples
            End If
        End Sub

        Private Sub ShowFirstExample()
            treeList1.ExpandAll()
            If treeList1.Nodes.Count > 0 Then treeList1.FocusedNode = treeList1.MoveFirst().FirstNode
        End Sub

        Private Sub OnNewExampleSelected(ByVal sender As Object, ByVal e As FocusedNodeChangedEventArgs)
            Dim newExample As CodeExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.Node), CodeExample)
            Dim oldExample As CodeExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.OldNode), CodeExample)
            If newExample Is Nothing Then Return
            Dim exampleCode As String = codeEditor.ShowExample(oldExample, newExample)
            codeExampleNameLbl.Text = ConvertStringToMoreHumanReadableForm(newExample.RegionName) & " example"
            Dim args As CodeEvaluationEventArgs = New CodeEvaluationEventArgs()
            InitializeCodeEvaluationEventArgs(args, newExample.RegionName)
            evaluator.ForceCompile(args)
        End Sub

        Private Sub InitializeCodeEvaluationEventArgs(ByVal e As CodeEvaluationEventArgs, ByVal regionName As String)
            e.Result = True
            e.Code = codeEditor.CurrentCodeEditor.Text
            e.Language = CurrentExampleLanguage
            e.EvaluationParameter = spreadsheet.Document
            e.RegionName = regionName
        End Sub

        Private Sub OnExampleEvaluatorQueryEvaluate(ByVal sender As Object, ByVal e As CodeEvaluationEventArgs)
            e.Result = False
            If codeEditor.RichEditTextChanged Then ' && compileComplete) {
                Dim span As TimeSpan = Date.Now - codeEditor.LastExampleCodeModifiedTime
                If span < TimeSpan.FromMilliseconds(1000) Then 'CompileTimeIntervalInMilliseconds  1900
                    codeEditor.ResetLastExampleModifiedTime()
                    Return
                End If

                'e.Result = true;
                InitializeCodeEvaluationEventArgs(e, e.RegionName)
            End If
        End Sub

'#Region "InitializeComponent"
        Private Sub InitializeComponent()
            horizontalSplitContainerControl1 = New SplitContainerControl()
            xtraTabControl1 = New XtraTabControl()
            xtraTabPage1 = New XtraTabPage()
            richEditControlCS = New RichEditControl()
            xtraTabPage2 = New XtraTabPage()
            richEditControlVB = New RichEditControl()
            codeExampleNameLbl = New LabelControl()
            spreadsheet = New SpreadsheetControl()
            verticalSplitContainerControl1 = New SplitContainerControl()
            treeList1 = New TreeList()
            CType(horizontalSplitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            horizontalSplitContainerControl1.SuspendLayout()
            CType(xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            xtraTabControl1.SuspendLayout()
            xtraTabPage1.SuspendLayout()
            xtraTabPage2.SuspendLayout()
            CType(verticalSplitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            verticalSplitContainerControl1.SuspendLayout()
            CType(treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            ' 
            ' horizontalSplitContainerControl1
            ' 
            horizontalSplitContainerControl1.Dock = DockStyle.Fill
            horizontalSplitContainerControl1.FixedPanel = SplitFixedPanel.Panel2
            horizontalSplitContainerControl1.Horizontal = False
            horizontalSplitContainerControl1.Location = New System.Drawing.Point(0, 0)
            horizontalSplitContainerControl1.Name = "horizontalSplitContainerControl1"
            horizontalSplitContainerControl1.Panel1.Controls.Add(xtraTabControl1)
            horizontalSplitContainerControl1.Panel1.Controls.Add(codeExampleNameLbl)
            horizontalSplitContainerControl1.Panel1.Text = "Panel1"
            horizontalSplitContainerControl1.Panel2.Controls.Add(spreadsheet)
            horizontalSplitContainerControl1.Panel2.Text = "Panel2"
            horizontalSplitContainerControl1.Size = New System.Drawing.Size(945, 655)
            horizontalSplitContainerControl1.SplitterPosition = 340
            horizontalSplitContainerControl1.TabIndex = 2
            horizontalSplitContainerControl1.Text = "splitContainerControl1"
            ' 
            ' xtraTabControl1
            ' 
            xtraTabControl1.AppearancePage.PageClient.BackColor = System.Drawing.Color.Transparent
            xtraTabControl1.AppearancePage.PageClient.BackColor2 = System.Drawing.Color.Transparent
            xtraTabControl1.AppearancePage.PageClient.BorderColor = System.Drawing.Color.Transparent
            xtraTabControl1.AppearancePage.PageClient.Options.UseBackColor = True
            xtraTabControl1.AppearancePage.PageClient.Options.UseBorderColor = True
            xtraTabControl1.Dock = DockStyle.Fill
            xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.True
            xtraTabControl1.Location = New System.Drawing.Point(0, 44)
            xtraTabControl1.Name = "xtraTabControl1"
            xtraTabControl1.SelectedTabPage = xtraTabPage1
            xtraTabControl1.Size = New System.Drawing.Size(945, 266)
            xtraTabControl1.TabIndex = 11
            xtraTabControl1.TabPages.AddRange(New XtraTabPage() {xtraTabPage1, xtraTabPage2})
            ' 
            ' xtraTabPage1
            ' 
            xtraTabPage1.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            xtraTabPage1.Appearance.HeaderActive.Options.UseFont = True
            xtraTabPage1.Controls.Add(richEditControlCS)
            xtraTabPage1.Name = "xtraTabPage1"
            xtraTabPage1.Size = New System.Drawing.Size(939, 238)
            xtraTabPage1.Text = "C#"
            ' 
            ' richEditControlCS
            ' 
            richEditControlCS.ActiveViewType = RichEditViewType.Draft
            richEditControlCS.Dock = DockStyle.Fill
            richEditControlCS.EnableToolTips = True
            richEditControlCS.Location = New System.Drawing.Point(0, 0)
            richEditControlCS.Name = "richEditControlCS"
            richEditControlCS.Options.Comments.ShowAllAuthors = False
            richEditControlCS.Options.CopyPaste.MaintainDocumentSectionSettings = False
            richEditControlCS.Options.Fields.UseCurrentCultureDateTimeFormat = False
            richEditControlCS.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden
            richEditControlCS.Options.MailMerge.KeepLastParagraph = False
            richEditControlCS.Size = New System.Drawing.Size(939, 238)
            richEditControlCS.TabIndex = 14
            ' 
            ' xtraTabPage2
            ' 
            xtraTabPage2.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            xtraTabPage2.Appearance.HeaderActive.Options.UseFont = True
            xtraTabPage2.Controls.Add(richEditControlVB)
            xtraTabPage2.Name = "xtraTabPage2"
            xtraTabPage2.Size = New System.Drawing.Size(939, 238)
            xtraTabPage2.Text = "VB"
            ' 
            ' richEditControlVB
            ' 
            richEditControlVB.ActiveViewType = RichEditViewType.Draft
            richEditControlVB.Dock = DockStyle.Fill
            richEditControlVB.EnableToolTips = True
            richEditControlVB.Location = New System.Drawing.Point(0, 0)
            richEditControlVB.Name = "richEditControlVB"
            richEditControlVB.Options.Comments.ShowAllAuthors = False
            richEditControlVB.Options.CopyPaste.MaintainDocumentSectionSettings = False
            richEditControlVB.Options.Fields.UseCurrentCultureDateTimeFormat = False
            richEditControlVB.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden
            richEditControlVB.Options.MailMerge.KeepLastParagraph = False
            richEditControlVB.Size = New System.Drawing.Size(939, 238)
            richEditControlVB.TabIndex = 15
            ' 
            ' codeExampleNameLbl
            ' 
            codeExampleNameLbl.Appearance.Font = New System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(204)))
            codeExampleNameLbl.Dock = DockStyle.Top
            codeExampleNameLbl.Location = New System.Drawing.Point(0, 0)
            codeExampleNameLbl.Margin = New Padding(3, 5, 3, 5)
            codeExampleNameLbl.Name = "codeExampleNameLbl"
            codeExampleNameLbl.Padding = New Padding(0, 0, 0, 12)
            codeExampleNameLbl.Size = New System.Drawing.Size(72, 44)
            codeExampleNameLbl.TabIndex = 10
            codeExampleNameLbl.Text = "label1"
            ' 
            ' spreadsheet
            ' 
            spreadsheet.Dock = DockStyle.Fill
            spreadsheet.Location = New System.Drawing.Point(0, 0)
            spreadsheet.Name = "spreadsheet"
            spreadsheet.Options.Culture = New CultureInfo("ru-RU")
            spreadsheet.Options.Export.Csv.Culture = New CultureInfo("")
            spreadsheet.Options.Export.Txt.Culture = New CultureInfo("")
            spreadsheet.Options.Import.Csv.Culture = New CultureInfo("")
            spreadsheet.Options.Import.Csv.Delimiter = ","c
            spreadsheet.Options.Import.Txt.Culture = New CultureInfo("")
            spreadsheet.Options.Import.Txt.Delimiter = ","c
            spreadsheet.Options.View.Charts.Antialiasing = DevExpress.XtraSpreadsheet.DocumentCapability.Enabled
            spreadsheet.Size = New System.Drawing.Size(945, 340)
            spreadsheet.TabIndex = 5
            ' 
            ' verticalSplitContainerControl1
            ' 
            verticalSplitContainerControl1.Dock = DockStyle.Fill
            verticalSplitContainerControl1.FixedPanel = SplitFixedPanel.Panel2
            verticalSplitContainerControl1.Location = New System.Drawing.Point(0, 0)
            verticalSplitContainerControl1.Name = "verticalSplitContainerControl1"
            verticalSplitContainerControl1.Panel1.Controls.Add(horizontalSplitContainerControl1)
            verticalSplitContainerControl1.Panel1.Text = "Panel1"
            verticalSplitContainerControl1.Panel2.Controls.Add(treeList1)
            verticalSplitContainerControl1.Panel2.Text = "Panel2"
            verticalSplitContainerControl1.Size = New System.Drawing.Size(1212, 655)
            verticalSplitContainerControl1.SplitterPosition = 262
            verticalSplitContainerControl1.TabIndex = 0
            verticalSplitContainerControl1.Text = "verticalSplitContainerControl1"
            ' 
            ' treeList1
            ' 
            treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline)
            treeList1.Appearance.FocusedCell.Options.UseFont = True
            treeList1.Dock = DockStyle.Fill
            treeList1.Location = New System.Drawing.Point(0, 0)
            treeList1.Name = "treeList1"
            treeList1.Size = New System.Drawing.Size(262, 655)
            treeList1.TabIndex = 11
            ' 
            ' Form1
            ' 
            AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            AutoScaleMode = AutoScaleMode.Font
            ClientSize = New System.Drawing.Size(1212, 655)
            Me.Controls.Add(verticalSplitContainerControl1)
            Name = "Form1"
            CType(horizontalSplitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            horizontalSplitContainerControl1.ResumeLayout(False)
            CType(xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            xtraTabControl1.ResumeLayout(False)
            xtraTabPage1.ResumeLayout(False)
            xtraTabPage2.ResumeLayout(False)
            CType(verticalSplitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            verticalSplitContainerControl1.ResumeLayout(False)
            CType(treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
        End Sub

'#End Region
        Private Sub xtraTabControl1_SelectedPageChanged(ByVal sender As Object, ByVal e As TabPageChangedEventArgs)
            Dim value As ExampleLanguage = CType(xtraTabControl1.SelectedTabPageIndex, ExampleLanguage)
            If codeEditor IsNot Nothing Then codeEditor.CurrentExampleLanguage = value
        End Sub

        Private Sub ChartAPIModule_Disposed(ByVal sender As Object, ByVal e As EventArgs)
            evaluator.Dispose()
        End Sub

        Private Sub DisableTabs(ByVal examplesCSCount As Integer, ByVal examplesVBCount As Integer)
            If examplesCSCount = 0 Then xtraTabControl1.TabPages(CInt(ExampleLanguage.Csharp)).PageEnabled = False
            If examplesVBCount = 0 Then xtraTabControl1.TabPages(CInt(ExampleLanguage.VB)).PageEnabled = False
        End Sub
    End Class
End Namespace
