Imports Microsoft.VisualBasic
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
        Inherits DevExpress.XtraEditors.XtraForm
        Private defaultCulture As New CultureInfo("en-US")
        'IContainer components = null;

#Region "Controls"
        Private spreadsheet As SpreadsheetControl
        Private treeList1 As TreeList
        Private xtraTabControl1 As XtraTabControl
        Private xtraTabPage1 As XtraTabPage
        Private richEditControlCS As RichEditControl
        Private xtraTabPage2 As XtraTabPage
        Private richEditControlVB As RichEditControl
        Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
        Friend WithEvents Root As DevExpress.XtraLayout.LayoutControlGroup
        Friend WithEvents codeExampleNameLbl As DevExpress.XtraLayout.SimpleLabelItem
        Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents SplitterItem1 As DevExpress.XtraLayout.SplitterItem
        Friend WithEvents SplitterItem2 As DevExpress.XtraLayout.SplitterItem
#End Region
        Private codeEditor As ExampleCodeEditor
        Private evaluator As ExampleEvaluatorByTimer
        Private examples As List(Of CodeExampleGroup)
        Private treeListRootNodeLoading As Boolean = True

        Public Sub New()
			InitializeComponent()
			Dim examplePath As String = CodeExampleDemoUtils.GetExamplePath("CodeExamples")

			'string examplePath = "D:\\VB\\CS\\SpreadsheetMainDemo\\CodeExamples";
			Dim examplesCS As Dictionary(Of String, FileInfo) = CodeExampleDemoUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp)
			Dim examplesVB As Dictionary(Of String, FileInfo) = CodeExampleDemoUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.VB)
			DisableTabs(examplesCS.Count, examplesVB.Count)
			Me.examples = CodeExampleDemoUtils.FindExamples(examplePath, examplesCS, examplesVB)
			RearrangeExamples()
			MergeGroups()
			ShowExamplesInTreeList(treeList1, examples)

			Me.codeEditor = New ExampleCodeEditor(richEditControlCS, richEditControlVB)
			CurrentExampleLanguage = CodeExampleDemoUtils.DetectExampleLanguage("SpreadsheetChartAPISamples")
			Me.evaluator = New SpreadsheetExampleEvaluatorByTimer() 'this.components

			AddHandler Me.evaluator.QueryEvaluate, AddressOf OnExampleEvaluatorQueryEvaluate
			AddHandler Me.evaluator.OnBeforeCompile, AddressOf evaluator_OnBeforeCompile
			AddHandler Me.evaluator.OnAfterCompile, AddressOf evaluator_OnAfterCompile

			ShowFirstExample()
			AddHandler xtraTabControl1.SelectedPageChanged, AddressOf xtraTabControl1_SelectedPageChanged
		End Sub

		Private Sub MergeGroups()
			Dim uniqueNameGroup = New Dictionary(Of String, CodeExampleGroup)()
			For Each n As CodeExampleGroup In examples
				If uniqueNameGroup.ContainsKey(n.Name) Then
					uniqueNameGroup(n.Name).Merge(n)
				Else
					uniqueNameGroup(n.Name) = n
				End If
			Next n

			examples.Clear()
			For Each value In uniqueNameGroup.Values
				examples.Add(value)
			Next value
		End Sub

		Private Sub RearrangeExamples()
			Dim i As Integer = 0
			Do While i < examples.Count
				Dim group As CodeExampleGroup = examples(i)
				If group.Name = "Charts" Then
					examples.RemoveAt(i)
					examples.Insert(0, group)
					Exit Do
				End If
				i += 1
			Loop
			i = 0
			Do While i < examples.Count
				Dim group As CodeExampleGroup = examples(i)
				If group.Name.StartsWith("Creation") Then
					examples.RemoveAt(i)
					examples.Insert(1, group)
					Exit Do
				End If
				i += 1
			Loop
		End Sub

		Private Sub evaluator_OnAfterCompile(ByVal sender As Object, ByVal args As OnAfterCompileEventArgs)
			Dim workbook As IWorkbook = spreadsheet.Document
			For Each sheet As Worksheet In workbook.Worksheets
				sheet.PrintOptions.PrintGridlines = True
			Next sheet

            Dim active As Worksheet = workbook.Worksheets.ActiveWorksheet
            If active IsNot Nothing Then
                Dim usedRange As CellRange = active.GetUsedRange()
                active.SelectedCell = usedRange(usedRange.RowCount * usedRange.ColumnCount - 1).Offset(1, 1)
            End If

            codeEditor.AfterCompile(args.Result)
			spreadsheet.EndUpdate()
		End Sub

		Private Sub evaluator_OnBeforeCompile(ByVal sender As Object, ByVal e As EventArgs)
			spreadsheet.BeginUpdate()
			codeEditor.BeforeCompile()

			Dim workbook As IWorkbook = spreadsheet.Document
			workbook.Options.Culture = defaultCulture
			Dim loaded As Boolean = workbook.LoadDocument("Document.xlsx")
			Debug.Assert(loaded)
		End Sub
		Private Property CurrentExampleLanguage() As ExampleLanguage
			Get
				Return CType(xtraTabControl1.SelectedTabPageIndex, ExampleLanguage)
			End Get
			Set(ByVal value As ExampleLanguage)
				Me.codeEditor.CurrentExampleLanguage = value
				xtraTabControl1.SelectedTabPageIndex = If((value = ExampleLanguage.Csharp), 0, 1)
			End Set
		End Property
		Private Sub ShowExamplesInTreeList(ByVal treeList As TreeList, ByVal examples As List(Of CodeExampleGroup))
'			#Region "InitializeTreeList"
			treeList.OptionsPrint.UsePrintStyles = True
			AddHandler treeList.FocusedNodeChanged, AddressOf OnNewExampleSelected
			treeList.OptionsView.ShowColumns = False
			treeList.OptionsView.ShowIndicator = False


			AddHandler treeList.VirtualTreeGetChildNodes, AddressOf treeList_VirtualTreeGetChildNodes
			AddHandler treeList.VirtualTreeGetCellValue, AddressOf treeList_VirtualTreeGetCellValue
'			#End Region
			Dim col1 As New TreeListColumn()
			col1.VisibleIndex = 0
			col1.OptionsColumn.AllowEdit = False
			col1.OptionsColumn.AllowMove = False
			col1.OptionsColumn.ReadOnly = True
			treeList.Columns.AddRange(New TreeListColumn() { col1 })

			treeList.DataSource = New Object()
			treeList.ExpandAll()
		End Sub

		Private Sub treeList_VirtualTreeGetCellValue(ByVal sender As Object, ByVal args As VirtualTreeGetCellValueInfo)
			Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
			If group IsNot Nothing Then
				args.CellData = group.Name
			End If

			Dim example As CodeExample = TryCast(args.Node, CodeExample)
			If example IsNot Nothing Then
				args.CellData = example.RegionName
			End If
		End Sub

		Private Sub treeList_VirtualTreeGetChildNodes(ByVal sender As Object, ByVal args As VirtualTreeGetChildNodesInfo)
			If treeListRootNodeLoading Then
				args.Children = examples
				treeListRootNodeLoading = False
			Else
				If args.Node Is Nothing Then
					Return
				End If
				Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
				If group IsNot Nothing Then
					args.Children = group.Examples
				End If
			End If
		End Sub
		Private Sub ShowFirstExample()
			treeList1.ExpandAll()
			If treeList1.Nodes.Count > 0 Then
				treeList1.FocusedNode = treeList1.MoveFirst().FirstNode
			End If
		End Sub
		Private Sub OnNewExampleSelected(ByVal sender As Object, ByVal e As FocusedNodeChangedEventArgs)
			Dim newExample As CodeExample = TryCast((TryCast(sender, TreeList)).GetDataRecordByNode(e.Node), CodeExample)
			Dim oldExample As CodeExample = TryCast((TryCast(sender, TreeList)).GetDataRecordByNode(e.OldNode), CodeExample)

			If newExample Is Nothing Then
				Return
			End If

			Dim exampleCode As String = codeEditor.ShowExample(oldExample, newExample)
			codeExampleNameLbl.Text = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(newExample.RegionName) & " example"
			Dim args As New CodeEvaluationEventArgs()
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
			If codeEditor.RichEditTextChanged Then
				Dim span As TimeSpan = DateTime.Now - codeEditor.LastExampleCodeModifiedTime

				If span < TimeSpan.FromMilliseconds(1000) Then
					codeEditor.ResetLastExampleModifiedTime()
					Return
				End If
				'e.Result = true;
				InitializeCodeEvaluationEventArgs(e, e.RegionName)
			End If
		End Sub
		#Region "InitializeComponent"
		Private Sub InitializeComponent()
            Me.xtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.xtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlCS = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlVB = New DevExpress.XtraRichEdit.RichEditControl()
            Me.spreadsheet = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
            Me.Root = New DevExpress.XtraLayout.LayoutControlGroup()
            Me.codeExampleNameLbl = New DevExpress.XtraLayout.SimpleLabelItem()
            Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.SplitterItem1 = New DevExpress.XtraLayout.SplitterItem()
            Me.SplitterItem2 = New DevExpress.XtraLayout.SplitterItem()
            CType(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabControl1.SuspendLayout()
            Me.xtraTabPage1.SuspendLayout()
            Me.xtraTabPage2.SuspendLayout()
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.LayoutControl1.SuspendLayout()
            CType(Me.Root, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.codeExampleNameLbl, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SplitterItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SplitterItem2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'xtraTabControl1
            '
            Me.xtraTabControl1.AppearancePage.PageClient.BackColor = System.Drawing.Color.Transparent
            Me.xtraTabControl1.AppearancePage.PageClient.BackColor2 = System.Drawing.Color.Transparent
            Me.xtraTabControl1.AppearancePage.PageClient.BorderColor = System.Drawing.Color.Transparent
            Me.xtraTabControl1.AppearancePage.PageClient.Options.UseBackColor = True
            Me.xtraTabControl1.AppearancePage.PageClient.Options.UseBorderColor = True
            Me.xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.[True]
            Me.xtraTabControl1.Location = New System.Drawing.Point(12, 48)
            Me.xtraTabControl1.Name = "xtraTabControl1"
            Me.xtraTabControl1.SelectedTabPage = Me.xtraTabPage1
            Me.xtraTabControl1.Size = New System.Drawing.Size(897, 357)
            Me.xtraTabControl1.TabIndex = 11
            Me.xtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.xtraTabPage1, Me.xtraTabPage2})
            '
            'xtraTabPage1
            '
            Me.xtraTabPage1.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
            Me.xtraTabPage1.Appearance.HeaderActive.Options.UseFont = True
            Me.xtraTabPage1.Controls.Add(Me.richEditControlCS)
            Me.xtraTabPage1.Name = "xtraTabPage1"
            Me.xtraTabPage1.Size = New System.Drawing.Size(895, 332)
            Me.xtraTabPage1.Text = "C#"
            '
            'richEditControlCS
            '
            Me.richEditControlCS.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Draft
            Me.richEditControlCS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlCS.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            Me.richEditControlCS.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlCS.Name = "richEditControlCS"
            Me.richEditControlCS.Options.Annotations.ShowAllAuthors = False
            Me.richEditControlCS.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlCS.Size = New System.Drawing.Size(895, 332)
            Me.richEditControlCS.TabIndex = 14
            '
            'xtraTabPage2
            '
            Me.xtraTabPage2.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
            Me.xtraTabPage2.Appearance.HeaderActive.Options.UseFont = True
            Me.xtraTabPage2.Controls.Add(Me.richEditControlVB)
            Me.xtraTabPage2.Name = "xtraTabPage2"
            Me.xtraTabPage2.Size = New System.Drawing.Size(938, 236)
            Me.xtraTabPage2.Text = "VB"
            '
            'richEditControlVB
            '
            Me.richEditControlVB.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Draft
            Me.richEditControlVB.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlVB.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            Me.richEditControlVB.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlVB.Name = "richEditControlVB"
            Me.richEditControlVB.Options.Annotations.ShowAllAuthors = False
            Me.richEditControlVB.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlVB.Size = New System.Drawing.Size(938, 236)
            Me.richEditControlVB.TabIndex = 15
            '
            'spreadsheet
            '
            Me.spreadsheet.Location = New System.Drawing.Point(12, 419)
            Me.spreadsheet.Name = "spreadsheet"
            Me.spreadsheet.Options.Culture = New System.Globalization.CultureInfo("ru-RU")
            Me.spreadsheet.Options.Export.Csv.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheet.Options.Export.Txt.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheet.Options.Import.Csv.AutoDetectDelimiter = False
            Me.spreadsheet.Options.Import.Csv.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheet.Options.Import.Csv.Delimiter = Global.Microsoft.VisualBasic.ChrW(44)
            Me.spreadsheet.Options.Import.Txt.AutoDetectDelimiter = False
            Me.spreadsheet.Options.Import.Txt.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheet.Options.Import.Txt.Delimiter = Global.Microsoft.VisualBasic.ChrW(44)
            Me.spreadsheet.Options.View.Charts.Antialiasing = DevExpress.XtraSpreadsheet.DocumentCapability.Enabled
            Me.spreadsheet.Size = New System.Drawing.Size(897, 224)
            Me.spreadsheet.TabIndex = 5
            '
            'treeList1
            '
            Me.treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Underline)
            Me.treeList1.Appearance.FocusedCell.Options.UseFont = True
            Me.treeList1.Location = New System.Drawing.Point(923, 12)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.Size = New System.Drawing.Size(277, 631)
            Me.treeList1.TabIndex = 11
            '
            'LayoutControl1
            '
            Me.LayoutControl1.Controls.Add(Me.treeList1)
            Me.LayoutControl1.Controls.Add(Me.spreadsheet)
            Me.LayoutControl1.Controls.Add(Me.xtraTabControl1)
            Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
            Me.LayoutControl1.Name = "LayoutControl1"
            Me.LayoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(488, 89, 650, 400)
            Me.LayoutControl1.Root = Me.Root
            Me.LayoutControl1.Size = New System.Drawing.Size(1212, 655)
            Me.LayoutControl1.TabIndex = 1
            Me.LayoutControl1.Text = "LayoutControl1"
            '
            'Root
            '
            Me.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
            Me.Root.GroupBordersVisible = False
            Me.Root.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.codeExampleNameLbl, Me.LayoutControlItem1, Me.LayoutControlItem2, Me.LayoutControlItem3, Me.SplitterItem1, Me.SplitterItem2})
            Me.Root.Name = "Root"
            Me.Root.Size = New System.Drawing.Size(1212, 655)
            Me.Root.TextVisible = False
            '
            'codeExampleNameLbl
            '
            Me.codeExampleNameLbl.AllowHotTrack = False
            Me.codeExampleNameLbl.AppearanceItemCaption.Font = New System.Drawing.Font("Arial", 20.25!)
            Me.codeExampleNameLbl.AppearanceItemCaption.Options.UseFont = True
            Me.codeExampleNameLbl.Location = New System.Drawing.Point(0, 0)
            Me.codeExampleNameLbl.Name = "codeExampleNameLbl"
            Me.codeExampleNameLbl.Size = New System.Drawing.Size(901, 36)
            Me.codeExampleNameLbl.TextSize = New System.Drawing.Size(335, 32)
            '
            'LayoutControlItem1
            '
            Me.LayoutControlItem1.Control = Me.xtraTabControl1
            Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 36)
            Me.LayoutControlItem1.Name = "LayoutControlItem1"
            Me.LayoutControlItem1.Size = New System.Drawing.Size(901, 361)
            Me.LayoutControlItem1.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem1.TextVisible = False
            '
            'LayoutControlItem2
            '
            Me.LayoutControlItem2.Control = Me.spreadsheet
            Me.LayoutControlItem2.Location = New System.Drawing.Point(0, 407)
            Me.LayoutControlItem2.Name = "LayoutControlItem2"
            Me.LayoutControlItem2.Size = New System.Drawing.Size(901, 228)
            Me.LayoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem2.TextVisible = False
            '
            'LayoutControlItem3
            '
            Me.LayoutControlItem3.Control = Me.treeList1
            Me.LayoutControlItem3.Location = New System.Drawing.Point(911, 0)
            Me.LayoutControlItem3.Name = "LayoutControlItem3"
            Me.LayoutControlItem3.Size = New System.Drawing.Size(281, 635)
            Me.LayoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem3.TextVisible = False
            '
            'SplitterItem1
            '
            Me.SplitterItem1.AllowHotTrack = True
            Me.SplitterItem1.Location = New System.Drawing.Point(901, 0)
            Me.SplitterItem1.Name = "SplitterItem1"
            Me.SplitterItem1.Size = New System.Drawing.Size(10, 635)
            '
            'SplitterItem2
            '
            Me.SplitterItem2.AllowHotTrack = True
            Me.SplitterItem2.Location = New System.Drawing.Point(0, 397)
            Me.SplitterItem2.Name = "SplitterItem2"
            Me.SplitterItem2.Size = New System.Drawing.Size(901, 10)
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1212, 655)
            Me.Controls.Add(Me.LayoutControl1)
            Me.Name = "Form1"
            CType(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabControl1.ResumeLayout(False)
            Me.xtraTabPage1.ResumeLayout(False)
            Me.xtraTabPage2.ResumeLayout(False)
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.LayoutControl1.ResumeLayout(False)
            CType(Me.Root, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.codeExampleNameLbl, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.SplitterItem1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.SplitterItem2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub
#End Region

        Private Sub xtraTabControl1_SelectedPageChanged(ByVal sender As Object, ByVal e As TabPageChangedEventArgs)
			Dim value As ExampleLanguage = CType(xtraTabControl1.SelectedTabPageIndex, ExampleLanguage)
			If codeEditor IsNot Nothing Then
				Me.codeEditor.CurrentExampleLanguage = value
			End If
		End Sub
		Private Sub ChartAPIModule_Disposed(ByVal sender As Object, ByVal e As EventArgs)
			evaluator.Dispose()
		End Sub
		Private Sub DisableTabs(ByVal examplesCSCount As Integer, ByVal examplesVBCount As Integer)
			If examplesCSCount = 0 Then
				xtraTabControl1.TabPages(CInt(Fix(ExampleLanguage.Csharp))).PageEnabled = False
			End If
			If examplesVBCount = 0 Then
				xtraTabControl1.TabPages(CInt(Fix(ExampleLanguage.VB))).PageEnabled = False
			End If
		End Sub
	End Class
End Namespace
