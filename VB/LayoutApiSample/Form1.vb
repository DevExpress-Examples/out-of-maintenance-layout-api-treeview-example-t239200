Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Layout
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace LayoutApiSample
	Partial Public Class Form1
		Inherits DevExpress.XtraBars.Ribbon.RibbonForm

		Private layoutTreeDictionary As Dictionary(Of TreeNode, LayoutElement) = New Dictionary(Of TreeNode,LayoutElement)()
		Private rebuild As Boolean = True

		Public Sub New()
			InitializeComponent()

			treeView1.ShowNodeToolTips = True

			AddHandler richEditControl1.MouseMove, AddressOf richEditControl1_MouseMove

			richEditControl1.LoadDocument("FloatingObjects.docx")
			richEditControl1.Document.Sections(0).LineNumbering.CountBy = 1
			AddHandler richEditControl1.DocumentLayout.DocumentFormatted, AddressOf DocumentLayout_DocumentFormatted
			AddHandler treeView1.AfterSelect, AddressOf treeView1_AfterSelect
		End Sub
		#Region "#mousemove"
		Private Sub richEditControl1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
			Dim pos As PageLayoutPosition = richEditControl1.ActiveView.GetDocumentLayoutPosition(e.Location)
			If pos IsNot Nothing Then
				Me.barStaticItem1.Caption = System.String.Format("Mouse is over page {0}, position {1}", pos.PageIndex, pos.Position)
			Else
				Me.barStaticItem1.Caption = ""
			End If
		End Sub
		#End Region ' #mousemove

		Private Sub treeView1_AfterSelect(ByVal sender As Object, ByVal e As TreeViewEventArgs)
			If Not (e.Node.IsSelected) OrElse e.Node.Tag Is Nothing Then
				Return
			End If
			' Rebuild the tree if required.
			rebuild = True

			richEditControl1.Document.ChangeActiveDocument(richEditControl1.Document)
			Dim element As LayoutElement = layoutTreeDictionary(e.Node)
			Dim rangedElement As RangedLayoutElement = TryCast(element, RangedLayoutElement)

			' Select a range or scroll to a document position, according to the type of the layout element.
			If DirectCast(e.Node.Tag, ContentDisplayAction) = ContentDisplayAction.ScrollTo Then
				Dim page As LayoutPage = element.GetParentByType(Of LayoutPage)()
				Dim nearestPageArea As LayoutPageArea = page.PageAreas(0)

				If (element.Type = LayoutType.Header) OrElse (element.GetParentByType(Of LayoutHeader)() IsNot Nothing) Then
					ScrollToPosition(nearestPageArea.Range.Start)
				End If
				If (element.Type = LayoutType.Footer) OrElse (element.GetParentByType(Of LayoutFooter)() IsNot Nothing) Then
					ScrollToPosition(nearestPageArea.Range.Start + nearestPageArea.Range.Length)
				End If
				Dim layoutFloatingObject As LayoutFloatingObject = TryCast(element, LayoutFloatingObject)
				If layoutFloatingObject IsNot Nothing Then
'INSTANT VB NOTE: The variable anchor was renamed since Visual Basic does not handle local variables named the same as class members well:
					Dim anchor_Renamed As FloatingObjectAnchorBox = layoutFloatingObject.AnchorBox
					ScrollToPosition(anchor_Renamed.Range.Start)
					richEditControl1.Document.Selection = richEditControl1.Document.CreateRange(anchor_Renamed.Range.Start, anchor_Renamed.Range.Length)
				End If
				Dim layoutComment As LayoutComment = TryCast(element, LayoutComment)
				If layoutComment IsNot Nothing Then
					Dim comment As Comment = layoutComment.GetDocumentComment()
					ScrollToPosition(comment.Range.Start.ToInt())
				End If

				Dim textBox As LayoutTextBox = element.GetParentByType(Of LayoutTextBox)()
				If textBox IsNot Nothing Then
					' Do not rebuild the tree.
					rebuild = False
					ScrollToPosition(textBox.AnchorBox.Range.Start)
					richEditControl1.Document.ChangeActiveDocument(textBox.Document)
					richEditControl1.Document.Selection = textBox.Document.CreateRange(rangedElement.Range.Start, rangedElement.Range.Length)
				End If

			Else
				If rangedElement Is Nothing Then
					Return
				End If
				ScrollToPosition(rangedElement.Range.Start)
				richEditControl1.Document.Selection = richEditControl1.Document.CreateRange(rangedElement.Range.Start, rangedElement.Range.Length)
			End If
		End Sub

		Private Sub ScrollToPosition(ByVal position As Integer)
			richEditControl1.Document.CaretPosition = richEditControl1.Document.CreatePosition(position)
			richEditControl1.ScrollToCaret(0.5F)
		End Sub

		#Region "#DocumentFormatted"
		Private Sub DocumentLayout_DocumentFormatted(ByVal sender As Object, ByVal e As System.EventArgs)
			' Do not rebuild the tree if the textbox content has been selected.
			If Not rebuild Then
				Return
			End If

			richEditControl1.BeginInvoke(New Action(Sub()
					' Create a new instance of a custom visitor.
					' Traverse the document layout tree.
					' Create a tree node / layout element dictionary for later use.
				treeView1.Nodes.Clear()
				Dim pageCount As Integer = richEditControl1.DocumentLayout.GetFormattedPageCount()
				For i As Integer = 0 To pageCount - 1
					Dim collector As New TreeViewCollector(treeView1, richEditControl1)
					collector.Visit(richEditControl1.DocumentLayout.GetPage(i))
					collector.Dictionary.ToList().ForEach(Sub(x) layoutTreeDictionary.Add(x.Value, x.Key))
				Next i
			End Sub))
		End Sub
		#End Region ' #DocumentFormatted
	End Class

	#Region "#TreeViewCollector"
	Friend Class TreeViewCollector
		Inherits LayoutVisitor

'INSTANT VB NOTE: The variable view was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private ReadOnly view_Renamed As TreeView
'INSTANT VB NOTE: The variable dictionary was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private dictionary_Renamed As Dictionary(Of LayoutElement, TreeNode)
'INSTANT VB NOTE: The variable richEdit was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private richEdit_Renamed As RichEditControl

		Public Sub New(ByVal view As TreeView, ByVal richEdit As RichEditControl)
			Me.view_Renamed = view
			Me.dictionary_Renamed = New Dictionary(Of LayoutElement, TreeNode)()
			Me.richEdit_Renamed = richEdit
		End Sub

		Public ReadOnly Property View() As TreeView
			Get
				Return view_Renamed
			End Get
		End Property
		Public ReadOnly Property Dictionary() As Dictionary(Of LayoutElement, TreeNode)
			Get
				Return dictionary_Renamed
			End Get
		End Property
		Public ReadOnly Property RichEdit() As RichEditControl
			Get
				Return richEdit_Renamed
			End Get
		End Property

		Protected Overrides Sub VisitPage(ByVal page As LayoutPage)
			Dim item As New TreeNode()
			item.Text = String.Format("{0} #{1}", "Page", page.Index + 1)

			Dictionary.Add(page, item)
			MyBase.VisitPage(page)
			Me.View.Nodes.Add(item)
		End Sub
		Protected Overrides Sub VisitHeader(ByVal header As LayoutHeader)
			AddTreeNode(header, ContentDisplayAction.ScrollTo)
			MyBase.VisitHeader(header)
		End Sub
		Protected Overrides Sub VisitFooter(ByVal footer As LayoutFooter)
			AddTreeNode(footer, ContentDisplayAction.ScrollTo)
			MyBase.VisitFooter(footer)
		End Sub
		Protected Overrides Sub VisitPageArea(ByVal pageArea As LayoutPageArea)
			AddTreeNode(pageArea, ContentDisplayAction.Select)
			MyBase.VisitPageArea(pageArea)
		End Sub
		Protected Overrides Sub VisitBookmarkEndBox(ByVal bookmarkEndBox As BookmarkBox)
			AddTreeNode(bookmarkEndBox, ContentDisplayAction.Select)
			MyBase.VisitBookmarkEndBox(bookmarkEndBox)
		End Sub
		Protected Overrides Sub VisitBookmarkStartBox(ByVal bookmarkStartBox As BookmarkBox)
			AddTreeNode(bookmarkStartBox, ContentDisplayAction.Select)
			MyBase.VisitBookmarkStartBox(bookmarkStartBox)
		End Sub
		Protected Overrides Sub VisitColumn(ByVal column As LayoutColumn)
			AddTreeNode(column, ContentDisplayAction.Select)
			MyBase.VisitColumn(column)
		End Sub
		Protected Overrides Sub VisitColumnBreakBox(ByVal columnBreakBox As PlainTextBox)
			AddTreeNode(columnBreakBox, ContentDisplayAction.Select)
			MyBase.VisitColumnBreakBox(columnBreakBox)
		End Sub
		Protected Overrides Sub VisitComment(ByVal comment As LayoutComment)
			AddTreeNode(comment, ContentDisplayAction.ScrollTo)
			MyBase.VisitComment(comment)
		End Sub
		Protected Overrides Sub VisitCommentEndBox(ByVal commentEndBox As CommentBox)
			AddTreeNode(commentEndBox, ContentDisplayAction.Select)
			MyBase.VisitCommentEndBox(commentEndBox)
		End Sub
		Protected Overrides Sub VisitCommentHighlightAreaBox(ByVal commentHighlightAreaBox As CommentHighlightAreaBox)
			AddTreeNode(commentHighlightAreaBox, ContentDisplayAction.Select)
			MyBase.VisitCommentHighlightAreaBox(commentHighlightAreaBox)
		End Sub
		Protected Overrides Sub VisitCommentStartBox(ByVal commentStartBox As CommentBox)
			AddTreeNode(commentStartBox, ContentDisplayAction.Select)
			MyBase.VisitCommentStartBox(commentStartBox)
		End Sub
		Protected Overrides Sub VisitFieldHighlightAreaBox(ByVal fieldHighlightAreaBox As FieldHighlightAreaBox)
			AddTreeNode(fieldHighlightAreaBox, ContentDisplayAction.Select)
			MyBase.VisitFieldHighlightAreaBox(fieldHighlightAreaBox)
		End Sub
		Protected Overrides Sub VisitFloatingObjectAnchorBox(ByVal floatingObjectAnchorBox As FloatingObjectAnchorBox)
			AddTreeNode(floatingObjectAnchorBox, ContentDisplayAction.Select)
			MyBase.VisitFloatingObjectAnchorBox(floatingObjectAnchorBox)
		End Sub
		Protected Overrides Sub VisitFloatingPicture(ByVal floatingPicture As LayoutFloatingPicture)
			AddTreeNode(floatingPicture, ContentDisplayAction.ScrollTo)
			MyBase.VisitFloatingPicture(floatingPicture)
		End Sub
		Protected Overrides Sub VisitHiddenTextUnderlineBox(ByVal hiddenTextUnderlineBox As HiddenTextUnderlineBox)
			AddTreeNode(hiddenTextUnderlineBox, ContentDisplayAction.Select)
			MyBase.VisitHiddenTextUnderlineBox(hiddenTextUnderlineBox)
		End Sub
		Protected Overrides Sub VisitHighlightAreaBox(ByVal highlightAreaBox As HighlightAreaBox)
			AddTreeNode(highlightAreaBox, ContentDisplayAction.Select)
			MyBase.VisitHighlightAreaBox(highlightAreaBox)
		End Sub
		Protected Overrides Sub VisitHyphenBox(ByVal hyphen As PlainTextBox)
			AddTreeNode(hyphen, ContentDisplayAction.Select)
			MyBase.VisitHyphenBox(hyphen)
		End Sub
		Protected Overrides Sub VisitInlinePictureBox(ByVal inlinePictureBox As InlinePictureBox)
			AddTreeNode(inlinePictureBox, ContentDisplayAction.Select)
			MyBase.VisitInlinePictureBox(inlinePictureBox)
		End Sub
		Protected Overrides Sub VisitLineBreakBox(ByVal lineBreakBox As PlainTextBox)
			AddTreeNode(lineBreakBox, ContentDisplayAction.Select)
			MyBase.VisitLineBreakBox(lineBreakBox)
		End Sub
		Protected Overrides Sub VisitLineNumberBox(ByVal lineNumberBox As LineNumberBox)
			AddTreeNode(lineNumberBox, ContentDisplayAction.Select)
			MyBase.VisitLineNumberBox(lineNumberBox)
		End Sub
		Protected Overrides Sub VisitNumberingListMarkBox(ByVal numberingListMarkBox As NumberingListMarkBox)
			AddTreeNode(numberingListMarkBox, ContentDisplayAction.Select)
			MyBase.VisitNumberingListMarkBox(numberingListMarkBox)
		End Sub
		Protected Overrides Sub VisitNumberingListWithSeparatorBox(ByVal numberingListWithSeparatorBox As NumberingListWithSeparatorBox)
			AddTreeNode(numberingListWithSeparatorBox, ContentDisplayAction.Select)
			MyBase.VisitNumberingListWithSeparatorBox(numberingListWithSeparatorBox)
		End Sub
		Protected Overrides Sub VisitPageBreakBox(ByVal pageBreakBox As PlainTextBox)
			AddTreeNode(pageBreakBox, ContentDisplayAction.Select)
			MyBase.VisitPageBreakBox(pageBreakBox)
		End Sub
		Protected Overrides Sub VisitPageNumberBox(ByVal pageNumberBox As PlainTextBox)
			AddTreeNode(pageNumberBox, ContentDisplayAction.Select)
			MyBase.VisitPageNumberBox(pageNumberBox)
		End Sub
		Protected Overrides Sub VisitParagraphMarkBox(ByVal paragraphMarkBox As PlainTextBox)
			AddTreeNode(paragraphMarkBox, ContentDisplayAction.Select)
			MyBase.VisitParagraphMarkBox(paragraphMarkBox)
		End Sub
		Protected Overrides Sub VisitPlainTextBox(ByVal plainTextBox As PlainTextBox)
			AddTreeNode(plainTextBox, ContentDisplayAction.Select)
			MyBase.VisitPlainTextBox(plainTextBox)
		End Sub
		Protected Overrides Sub VisitRangePermissionEndBox(ByVal rangePermissionEndBox As RangePermissionBox)
			AddTreeNode(rangePermissionEndBox, ContentDisplayAction.Select)
			MyBase.VisitRangePermissionEndBox(rangePermissionEndBox)
		End Sub
		Protected Overrides Sub VisitRangePermissionHighlightAreaBox(ByVal rangePermissionHighlightAreaBox As RangePermissionHighlightAreaBox)
			AddTreeNode(rangePermissionHighlightAreaBox, ContentDisplayAction.Select)
			MyBase.VisitRangePermissionHighlightAreaBox(rangePermissionHighlightAreaBox)
		End Sub
		Protected Overrides Sub VisitRangePermissionStartBox(ByVal rangePermissionStartBox As RangePermissionBox)
			AddTreeNode(rangePermissionStartBox, ContentDisplayAction.Select)
			MyBase.VisitRangePermissionStartBox(rangePermissionStartBox)
		End Sub
		Protected Overrides Sub VisitRow(ByVal row As LayoutRow)
			AddTreeNode(row, ContentDisplayAction.Select)
			MyBase.VisitRow(row)
		End Sub
		Protected Overrides Sub VisitSectionBreakBox(ByVal sectionBreakBox As PlainTextBox)
			AddTreeNode(sectionBreakBox, ContentDisplayAction.Select)
			MyBase.VisitSectionBreakBox(sectionBreakBox)
		End Sub
		Protected Overrides Sub VisitSpaceBox(ByVal spaceBox As PlainTextBox)
			AddTreeNode(spaceBox, ContentDisplayAction.Select)
			MyBase.VisitSpaceBox(spaceBox)
		End Sub
		Protected Overrides Sub VisitSpecialTextBox(ByVal specialTextBox As PlainTextBox)
			AddTreeNode(specialTextBox, ContentDisplayAction.Select)
			MyBase.VisitSpecialTextBox(specialTextBox)
		End Sub
		Protected Overrides Sub VisitStrikeoutBox(ByVal strikeoutBox As StrikeoutBox)
			AddTreeNode(strikeoutBox, ContentDisplayAction.Select)
			MyBase.VisitStrikeoutBox(strikeoutBox)
		End Sub
		Protected Overrides Sub VisitTable(ByVal table As LayoutTable)
			AddTreeNode(table, ContentDisplayAction.Select)
			MyBase.VisitTable(table)
		End Sub
		Protected Overrides Sub VisitTableCell(ByVal tableCell As LayoutTableCell)
			AddTreeNode(tableCell, ContentDisplayAction.Select)
			MyBase.VisitTableCell(tableCell)
		End Sub
		Protected Overrides Sub VisitTableRow(ByVal tableRow As LayoutTableRow)
			AddTreeNode(tableRow, ContentDisplayAction.Select)
			MyBase.VisitTableRow(tableRow)
		End Sub
		Protected Overrides Sub VisitTabSpaceBox(ByVal tabSpaceBox As PlainTextBox)
			AddTreeNode(tabSpaceBox, ContentDisplayAction.Select)
			MyBase.VisitTabSpaceBox(tabSpaceBox)
		End Sub
		Protected Overrides Sub VisitTextBox(ByVal textBox As LayoutTextBox)
			AddTreeNode(textBox, ContentDisplayAction.ScrollTo)
			MyBase.VisitTextBox(textBox)
		End Sub
		Protected Overrides Sub VisitUnderlineBox(ByVal underlineBox As UnderlineBox)
			AddTreeNode(underlineBox, ContentDisplayAction.Select)
			MyBase.VisitUnderlineBox(underlineBox)
		End Sub

		Private Sub AddTreeNode(ByVal element As LayoutElement, ByVal displayActionType As ContentDisplayAction)
			Dim item As New TreeNode()
			' Store the attribute that determines whether document range selection is allowed for this node.
			item.Tag = displayActionType
			Dim bounds As Rectangle = element.Bounds
			item.ToolTipText = String.Format("X = {0}" & ControlChars.Lf & "Y = {1}" & ControlChars.Lf & "Width = {2}" & ControlChars.Lf & "Height = {3}", bounds.X, bounds.Y, bounds.Width, bounds.Height)
			' Update the layout element / tree node dictionary.
			Dictionary.Add(element, item)

			Dim parentItem As TreeNode = Dictionary(element.Parent)
			If parentItem IsNot Nothing Then
				' Add a new node to the tree.
				Dim index As Integer = parentItem.Nodes.Add(item)
				' Specify the node caption.
				Dim typeName As String = element.Type.ToString()
				Select Case element.Type
					Case LayoutType.Column
						item.Text = String.Format("{0} #{1}", typeName, index)
					Case LayoutType.Row
						item.Text = String.Format("{0} #{1}", typeName, index + 1)
					Case LayoutType.TableRow
						item.Text = String.Format("{0} #{1}", typeName, index)
					Case Else
						item.Text = typeName
				End Select

				If parentItem.Tag IsNot Nothing Then
					' As for the node that does not allow range selection in the document, all its child nodes should have the same attribute.
					Dim parentDisplayActionType As ContentDisplayAction = DirectCast(parentItem.Tag, ContentDisplayAction)
					Dim actionType As ContentDisplayAction = If(parentDisplayActionType = ContentDisplayAction.ScrollTo, parentDisplayActionType, displayActionType)
					item.Tag = actionType
				End If
			End If
		End Sub
	End Class

	Public Enum ContentDisplayAction
		[Select]
		ScrollTo
	End Enum
	#End Region ' #TreeViewCollector
End Namespace
