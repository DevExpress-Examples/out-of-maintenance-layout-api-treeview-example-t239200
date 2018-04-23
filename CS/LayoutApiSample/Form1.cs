using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Layout;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LayoutApiSample
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        Dictionary<TreeNode, LayoutElement> layoutTreeDictionary = new Dictionary<TreeNode,LayoutElement>();
        bool rebuild = true;

        public Form1()
        {
            InitializeComponent();
            richEditControl1.LoadDocument("FloatingObjects.docx");
            richEditControl1.Document.Sections[0].LineNumbering.CountBy = 1;
            richEditControl1.DocumentLayout.DocumentFormatted += DocumentLayout_DocumentFormatted;
            treeView1.AfterSelect += treeView1_AfterSelect;
        }

        void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (!(e.Node.IsSelected) || e.Node.Tag == null)
                return;
            // Rebuild the tree if required.
            rebuild = true;

            richEditControl1.Document.ChangeActiveDocument(richEditControl1.Document);
            LayoutElement element = layoutTreeDictionary[e.Node];
            RangedLayoutElement rangedElement = element as RangedLayoutElement;

            // Select a range or scroll to a document position, according to the type of the layout element.
            if ((ContentDisplayAction)e.Node.Tag == ContentDisplayAction.ScrollTo)
            {
                LayoutPage page = element.GetParentByType<LayoutPage>();
                LayoutPageArea nearestPageArea = page.PageAreas[0];

                if ((element.Type == LayoutType.Header) || (element.GetParentByType<LayoutHeader>() != null))
                    ScrollToPosition(nearestPageArea.Range.Start);
                if ((element.Type == LayoutType.Footer) || (element.GetParentByType<LayoutFooter>() != null))
                    ScrollToPosition(nearestPageArea.Range.End);
                LayoutFloatingObject layoutFloatingObject = element as LayoutFloatingObject;
                if (layoutFloatingObject != null)
                {
                    FloatingObjectAnchorBox anchor = layoutFloatingObject.AnchorBox;
                    ScrollToPosition(anchor.Range.Start);
                    richEditControl1.Document.Selection = richEditControl1.Document.CreateRange(anchor.Range.Start, anchor.Range.Length);
                }
                LayoutComment layoutComment = element as LayoutComment;
                if (layoutComment != null)
                {
                    Comment comment = layoutComment.GetDocumentComment();
                    ScrollToPosition(comment.Range.Start.ToInt());
                }
                
                LayoutTextBox textBox = element.GetParentByType<LayoutTextBox>();
                if (textBox != null)
                {
                    // Do not rebuild the tree.
                    rebuild = false;
                    ScrollToPosition(textBox.AnchorBox.Range.Start);
                    richEditControl1.Document.ChangeActiveDocument(textBox.Document);
                    richEditControl1.Document.Selection = textBox.Document.CreateRange(rangedElement.Range.Start, rangedElement.Range.Length);
                }

            }
            else
            {
                if (rangedElement == null)
                    return;
                ScrollToPosition(rangedElement.Range.Start);
                richEditControl1.Document.Selection = richEditControl1.Document.CreateRange(rangedElement.Range.Start, rangedElement.Range.Length);
            }
        }
        
        void ScrollToPosition(int position)
        {
            richEditControl1.Document.CaretPosition = richEditControl1.Document.CreatePosition(position);
            richEditControl1.ScrollToCaret(0.5f);
        }

        #region #DocumentFormatted
        void DocumentLayout_DocumentFormatted(object sender, System.EventArgs e)
        {
            // Do not rebuild the tree if the textbox content has been selected.
            if (!rebuild) return;

            richEditControl1.BeginInvoke(new Action(() =>
            {
                treeView1.Nodes.Clear();
                int pageCount = richEditControl1.DocumentLayout.GetFormattedPageCount();
                for (int i = 0; i < pageCount; i++)
                {
                    // Create a new instance of a custom visitor.
                    TreeViewCollector collector = new TreeViewCollector(treeView1, richEditControl1);
                    // Traverse the document layout tree.
                    collector.Visit(richEditControl1.DocumentLayout.GetPage(i));
                    // Create a tree node / layout element dictionary for later use.
                    collector.Dictionary.ToList().ForEach(x => layoutTreeDictionary.Add(x.Value, x.Key));
                }
            }));
        }
        #endregion #DocumentFormatted
    }

    #region #TreeViewCollector
    class TreeViewCollector : LayoutVisitor
    {
        readonly TreeView view;
        Dictionary<LayoutElement, TreeNode> dictionary;
        RichEditControl richEdit;

        public TreeViewCollector(TreeView view, RichEditControl richEdit)
        {
            this.view = view;
            this.dictionary = new Dictionary<LayoutElement, TreeNode>();
            this.richEdit = richEdit;
        }

        public TreeView View { get { return view; } }
        public Dictionary<LayoutElement, TreeNode> Dictionary { get { return dictionary; } }
        public RichEditControl RichEdit { get { return richEdit; } }

        protected override void VisitPage(LayoutPage page)
        {
            TreeNode item = new TreeNode();
            item.Text = String.Format("{0} #{1}", "Page", page.Index + 1); ;
            Dictionary.Add(page, item);
            base.VisitPage(page);
            View.Nodes.Add(item);
        }
        protected override void VisitHeader(LayoutHeader header)
        {
            AddTreeNode(header, ContentDisplayAction.ScrollTo);
            base.VisitHeader(header);
        }
        protected override void VisitFooter(LayoutFooter footer)
        {
            AddTreeNode(footer, ContentDisplayAction.ScrollTo);
            base.VisitFooter(footer);
        }
        protected override void VisitPageArea(LayoutPageArea pageArea)
        {
            AddTreeNode(pageArea, ContentDisplayAction.Select);
            base.VisitPageArea(pageArea);
        }
        protected override void VisitBookmarkEndBox(BookmarkBox bookmarkEndBox)
        {
            AddTreeNode(bookmarkEndBox, ContentDisplayAction.Select);
            base.VisitBookmarkEndBox(bookmarkEndBox);
        }
        protected override void VisitBookmarkStartBox(BookmarkBox bookmarkStartBox)
        {
            AddTreeNode(bookmarkStartBox, ContentDisplayAction.Select);
            base.VisitBookmarkStartBox(bookmarkStartBox);
        }
        protected override void VisitColumn(LayoutColumn column)
        {
            AddTreeNode(column, ContentDisplayAction.Select);
            base.VisitColumn(column);
        }
        protected override void VisitColumnBreakBox(PlainTextBox columnBreakBox)
        {
            AddTreeNode(columnBreakBox, ContentDisplayAction.Select);
            base.VisitColumnBreakBox(columnBreakBox);
        }
        protected override void VisitComment(LayoutComment comment)
        {
            AddTreeNode(comment, ContentDisplayAction.ScrollTo);
            base.VisitComment(comment);
        }
        protected override void VisitCommentEndBox(CommentBox commentEndBox)
        {
            AddTreeNode(commentEndBox, ContentDisplayAction.Select);
            base.VisitCommentEndBox(commentEndBox);
        }
        protected override void VisitCommentHighlightAreaBox(CommentHighlightAreaBox commentHighlightAreaBox)
        {
            AddTreeNode(commentHighlightAreaBox, ContentDisplayAction.Select);
            base.VisitCommentHighlightAreaBox(commentHighlightAreaBox);
        }
        protected override void VisitCommentStartBox(CommentBox commentStartBox)
        {
            AddTreeNode(commentStartBox, ContentDisplayAction.Select);
            base.VisitCommentStartBox(commentStartBox);
        }
        protected override void VisitFieldHighlightAreaBox(FieldHighlightAreaBox fieldHighlightAreaBox)
        {
            AddTreeNode(fieldHighlightAreaBox, ContentDisplayAction.Select);
            base.VisitFieldHighlightAreaBox(fieldHighlightAreaBox);
        }
        protected override void VisitFloatingObjectAnchorBox(FloatingObjectAnchorBox floatingObjectAnchorBox)
        {
            AddTreeNode(floatingObjectAnchorBox, ContentDisplayAction.Select);
            base.VisitFloatingObjectAnchorBox(floatingObjectAnchorBox);
        }
        protected override void VisitFloatingPicture(LayoutFloatingPicture floatingPicture)
        {
            AddTreeNode(floatingPicture, ContentDisplayAction.ScrollTo);
            base.VisitFloatingPicture(floatingPicture);
        }
        protected override void VisitHiddenTextUnderlineBox(HiddenTextUnderlineBox hiddenTextUnderlineBox)
        {
            AddTreeNode(hiddenTextUnderlineBox, ContentDisplayAction.Select);
            base.VisitHiddenTextUnderlineBox(hiddenTextUnderlineBox);
        }
        protected override void VisitHighlightAreaBox(HighlightAreaBox highlightAreaBox)
        {
            AddTreeNode(highlightAreaBox, ContentDisplayAction.Select);
            base.VisitHighlightAreaBox(highlightAreaBox);
        }
        protected override void VisitHyphenBox(PlainTextBox hyphen)
        {
            AddTreeNode(hyphen, ContentDisplayAction.Select);
            base.VisitHyphenBox(hyphen);
        }
        protected override void VisitInlinePictureBox(InlinePictureBox inlinePictureBox)
        {
            AddTreeNode(inlinePictureBox, ContentDisplayAction.Select);
            base.VisitInlinePictureBox(inlinePictureBox);
        }
        protected override void VisitLineBreakBox(PlainTextBox lineBreakBox)
        {
            AddTreeNode(lineBreakBox, ContentDisplayAction.Select);
            base.VisitLineBreakBox(lineBreakBox);
        }
        protected override void VisitLineNumberBox(LineNumberBox lineNumberBox)
        {
            AddTreeNode(lineNumberBox, ContentDisplayAction.Select);
            base.VisitLineNumberBox(lineNumberBox);
        }
        protected override void VisitNumberingListMarkBox(NumberingListMarkBox numberingListMarkBox)
        {
            AddTreeNode(numberingListMarkBox, ContentDisplayAction.Select);
            base.VisitNumberingListMarkBox(numberingListMarkBox);
        }
        protected override void VisitNumberingListWithSeparatorBox(NumberingListWithSeparatorBox numberingListWithSeparatorBox)
        {
            AddTreeNode(numberingListWithSeparatorBox, ContentDisplayAction.Select);
            base.VisitNumberingListWithSeparatorBox(numberingListWithSeparatorBox);
        }
        protected override void VisitPageBreakBox(PlainTextBox pageBreakBox)
        {
            AddTreeNode(pageBreakBox, ContentDisplayAction.Select);
            base.VisitPageBreakBox(pageBreakBox);
        }
        protected override void VisitPageNumberBox(PlainTextBox pageNumberBox)
        {
            AddTreeNode(pageNumberBox, ContentDisplayAction.Select);
            base.VisitPageNumberBox(pageNumberBox);
        }
        protected override void VisitParagraphMarkBox(PlainTextBox paragraphMarkBox)
        {
            AddTreeNode(paragraphMarkBox, ContentDisplayAction.Select);
            base.VisitParagraphMarkBox(paragraphMarkBox);
        }
        protected override void VisitPlainTextBox(PlainTextBox plainTextBox)
        {
            AddTreeNode(plainTextBox, ContentDisplayAction.Select);
            base.VisitPlainTextBox(plainTextBox);
        }
        protected override void VisitRangePermissionEndBox(RangePermissionBox rangePermissionEndBox)
        {
            AddTreeNode(rangePermissionEndBox, ContentDisplayAction.Select);
            base.VisitRangePermissionEndBox(rangePermissionEndBox);
        }
        protected override void VisitRangePermissionHighlightAreaBox(RangePermissionHighlightAreaBox rangePermissionHighlightAreaBox)
        {
            AddTreeNode(rangePermissionHighlightAreaBox, ContentDisplayAction.Select);
            base.VisitRangePermissionHighlightAreaBox(rangePermissionHighlightAreaBox);
        }
        protected override void VisitRangePermissionStartBox(RangePermissionBox rangePermissionStartBox)
        {
            AddTreeNode(rangePermissionStartBox, ContentDisplayAction.Select);
            base.VisitRangePermissionStartBox(rangePermissionStartBox);
        }
        protected override void VisitRow(LayoutRow row)
        {
            AddTreeNode(row, ContentDisplayAction.Select);
            base.VisitRow(row);
        }
        protected override void VisitSectionBreakBox(PlainTextBox sectionBreakBox)
        {
            AddTreeNode(sectionBreakBox, ContentDisplayAction.Select);
            base.VisitSectionBreakBox(sectionBreakBox);
        }
        protected override void VisitSpaceBox(PlainTextBox spaceBox)
        {
            AddTreeNode(spaceBox, ContentDisplayAction.Select);
            base.VisitSpaceBox(spaceBox);
        }
        protected override void VisitSpecialTextBox(PlainTextBox specialTextBox)
        {
            AddTreeNode(specialTextBox, ContentDisplayAction.Select);
            base.VisitSpecialTextBox(specialTextBox);
        }
        protected override void VisitStrikeoutBox(StrikeoutBox strikeoutBox)
        {
            AddTreeNode(strikeoutBox, ContentDisplayAction.Select);
            base.VisitStrikeoutBox(strikeoutBox);
        }
        protected override void VisitTable(LayoutTable table)
        {
            AddTreeNode(table, ContentDisplayAction.Select);
            base.VisitTable(table);
        }
        protected override void VisitTableCell(LayoutTableCell tableCell)
        {
            AddTreeNode(tableCell, ContentDisplayAction.Select);
            base.VisitTableCell(tableCell);
        }
        protected override void VisitTableRow(LayoutTableRow tableRow)
        {
            AddTreeNode(tableRow, ContentDisplayAction.Select);
            base.VisitTableRow(tableRow);
        }
        protected override void VisitTabSpaceBox(PlainTextBox tabSpaceBox)
        {
            AddTreeNode(tabSpaceBox, ContentDisplayAction.Select);
            base.VisitTabSpaceBox(tabSpaceBox);
        }
        protected override void VisitTextBox(LayoutTextBox textBox)
        {
            AddTreeNode(textBox, ContentDisplayAction.ScrollTo);
            base.VisitTextBox(textBox);
        }
        protected override void VisitUnderlineBox(UnderlineBox underlineBox)
        {
            AddTreeNode(underlineBox, ContentDisplayAction.Select);
            base.VisitUnderlineBox(underlineBox);
        }

        void AddTreeNode(LayoutElement element, ContentDisplayAction displayActionType)
        {
            TreeNode item = new TreeNode();
            // Store the attribute that determines whether document range selection is allowed for this node.
            item.Tag = displayActionType;
            Rectangle bounds = element.Bounds;
            item.ToolTipText = String.Format("X = {0}\nY = {1}\nWidth = {2}\nHeight = {3}", bounds.X, bounds.Y, bounds.Width, bounds.Height);
            // Update the layout element / tree node dictionary.
            Dictionary.Add(element, item);

            TreeNode parentItem = Dictionary[element.Parent];
            if (parentItem != null) 
            {
                // Add a new node to the tree.
                int index = parentItem.Nodes.Add(item);
                // Specify the node caption.
                string typeName = element.Type.ToString();
                switch (element.Type)
                {
                    case LayoutType.Column:
                        item.Text = String.Format("{0} #{1}", typeName, index);
                        break;
                    case LayoutType.Row:
                        item.Text = String.Format("{0} #{1}", typeName, index + 1);
                        break;
                    case LayoutType.TableRow:
                        item.Text = String.Format("{0} #{1}", typeName, index);
                        break;
                    default:
                        item.Text = typeName;
                        break;
                }

                if (parentItem.Tag != null)
                {
                    // As for the node that does not allow range selection in the document, all its child nodes should have the same attribute.
                    ContentDisplayAction parentDisplayActionType = (ContentDisplayAction)parentItem.Tag;
                    ContentDisplayAction actionType = parentDisplayActionType == ContentDisplayAction.ScrollTo ? parentDisplayActionType : displayActionType;
                    item.Tag = actionType;
                }
            }
        }
    }

    public enum ContentDisplayAction { Select, ScrollTo }
    #endregion #TreeViewCollector
}
