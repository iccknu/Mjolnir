using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;

namespace Office
{
    public enum MyJustification { Left, Center, Right }

    public enum MyTypographicalEmphasis { Normal, Bold }

    public class MyWord
    {
        WordprocessingDocument wordProcessingDocument;
        MainDocumentPart mainDocumentPart;
        Body body;
        string path;
        public string BodyXml { get { return mainDocumentPart.Document.Body.OuterXml; } }

        public MyWord(string path, MyWordDocumentProperties pageProp = null)
        {
            //initializing word document object
            this.path = path;
            wordProcessingDocument = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
            mainDocumentPart = wordProcessingDocument.AddMainDocumentPart();
            mainDocumentPart.Document = new Document();
            body = mainDocumentPart.Document.AppendChild(new Body());

            if (pageProp == null) return;

            //set some document properties
            PageSize pageSize = GetPageSize(pageProp);

            body.AppendChild(new SectionProperties(pageSize,
                new PageMargin()
                {
                    Top = pageProp.Margins.TopMargin,
                    Bottom = pageProp.Margins.BottomMargin,
                    Left = (uint)pageProp.Margins.LeftMargin,
                    Right = (uint)pageProp.Margins.RightMargin
                }
            ));
        }

        public MyWord(string documentToClonePath, string newDocumentPath)
        {
            this.path = newDocumentPath;
            using (var mainDoc = WordprocessingDocument.Open(documentToClonePath, false))
            {
                this.wordProcessingDocument = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
                // copy parts from source document to new document
                foreach (var part in mainDoc.Parts)
                    wordProcessingDocument.AddPart(part.OpenXmlPart, part.RelationshipId);
                // Save не работает -- почему?
                //this.wordProcessingDocument.Dispose();
                //this.wordProcessingDocument = WordprocessingDocument.Open(newDocumentPath, true);
                this.mainDocumentPart = wordProcessingDocument.MainDocumentPart;
                this.body = this.wordProcessingDocument.MainDocumentPart.Document.Body;
            }
        }

        public void InsertAPicture(string fileName)
        {
            ImagePart imagePart = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
            imagePart.FeedData(stream);
            stream.Close();

            Bitmap bm = new Bitmap(fileName);

            AddImageToBody(mainDocumentPart.GetIdOfPart(imagePart), (long)bm.Width * (long)((float)400000 / bm.HorizontalResolution),
               (long)bm.Height * (long)((float)400000 / bm.VerticalResolution));
        }


        private void AddImageToBody(string relationshipId, long width, long height)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = width, Cy = height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState = A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = width, Cy = height }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordProcessingDocument.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }


        /// <summary>
        /// Searches for the specific substring in the file and replaces it.
        /// </summary>
        /// <param name="oldString">String to replace.</param>
        /// <param name="newString">String that will be placed instead the old one.</param>
        public void Replace(string oldString, string newString)
        {
            var r = body.Elements<Paragraph>();

            foreach (var p in r)
            {
                if (p.InnerText.Contains(oldString))
                {
                    string text = p.InnerText.Replace(oldString, newString);

                    ParagraphProperties prop = p.ParagraphProperties;
                    RunProperties runprop = (RunProperties)p.Elements<Run>().ToList()[0].RunProperties?.CloneNode(true);

                    p.RemoveAllChildren();
                    p.Append(prop);
                    p.Append(new Run(runprop, new Text(text)));
                }
            }
        }

        /// <summary>
        /// Searches for the specific substring in the document and replaces it with saving styles
        /// </summary>
        /// <param name="oldString">String to replace.</param>
        /// <param name="newString">String that will be placed instead the old one.</param>
        public void ReplaceWord(string oldString, string newString)
        {
            var allParagraphs = wordProcessingDocument.MainDocumentPart.Document.Body.Elements<Paragraph>();
            foreach (Paragraph p in allParagraphs)
            {
                ReplaceTextInParagraph(oldString, newString, p);
            }
        }

        /// <summary>
        /// Searches for the specific substring in paragraph and replaces it. 
        /// </summary>
        /// <param name="oldString"></param>
        /// <param name="newString"></param>
        /// <param name="p"></param>
        private static void ReplaceTextInParagraph(string oldString, string newString, Paragraph p)
        {
            //the main problem is that the text in paragraph is randomly distributed in Run elements
            //and oldString can be placed in several Runs
            if (oldString.Length == 0) throw new Exception("Empty string to replace");
            var allchildren = p.Elements<Run>().ToList();
            string[] texts = new string[allchildren.Count];
            for (int i = 0; i < allchildren.Count; i++)
            {
                texts[i] = allchildren[i].InnerText;
            }
            string textInParagraph = String.Join("", texts, 0, texts.Length);
            if (!textInParagraph.Contains(oldString)) return;
            //we iterating through characters just to be sure we do not touch already processed characters
            for (int i = 0; i < textInParagraph.Length; i++)
            {
                int pos = textInParagraph.IndexOf(oldString, i); //zerobased index of first character of oldString in textInParagraph
                if (pos == -1) break;
                int startIndex = 0; // index of first run block which contains (at least part of) oldString
                int endIndex = 0; // index of last run block which contains (at least part of) oldString
                int ax = 0;
                for (startIndex = 0; startIndex < texts.Length; startIndex++)
                {
                    ax += texts[startIndex].Length;
                    if (ax > pos) { break; }
                }
                ax = 0;
                for (endIndex = 0; endIndex < texts.Length; endIndex++)
                {
                    ax += texts[endIndex].Length;
                    if (ax > pos + oldString.Length - 1) { break; }
                }
                //at this point blocks from startindex to endIndex contains oldString whitin
                {
                    int relativepos = pos; //zerobased index of first character of oldstring relative to startindex block
                    //use relativepos to correctly replace oldstring (for example it helps to replace oldstring only once when oldstring is equal to newstring)

                    for (int j = 0; j < startIndex; j++)
                    {
                        relativepos -= texts[j].Length;
                    }
                    string text = String.Join("", texts, startIndex, endIndex - startIndex + 1);
                    string leftPart = text.Substring(0, relativepos + oldString.Length);
                    string rightPart = text.Substring(relativepos + oldString.Length);
                    texts[startIndex] = leftPart;
                    if (endIndex > startIndex) texts[endIndex] = rightPart;
                    else texts[startIndex] += rightPart;
                    //erasing bloks between startindex and endIndex, because all information contains in startindex block and endIndex block
                    for (int j = startIndex + 1; j < endIndex; j++)
                    {
                        texts[j] = "";
                        allchildren[j].Remove();
                    }
                    //write all information (including replaced string) to startindex block and endIndex block
                    allchildren[startIndex].RemoveAllChildren<Text>();
                    allchildren[startIndex].Append(new Text(texts[startIndex].Replace(oldString, newString)));
                    texts[startIndex] = texts[startIndex].Replace(oldString, newString);
                    if (endIndex > startIndex)
                    {
                        allchildren[endIndex].RemoveAllChildren<Text>();
                        allchildren[endIndex].Append(new Text(texts[endIndex]));
                    }
                }
                textInParagraph = String.Join("", texts, 0, texts.Length); // refresh textInParagraph whith replaced word
                i = pos + newString.Length - 1; // set counter just at the end of replaced word
            }
        }

        public void AppendRow(string[] keyWordsRow, string[] newWordsRow)
        {
            TableRow row = FindRow(keyWordsRow);
            if (row == null) return;
            Table tb = row.Parent as Table;
            TableRow newRow = (TableRow)row.CloneNode(deep: true);
            TableCell[] cells = newRow.Elements<TableCell>().ToArray();
            for (int i = 0; i < cells.Length; i++)
            {
                ReplaceTextInParagraph(keyWordsRow[i], newWordsRow[i], cells[i].Elements<Paragraph>().First());
            }
            tb.AppendChild(newRow);
        }

        public void DeleteRow(string[] keyWordsRow)
        {
            TableRow row = FindRow(keyWordsRow);
            if (row != null)
            {
                row.Remove();
            }
        }

        /// <summary>
        /// Returns first occurence of TableRow that have exactly same words
        /// </summary>
        /// <param name="keyWordsRow"></param>
        /// <returns></returns>
        TableRow FindRow(string[] keyWordsRow)
        {
            Stack<OpenXmlElement> stack = new Stack<OpenXmlElement>();
            stack.Push(body);
            while (stack.Count > 0)
            {
                OpenXmlElement el = stack.Pop();
                foreach (var child in el.Elements())
                {
                    stack.Push(child);
                }
                if (!(el is Table)) continue;
                Table tb = el as Table;
                TableRow[] rows = tb.Elements<TableRow>().ToArray();
                foreach (var row in rows)
                {
                    if (row.Elements<TableCell>().Count() != keyWordsRow.Length) continue;
                    List<string> words = new List<string>(keyWordsRow.Length);
                    foreach (TableCell cell in row.Elements<TableCell>().ToArray())
                    {
                        words.Add(cell.InnerText);
                    }
                    bool same = true;
                    for (int i = 0; i < words.Count; i++)
                    {
                        if (words[i] != keyWordsRow[i])
                        {
                            same = false;
                            break;
                        }
                    }
                    if (same)
                    {
                        return row;
                    }
                }
            }
            return null;
        }

        public void Save()
        {
            mainDocumentPart.Document.Save();
            wordProcessingDocument.Save();
            //rewriting whole file
            using (Stream stream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                using (OpenXmlPackage docfile = wordProcessingDocument.Clone(stream))
                {
                    docfile.Save();
                }
            }//stream is not automaticaly closed when document closed (when using Clone)   
        }

        public void Close()
        {
            Save();
            wordProcessingDocument.Close();
        }

        private PageSize GetPageSize(MyWordDocumentProperties pageProp)
        {
            PageSize pageSize = new PageSize();
            switch (pageProp.PageSize)
            {
                case MyPageSize.Letter:
                    pageSize.Width = 12240; pageSize.Height = 15840;
                    break;
                default:
                    pageSize.Width = 11906; pageSize.Height = 16838; // A4 page size
                    break;
            }
            if (pageProp.PageOrientation == MyPageOrientation.Portrait)
            {
                pageSize.Orient = PageOrientationValues.Portrait;
            }
            else
            {
                pageSize.Orient = PageOrientationValues.Landscape;
                UInt32Value temp = pageSize.Width;
                pageSize.Width = pageSize.Height;
                pageSize.Height = temp;
            }
            return pageSize;
        }

        public void AddParagraph(string text, int textSize, string font, MyTypographicalEmphasis typoEmp, MyJustification justification)
        {
            RunProperties runProp = GetRunProperties(font, textSize, typoEmp);
            Run run = new Run(runProp, new Text(text));

            ParagraphProperties paraProp = new ParagraphProperties(new Justification()
            {
                Val = justification == MyJustification.Left ? JustificationValues.Left :
                        justification == MyJustification.Right ? JustificationValues.Right :
                        JustificationValues.Center
            });

            Paragraph paragraph = new Paragraph(paraProp, run);

            body.AppendChild(paragraph);
        }

        public void AddParagraph(string text, int fontSize, string font)
        {
            AddParagraph(text, fontSize, font, MyTypographicalEmphasis.Normal, MyJustification.Left);
        }

        public void AddParagraph(string text)
        {
            AddParagraph(text, 11, "Arial", MyTypographicalEmphasis.Normal, MyJustification.Left);
        }

        public void AddTable(MyTable myTable)
        {
            Table table = new Table(GetTableProperties());

            TableRow headersRow = new TableRow(new TableRowProperties(new TableHeader()));
            FillRow(headersRow, myTable.Headers);
            table.Append(headersRow);

            foreach (var row in myTable.Rows)
            {
                TableRow tableRow = new TableRow();
                FillRow(tableRow, row);
                table.Append(tableRow);
            }

            body.Append(table);
        }


        private void FillRow(TableRow targetRow, MyTableRow sourceRow)
        {
            foreach (var cell in sourceRow.Cells)
            {
                TableCellProperties tableCellProperties = new TableCellProperties(
                    new TableCellMargin()
                    {
                        LeftMargin = new LeftMargin() { Width = cell.Style.LeftPadding.ToString() },
                        RightMargin = new RightMargin() { Width = cell.Style.RightPadding.ToString() },
                        TopMargin = new TopMargin() { Width = cell.Style.TopPadding.ToString() },
                        BottomMargin = new BottomMargin() { Width = cell.Style.BottomPadding.ToString() }
                    },
                    new GridSpan() { Val = cell.Style.Colspan }
                );
                TableCell tableCell = new TableCell(tableCellProperties);

                ParagraphProperties paraProp = new ParagraphProperties(new Justification()
                {
                    Val = cell.Style.Justification == MyJustification.Left ? JustificationValues.Left :
                        cell.Style.Justification == MyJustification.Right ? JustificationValues.Right :
                        JustificationValues.Center
                })
                { SpacingBetweenLines = new SpacingBetweenLines() { After = "50", Line = "240" } };

                RunProperties runProp = GetRunProperties(cell.Style.Font, cell.Style.FontSize, cell.Style.Emphasis);

                Paragraph para = new Paragraph(paraProp, new Run(runProp, new Text(cell.Text)));
                tableCell.Append(para);
                targetRow.Append(tableCell);
            }
        }

        private TableProperties GetTableProperties()
        {
            TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 }
                    ),
                    new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct }
                );
            return tblProp;
        }

        private RunProperties GetRunProperties(string font, int fontSize, MyTypographicalEmphasis typoEmp)
        {
            RunProperties runProp = new RunProperties(
                new FontSize() { Val = (fontSize * 2).ToString() },
                new RunFonts() { Ascii = font, HighAnsi = font }
            );
            if (typoEmp == MyTypographicalEmphasis.Bold) runProp.Bold = new Bold();
            return runProp;
        }

    }

    public class MyTableCell
    {
        public CellStyle Style { get; private set; }
        public string Text { get; private set; }
        public MyTableCell(string text)
        {
            Text = text;
            this.Style = new CellStyle();
        }
        public MyTableCell(string text, CellStyle style)
        {
            Text = text;
            this.Style = style;
        }
    }



    public class MyTableRow
    {
        public List<MyTableCell> Cells { get; set; } = new List<MyTableCell>();

        public MyTableRow() { }

        public MyTableRow(IEnumerable<string> data)
        {
            foreach (string cellText in data)
            {
                Cells.Add(new MyTableCell(cellText));
            }
        }

        public MyTableRow(IEnumerable<MyTableCell> cells)
        {
            AddCells(cells);
        }

        public void AddCells(IEnumerable<MyTableCell> cells)
        {
            foreach (MyTableCell cell in cells)
            {
                AddCell(cell);
            }
        }

        public void AddCell(MyTableCell cell)
        {
            Cells.Add(cell);
        }
    }



    public class MyTable
    {
        public MyTableRow Headers { get; set; }
        public List<MyTableRow> Rows { get; set; } = new List<MyTableRow>();

        public MyTable(MyTableRow headers)
        {
            Headers = headers;
        }

        public MyTable(MyTableRow headers, IEnumerable<MyTableRow> rows)
        {
            Headers = headers;
            AddRows(rows);
        }

        public void AddRows(IEnumerable<MyTableRow> rows)
        {
            foreach (MyTableRow row in rows)
            {
                AddRow(row);
            }
        }

        public void AddRow(MyTableRow row)
        {
            Rows.Add(row);
        }
    }

	
    public class CellStyle
    {
        public short[] Padding { get; set; }

        public short TopPadding
        {
            get { return (Padding is null) ? (short)5 : Padding[0]; }
            set { Padding[0] = (value > 0) ? value : (short)0; }
        }
        public short RightPadding
        {
            get { return (Padding is null) ? (short)5 : Padding[1]; }
            set { Padding[1] = (value > 0) ? value : (short)0; }
        }
        public short BottomPadding
        {
            get { return (Padding is null) ? (short)5 : Padding[2]; }
            set { Padding[2] = (value > 0) ? value : (short)0; }
        }
        public short LeftPadding
        {
            get { return (Padding is null) ? (short)5 : Padding[3]; }
            set { Padding[3] = (value > 0) ? value : (short)0; }
        }

        public string Font { get; set; }
        public int FontSize { get; set; }
        public MyJustification Justification { get; set; }
        public MyTypographicalEmphasis Emphasis { get; set; }
        public short Colspan { get; set; }

        public CellStyle() : this(null, "Times New Roman", 12,
            MyJustification.Left, MyTypographicalEmphasis.Normal, 1)
        { }


        public CellStyle(short[] padding = null, string font = "Times New Roman", int fontSize = 12,
            MyJustification justification = MyJustification.Left,
            MyTypographicalEmphasis emphasis = MyTypographicalEmphasis.Normal, short colspan = 1)
        {
            Font = font;
            FontSize = fontSize;
            Justification = justification;
            Emphasis = emphasis;
            Colspan = colspan;

            if (padding is null)
                padding = new short[] { 0, 0, 0, 0 };
            Padding = padding;
        }
    }

	
    public enum MyPageOrientation { Portrait, Landscape }

    public enum MyPageSize { Letter, A4 }

    public class MyMargin
    {
        public int LeftMargin { get; set; }
        public int RightMargin { get; set; }
        public int TopMargin { get; set; }
        public int BottomMargin { get; set; }

        public MyMargin() { }

        public MyMargin(int LeftMargin, int RightMargin, int TopMargin, int BottomMargin)
        {
            this.LeftMargin = LeftMargin;
            this.RightMargin = RightMargin;
            this.TopMargin = TopMargin;
            this.BottomMargin = BottomMargin;
        }
    }

    public class MyWordDocumentProperties
    {
        public MyPageOrientation PageOrientation { get; set; }
        public MyPageSize PageSize { get; set; }
        public MyMargin Margins { get; set; }

        public MyWordDocumentProperties()
        {
            PageOrientation = MyPageOrientation.Portrait;
            PageSize = MyPageSize.A4;
            Margins = new MyMargin(850, 850, 850, 850);
        }
    }

}
