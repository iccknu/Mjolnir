using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace Office
{
    public class MyWord
    {
        WordprocessingDocument wordProcessingDocument;
        MainDocumentPart mainDocumentPart;
        Body body;
        string path;

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
