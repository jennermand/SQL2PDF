using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQL2PDFReport
{
    public class PdfPageHelper : PdfPageEventHelper
    {
        //public override void OnEndPage(PdfWriter writer, iTextSharp.text.Document document)
        //{
        //    base.OnEndPage(writer, document);
        //}

        //public override void OnStartPage(PdfWriter writer, iTextSharp.text.Document document)
        //{
        //    base.OnStartPage(writer, document);
        //}
        /*
 * We use a __single__ Image instance that's assigned __once__;
 * the image bytes added **ONCE** to the PDF file. If you create 
 * separate Image instances in OnEndPage()/OnEndPage(), for example,
 * you'll end up with a much bigger file size.
 */
        public Image ImageHeader { get; set; }
        public Image ImageFooter { get; set; }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            // cell height 
            float cellHeight = document.TopMargin;
            // PDF document size      
            Rectangle page = document.PageSize;
            
            // create two column table
            PdfPTable head = new PdfPTable(1);
            head.TotalWidth = page.Width-50;

            // add image; PdfPCell() overload sizes image to fit cell
            PdfPCell c = new PdfPCell(ImageHeader, true);
            c.HorizontalAlignment = Element.ALIGN_RIGHT;
            c.VerticalAlignment = Element.ALIGN_BOTTOM;
            c.FixedHeight = cellHeight;
            c.Border = PdfPCell.NO_BORDER;
            head.AddCell(c);

            //// add the header text
            //c = new PdfPCell(new Phrase(
            //  DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + " GMT"
            //));
            //c.Border = PdfPCell.NO_BORDER;
            //c.VerticalAlignment = Element.ALIGN_BOTTOM;
            //c.FixedHeight = cellHeight;
            //head.AddCell(c);

            // since the table header is implemented using a PdfPTable, we call
            // WriteSelectedRows(), which requires absolute positions!
            head.WriteSelectedRows(
              0, -1,  // first/last row; -1 flags all write all rows
              25,      // left offset
                // ** bottom** yPos of the table
              page.Height - cellHeight + head.TotalHeight,
              
              writer.DirectContent
            );
        }
    }
}
