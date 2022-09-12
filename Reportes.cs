using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf.security;
using System.Dynamic;
using System.Runtime.Remoting.Messaging;
using System.Runtime.CompilerServices;
using Org.BouncyCastle.Asn1.Cmp;

namespace proyectoAsync.Reportes
{
    public class Reporte
    {
        private Font _standardFont = new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK);
        private Font _NFont = new Font(Font.FontFamily.HELVETICA, 9, Font.BOLD, BaseColor.BLACK);
        private Font _textFont = new Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, BaseColor.BLACK);
        private int MarginDocumentTop = 20;
        private int MarginDocumentLeft = 15;
        private int MarginDocumentRigth = 15;
        private int MarginDocumentBottom = 20;

        public Reporte()
        {
            Document doc = new Document(PageSize.LETTER.Rotate(),MarginDocumentLeft, MarginDocumentRigth, MarginDocumentTop, MarginDocumentBottom);
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\edson\Desktop\prueba.pdf", FileMode.Create));

            doc.AddTitle("Mi primer PDF");
            doc.AddCreator("Edson Dimas");
            doc.Open();

            doc.Add(CreateTableTitle("REPORTE DE PRODUCCION DE BARNIZ U.V",30));

            doc.Add(CreateSpaceWhite(0));

            doc.Add(CreateTable1("NUMERO DE O.I","2207012","PROYECTO","cajilla 12 packs","FECHA","14-jul-22"));

            doc.Add(CreateSpaceWhite(0));

            doc.Add(CreateTable2("OPERADOR", "M. Baron G.", "HOJAS MAQUINAS", "3100", "HORA INICIO", "08:46"));
            doc.Add(CreateTable2("1ER AYUDANTE", "M. Baron G.", "HOJAS IMPRESAS FRENTE", "3100", "HORA TERMINO", "08:46"));
            doc.Add(CreateTable2("FEEDER", "M. Baron G.", "HOJAS IMPRESAS VUELTA", "3100", "MERMA", "08:46"));

            doc.Add(CreateSpaceWhite(0));

            doc.Add(CreateTableTitle("REGISTRO DE CONTROLES DE PROCESO", 15));
            doc.Add(CreateTable1("BARNIZ", "M. Baron G.", "TIPO DE SUSTRATO", "3100", "TRATAMINETO DE SUSTRATO", "08:46"));

            doc.Close();
            writer.Close();
        }
        private PdfPTable CreateSpaceWhite(int space)
        {
            PdfPTable SpaceTable = new PdfPTable(1);
            SpaceTable.WidthPercentage = 100;

            Phrase Text = new Phrase("a");
            Text.Font.Color = BaseColor.WHITE;
            Text.Font.Size = 5;
            PdfPCell nothing = new PdfPCell(Text);
            nothing.BorderWidth = 0;
            nothing.Padding = space;

            SpaceTable.AddCell(nothing);
            return SpaceTable;
        }
        private PdfPTable CreateTableTitle(String _title, int size)
        {
            //configuracion de tabla
            PdfPTable TableTitle = new PdfPTable(1);
            TableTitle.WidthPercentage = 100;

            // configuracion de letras
            Phrase title = new Phrase(_title);
            title.Font.Size = size;
            title.Font.IsBold();

            //Configuracion de celda
            PdfPCell clTitle = new PdfPCell(title);
            clTitle.HorizontalAlignment = Element.ALIGN_CENTER;
            clTitle.VerticalAlignment = Element.ALIGN_CENTER;
            clTitle.BorderWidth = 1;
            clTitle.PaddingBottom = 5;
            

            TableTitle.AddCell(clTitle);

            return TableTitle;
        }
        private PdfPTable CreateTable1(String Option1,String Value1,String Option2,String Value2,String Option3,String Value3)
        {
            float[] widths = new float[] { 4f, 7f,4f,7f,4f,7f };
            PdfPTable table = new PdfPTable(6);
            table.WidthPercentage = 100;
            table.SetWidths(widths);


            PdfPCell clOptionOrden = new PdfPCell(new Phrase(Option1, _NFont));
            clOptionOrden.BorderWidth = 0.75f;
            clOptionOrden.Padding = 5;
            clOptionOrden.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell clValueOrden = new PdfPCell(new Phrase(Value1, _standardFont));
            clValueOrden.BorderWidth = 0.75f;
            clValueOrden.Padding = 5;
            clValueOrden.HorizontalAlignment = Element.ALIGN_CENTER;
            clValueOrden.VerticalAlignment = Element.ALIGN_CENTER;

            PdfPCell clOptionProyecto = new PdfPCell(new Phrase(Option2, _NFont));
            clOptionProyecto.BorderWidth = 0.75f;
            clOptionProyecto.Padding = 5;
            clOptionProyecto.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell clValueProyecto = new PdfPCell(new Phrase(Value2, _textFont));
            clValueProyecto.BorderWidth = 0.75f;
            clValueProyecto.Padding = 5;
            clValueProyecto.HorizontalAlignment = Element.ALIGN_CENTER;
            clValueProyecto.VerticalAlignment = Element.ALIGN_CENTER;

            PdfPCell clOptionFecha = new PdfPCell(new Phrase(Option3, _NFont));
            clOptionFecha.BorderWidth = 0.75f;
            clOptionFecha.Padding = 5;
            clOptionFecha.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell clValueFecha = new PdfPCell(new Phrase(Value3, _standardFont));
            clValueFecha.BorderWidth = 0.75f;
            clValueFecha.Padding = 5;
            clValueFecha.HorizontalAlignment = Element.ALIGN_CENTER;
            clValueFecha.VerticalAlignment = Element.ALIGN_CENTER;



            table.AddCell(clOptionOrden);
            table.AddCell(clValueOrden);

            table.AddCell(clOptionProyecto);
            table.AddCell(clValueProyecto);

            table.AddCell(clOptionFecha);
            table.AddCell(clValueFecha);

            return table;
        }
        private PdfPTable CreateTable2(String Option1,String Value1,String Option2,String Value2,String Option3,String Value3)
        {
            float[] widths = new float[] { 4f, 7f, 4f, 7f, 4f, 7f };
            PdfPTable table = new PdfPTable(6);
            table.WidthPercentage = 100;
            table.SetWidths(widths);


            PdfPCell clOptionOrden = new PdfPCell(new Phrase(Option1, _NFont));
            clOptionOrden.BorderWidth = 0.75f;
            clOptionOrden.Padding = 5;
            clOptionOrden.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell clValueOrden = new PdfPCell(new Phrase(Value1, _standardFont));
            clValueOrden.BorderWidth = 0.75f;
            clValueOrden.Padding = 5;
            clValueOrden.HorizontalAlignment = Element.ALIGN_CENTER;
            clValueOrden.VerticalAlignment = Element.ALIGN_CENTER;

            PdfPCell clOptionProyecto = new PdfPCell(new Phrase(Option2, _NFont));
            clOptionProyecto.BorderWidth = 0.75f;
            clOptionProyecto.Padding = 5;
            clOptionProyecto.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell clValueProyecto = new PdfPCell(new Phrase(Value2, _standardFont));
            clValueProyecto.BorderWidth = 0.75f;
            clValueProyecto.Padding = 5;
            clValueProyecto.HorizontalAlignment = Element.ALIGN_CENTER;
            clValueProyecto.VerticalAlignment = Element.ALIGN_CENTER;

            PdfPCell clOptionFecha = new PdfPCell(new Phrase(Option3, _NFont));
            clOptionFecha.BorderWidth = 0.75f;
            clOptionFecha.Padding = 5;
            clOptionFecha.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell clValueFecha = new PdfPCell(new Phrase(Value3, _standardFont));
            clValueFecha.BorderWidth = 0.75f;
            clValueFecha.Padding = 5;
            clValueFecha.HorizontalAlignment = Element.ALIGN_CENTER;
            clValueFecha.VerticalAlignment = Element.ALIGN_CENTER;



            table.AddCell(clOptionOrden);
            table.AddCell(clValueOrden);

            table.AddCell(clOptionProyecto);
            table.AddCell(clValueProyecto);

            table.AddCell(clOptionFecha);
            table.AddCell(clValueFecha);

            return table;
        }
    }
}
