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
using Org.BouncyCastle.Asn1.X509;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Linq.Expressions;
using System.Runtime.InteropServices;

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

        public Reporte(int indicateReport, string[][] value)
        {
            switch (indicateReport)
            {
                case 1://reporte de produccion Barniz
                    CreateReportProduccionBarniz();
                    break;
                case 2://reporte de control de proceso
                    CreateReportControlProcess(value);
                    break;
                default:
                    throw new ArgumentException("value default does not exist");
            }
        }
        //FUNCTION CREATE
        private void CreateReportProduccionBarniz()
        {
            //sup
            Document doc = new Document(PageSize.LETTER.Rotate(), MarginDocumentLeft, MarginDocumentRigth, MarginDocumentTop, MarginDocumentBottom);
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\edson\Desktop\prueba.pdf", FileMode.Create));

            doc.AddTitle("Mi primer PDF");
            doc.AddCreator("Edson Dimas");
            doc.Open();

            doc.Add(CreateTableTitle("REPORTE DE PRODUCCION DE BARNIZ U.V", 30));

            doc.Add(CreateSpaceWhite(0));

            doc.Add(CreateTable1("NUMERO DE O.I", "2207012", "PROYECTO", "cajilla 12 packs", "FECHA", "14-jul-22"));

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
        private void CreateReportControlProcess(string[][] value)
        {
            Document doc = new Document(PageSize.LETTER.Rotate(), MarginDocumentLeft, MarginDocumentRigth, MarginDocumentTop, MarginDocumentBottom);
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\edson\Desktop\prueba.pdf", FileMode.Create));
            Image image = Image.GetInstance(@"Picture\image.jpg");

            //image1.ScalePercent(50f);
            image.ScaleAbsoluteWidth(150);
            image.ScaleAbsoluteHeight(35);

            doc.AddTitle("Mi primer PDF");
            doc.AddCreator("Edson Dimas");
            doc.Open();

            doc.Add(CreateTableControlProcessTitle(
                image,
                "CONTROL EN PROCESO BARNIZ U.V",
                "CODIGO:ROP-11-01",_NFont));
            doc.Add(CreateTableControlProcess(
                "FECHA",
                "ORDEN DE IMPRESION",
                "NOMBRE DEL TRABAJO",
                "PRUEBA DE FROTE",
                "PRUEBA DE DOBLEZ",
                "CURADO DE BARNIZ",
                "POLVO",
                "REPINTE",
                "MATERIAL FRESCO",
                "OBSERVACIONES",_standardFont));
            for(int i = 0; i < value.Length; i++)
            {
                doc.Add(CreateTableControlProcess(
                    value[i][0], 
                    value[i][1], 
                    value[i][2], 
                    value[i][3], 
                    value[i][4], 
                    value[i][5], 
                    value[i][6], 
                    value[i][7], 
                    value[i][8],
                    value[i][9],
                    _standardFont));
            }
            doc.Add(CreateTableControlProcessLeyend("C:CUMPLE","NC:NO CUMPLE","N/A:NO APLICA",_NFont));
            doc.Add(new Phrase("NOTA: ESTE CONTROL EN PROCESO DEBE DE REALIZARSE CADA 300 HOJAS PROCESADAS SIN IMPORTAR LA CANTIDAD DE PEDIDO", _NFont));
            
            doc.Close();
            writer.Close();
        }
        private void CreateReportBitacora(string[][] value)
        {
            Document doc = new Document(PageSize.LETTER.Rotate(), MarginDocumentLeft, MarginDocumentRigth, MarginDocumentTop, MarginDocumentBottom);
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\edson\Desktop\prueba.pdf", FileMode.Create));
            Image image = Image.GetInstance(@"Picture\image.jpg");

            //image1.ScalePercent(50f);
            image.ScaleAbsoluteWidth(150);
            image.ScaleAbsoluteHeight(35);

            doc.AddTitle("Mi primer PDF");
            doc.AddCreator("Edson Dimas");

            for(int i = 0; i < value.Length; i++)
            {
                doc.Add(CreateTableReportBitacora(
                    value[i][0],
                    value[i][1],
                    value[i][2],
                    value[i][3],
                    value[i][4],
                    value[i][5],
                    _NFont,
                    _standardFont));
            }

            doc.Open();
            doc.Close();
            writer.Close();
        }
        //FUNCTION DRAW
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
        private PdfPTable CreateTable1(String Option1, string Value1,String Option2,String Value2,String Option3,String Value3)
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
        private PdfPTable CreateTableControlProcessTitle(Image a1, string a2, string a3, Font _font )
        {
            float[] widths = new float[] { 9, 12, 6 };
            PdfPTable table = new PdfPTable(3);
            table.WidthPercentage = 100;
            table.SetWidths(widths);

            PdfPCell pdfPCell1 = new PdfPCell(a1);
            pdfPCell1.BorderWidth = 0.75f;
            pdfPCell1.Padding = 0;
            pdfPCell1.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell2 = new PdfPCell(new Phrase(a2, _font));
            pdfPCell2.BorderWidth = 0.75f;
            pdfPCell2.Padding = 10;
            pdfPCell2.VerticalAlignment = Element.ALIGN_CENTER;
            pdfPCell2.HorizontalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell3 = new PdfPCell(new Phrase(a3, _font));
            pdfPCell3.BorderWidth = 0.75f;
            pdfPCell3.Padding = 10;
            pdfPCell3.VerticalAlignment = Element.ALIGN_CENTER;
            pdfPCell2.HorizontalAlignment = Element.ALIGN_CENTER;

            table.AddCell(pdfPCell1);
            table.AddCell(pdfPCell2);
            table.AddCell(pdfPCell3);

            return table;
        }
        private PdfPTable CreateTableControlProcess(string a1,string a2,string a3,string a4, string a5, string a6, string a7,string a8,string a9,string a10,Font _font)
        {
            int pading = 5;
            float[] widht = new float[] {2,2,5,2,2,2,2,2,2,6 };
            PdfPTable table = new PdfPTable(10);
            table.WidthPercentage = 100;
            table.SetWidths(widht);
            

            PdfPCell pdfPCell1 = new PdfPCell(new Phrase(a1,_font));
            pdfPCell1.BorderWidth = 0.75f;
            pdfPCell1.PaddingBottom = pading;
            pdfPCell1.PaddingTop = pading;

            pdfPCell1.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell1.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell2 = new PdfPCell(new Phrase(a2, _font));
            pdfPCell2.BorderWidth = 0.75f;
            pdfPCell2.Padding = pading;
            pdfPCell2.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell2.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell3 = new PdfPCell(new Phrase(a3, _font));
            pdfPCell3.BorderWidth = 0.75f;
            pdfPCell3.Padding = pading;
            pdfPCell3.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell3.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell4 = new PdfPCell(new Phrase(a4,_font));
            pdfPCell4.BorderWidth = 0.75f;
            pdfPCell4.Padding = pading;
            pdfPCell4.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell4.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell5 = new PdfPCell(new Phrase(a5, _font));
            pdfPCell5.BorderWidth = 0.75f;
            pdfPCell5.Padding = pading;
            pdfPCell5.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell5.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell6 = new PdfPCell(new Phrase(a6, _font));
            pdfPCell6.BorderWidth = 0.75f;
            pdfPCell6.Padding = pading;
            pdfPCell6.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell6.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell7 = new PdfPCell(new Phrase(a7, _font));
            pdfPCell7.BorderWidth = 0.75f;
            pdfPCell7.Padding = pading;
            pdfPCell7.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell7.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell8 = new PdfPCell(new Phrase(a8, _font));
            pdfPCell8.BorderWidth = 0.75f;
            pdfPCell8.Padding = pading;
            pdfPCell8.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell8.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell9 = new PdfPCell(new Phrase(a9, _font));
            pdfPCell9.BorderWidth = 0.75f;
            pdfPCell9.Padding = pading;
            pdfPCell9.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell9.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell10 = new PdfPCell(new Phrase(a10, _font)); 
            pdfPCell10.BorderWidth = 0.75f;
            pdfPCell10.Padding = pading;
            pdfPCell10.VerticalAlignment = Element.ALIGN_CENTER;

            table.AddCell(pdfPCell1);
            table.AddCell(pdfPCell2);
            table.AddCell(pdfPCell3);
            table.AddCell(pdfPCell4);
            table.AddCell(pdfPCell5);
            table.AddCell(pdfPCell6);
            table.AddCell(pdfPCell7);
            table.AddCell(pdfPCell8);
            table.AddCell(pdfPCell9);
            table.AddCell(pdfPCell10);

            return table;
        }
        private PdfPTable CreateTableControlProcessLeyend(String a1,String a2,string a3, Font _font)
        {
            float[] Widths = new float[] {2,2,2,7}; 
            PdfPTable table = new PdfPTable(4);
            table.WidthPercentage = 100;
            table.SetWidths(Widths);

            PdfPCell pdfPCell1 = new PdfPCell(new Phrase(a1, _font));
            pdfPCell1.BorderWidth = 0;
            pdfPCell1.Padding = 5;
            pdfPCell1.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell2 = new PdfPCell(new Phrase(a2, _font));
            pdfPCell2.BorderWidth = 0;
            pdfPCell2.Padding = 5;
            pdfPCell2.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell3 = new PdfPCell(new Phrase(a3, _font));
            pdfPCell3.BorderWidth = 0;
            pdfPCell3.Padding = 5;
            pdfPCell3.VerticalAlignment = Element.ALIGN_CENTER;
            PdfPCell pdfPCell4 = new PdfPCell(new Phrase("", _font));
            pdfPCell4.BorderWidth = 0;
            pdfPCell4.Padding = 5;
            pdfPCell4.VerticalAlignment = Element.ALIGN_CENTER;

            table.AddCell(pdfPCell1);
            table.AddCell(pdfPCell2);
            table.AddCell(pdfPCell3);
            table.AddCell(pdfPCell4);
            return table;
        }
        private PdfPTable CreateTableReportBitacoraTitle(string a1,string a2,Font font)
        {
            int padding = 5;
            float[] widths = new float[] { 3, 5 };
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 100;
            table.SetWidths(widths);

            PdfPCell pdfPCell1 = new PdfPCell(new Phrase(a1,font));
            pdfPCell1.BackgroundColor = BaseColor.GRAY;
            pdfPCell1.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfPCell1.VerticalAlignment = Element.ALIGN_CENTER;
            pdfPCell1.PaddingBottom = padding;
            pdfPCell1.PaddingTop = padding;

            PdfPCell pdfPCell2 = new PdfPCell(new Phrase(a2,font));
            pdfPCell2.BackgroundColor = BaseColor.GRAY;
            pdfPCell1.VerticalAlignment = Element.ALIGN_CENTER;
            pdfPCell2.HorizontalAlignment = Element.ALIGN_CENTER;

            table.AddCell(pdfPCell1);
            table.AddCell(pdfPCell2);

            return table; 
        }
        private PdfPTable CreateTableReportBitacora(string t1, string a1,string t2,string a2,string t3,string a3,Font title,Font text)
        {
            int padding = 3;
            float[] Widths = new float[] {2,3,2,3,2,3};  
            PdfPTable table = new PdfPTable(6);
            table.WidthPercentage = 100;
            table.SetWidths(Widths);

            PdfPCell pdfPCell1 = new PdfPCell(new Phrase(t1,title));
            pdfPCell1.BorderWidth = 0.75f;
            pdfPCell1.PaddingBottom = padding;
            pdfPCell1.PaddingTop = padding;
            PdfPCell pdfPCell2 = new PdfPCell(new Phrase(a1,text));
            pdfPCell2.BorderWidth = 0.75f;
            PdfPCell pdfPCell3 = new PdfPCell(new Phrase(t2,title));
            pdfPCell3.BorderWidth = 0.75f;
            PdfPCell pdfPCell4 = new PdfPCell(new Phrase(a2,text));
            pdfPCell4.BorderWidth = 0.75f;
            PdfPCell pdfPCell5 = new PdfPCell(new Phrase(t3,title));
            pdfPCell5.BorderWidth = 0.75f;
            PdfPCell pdfPCell6 = new PdfPCell(new Phrase(a3,text));
            pdfPCell6.BorderWidth = 0.75f;
            return table;
        }
    }
}

