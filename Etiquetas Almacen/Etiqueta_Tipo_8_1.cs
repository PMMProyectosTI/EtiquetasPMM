using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Printing;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Etiquetas_Almacen
{
    class Etiqueta_Tipo_8_1 : Etiqueta
    {
       //cambios 1 desde JuanPC
        //respuesta a cambio 1 desde MathusPC
        protected  Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        //protected   Image newImage = Image.FromFile("C:/Users/Juan/Desktop/Logos/PNG.png");
        //Image newImage = Image.FromFile("Logos.png");

        protected  int xi = 40;//int xi = 35;
        protected int yi = 35;//int yi = 35;
        protected int widthi = 120;//int widthi = 125;
        protected int heighti = 34;//int heighti = 40;

        public Image LogoPMM
        { get; set; }

        public Etiqueta_Tipo_8_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 587;
            this.Ancho = 403; 
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data" + @" Source=" + path.LocalPath + "; Extended Properties='Excel 12.0 Macro;HDR=NO;'";
            this.LogoPMM = Image.FromFile(@"C:/GitHub/EtiquetasPMM/logos/PMM_Logo.png");
        }

        public bool obtenerDatos(string cp)
        {
            OleDbCommand get = new OleDbCommand();
            get.Connection = conn;
            get.CommandText = "SELECT F2,F3,F4,F5,F6 FROM ["+hoja+"$] WHERE F1='"+cp+"'";
            try
            {
                conn.Open();
                OleDbDataReader reader = get.ExecuteReader();
                while (reader.Read())
                {
                    this.Cliente = reader["F2"].ToString();
                    this.Material = reader["F3"].ToString();
                    this.Calibre = reader["F4"].ToString();
                    this.Color = reader["F5"].ToString();
                    this.Corte = reader["F6"].ToString();
                }
                reader.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                if (conn.State == System.Data.ConnectionState.Open)
                    conn.Close();
                MessageBox.Show(ex.Message);
                return false;
            }
            
        }

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_8_1 etq)
        {
            int x = 25, y = 20;

            //Fuente de letra
            Font letraInfo = new Font("Arial", 6);
            Font letraCampos = new Font("Arial", 8, FontStyle.Bold);
            Font letraIndices = new Font("Arial", 14, FontStyle.Bold);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);
     
            
            //Rectángulos 
            Rectangle rectangulo_contorno = new Rectangle(x, y, etq.Ancho, etq.Largo);
            Rectangle rect_superior_1 = new Rectangle(x, y, 205, 126);
            Rectangle rect_superior_2 = new Rectangle(x + 205, y, 198, 126);
            Rectangle rectangulo_central_1 = new Rectangle(x, y + 126, etq.Ancho, 140);
            Rectangle rectangulo_central_2 = new Rectangle(x, rectangulo_central_1.Y + 140, etq.Ancho, 165);
            Rectangle rectangulo_central_3 = new Rectangle(x, rectangulo_central_2.Y + 165, etq.Ancho, 156);

            gfx.DrawRectangle(pluma, rectangulo_contorno);
            gfx.DrawRectangle(pluma, rect_superior_1);
            gfx.DrawRectangle(pluma, rect_superior_2);
            gfx.DrawRectangle(pluma, rectangulo_central_1);
            gfx.DrawRectangle(pluma, rectangulo_central_2);
            gfx.DrawRectangle(pluma, rectangulo_central_3);

            gfx.DrawImage(etq.LogoPMM, rectangulo_contorno.X + 35, rectangulo_contorno.Y + 3, 168, 40);

            //El interlineado varía mucho de campo a campo porque se está intentando imitar por completo una etiqueta
            gfx.DrawString("From:", letraCampos, Brush, new Point(x, y+3));
            gfx.DrawString("To:", letraCampos, Brush, new Point(x+205, y + 3));

            gfx.DrawString("Proveedora Mexicana de", letraCampos, Brush, new Point (x + 40, y + 45));
            gfx.DrawString("Monofilamentos S.A de C.V.", letraCampos, Brush, new Point(x + 40, y + 60));
            gfx.DrawString("Oriente 217 No. 190, Agricola Oriental,", letraInfo, Brush, new Point(x + 2, y + 80));
            gfx.DrawString("08500, Mexico City, Mexico.", letraInfo, Brush, new Point(x + 2, y + 90));
            gfx.DrawString("Tel. 00 5255 5763 8663 Fax 00 5255 5558 4483", letraInfo, Brush, new Point(x+ 2, y + 100));
            gfx.DrawString("pmm@pmm-mex.com", letraInfo, Brush, new Point(x+ 2, y + 110));

            x = rectangulo_central_1.X + 90;
            y = rectangulo_central_1.Y + 30;
            gfx.DrawString("PO:", letraIndices, Brush, new Point(x, y));

            x = rectangulo_central_2.X + 90;
            y = rectangulo_central_2.Y + 30;
            gfx.DrawString("SKU:", letraIndices, Brush, new Point(x, y));

            x = rectangulo_central_3.X + 90;
            y = rectangulo_central_3.Y + 30;
            gfx.DrawString("QTY:", letraIndices, Brush, new Point(x, y));          

            return e;
        }

        protected double getCenterXcoordinate(int etiquetaX, int textoX, int rectX)//anchura en pixeles del texto y de la etiqueta
        {
            double x = 0;
            x = (0.5 * (etiquetaX + rectX)) - (0.5 * textoX);
            return x;
        }
    }
}
