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
    class Etiqueta_Tipo_6_1 : Etiqueta
    {
        protected  Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        //protected   Image newImage = Image.FromFile("C:/Users/Juan/Desktop/Logos/PNG.png");
        //Image newImage = Image.FromFile("Logos.png");

        protected  int xi = 40;//int xi = 35;
        protected int yi = 35;//int yi = 35;
        protected int widthi = 120;//int widthi = 125;
        protected int heighti = 34;//int heighti = 40;

        public Etiqueta_Tipo_6_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 450;
            this.Ancho = 643; 
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data" + @" Source=" + path.LocalPath + "; Extended Properties='Excel 12.0 Macro;HDR=NO;'";
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

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_6_1 etq)
        {
            int x = 42, y = 20;

            //Fuente de letra
            Font letraGrande = new Font("Arial", 22, FontStyle.Bold);
            Font letraMedia = new Font("Arial", 20, FontStyle.Bold);
            Font letraPequena = new Font("Arial", 10, FontStyle.Bold);
            Font letraCliente = new Font("Arial", 22, FontStyle.Underline);
            Font letraFecha = new Font("Arial", 14, FontStyle.Bold);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);
     
            
            //Rectángulos 
            Rectangle rectangulo_1 = new Rectangle(x, y, etq.Ancho, 1);
            Rectangle rectangulo_2 = new Rectangle(x, etq.Largo, etq.Ancho, 1);
            

            gfx.DrawRectangle(pluma, rectangulo_1);
            gfx.DrawRectangle(pluma, rectangulo_2);

            etq.Cliente = "ORAL-B LABORATORIES";
            //Cliente

            //Campos en Recuadro
            gfx.DrawString("PALLET INTERNA", letraPequena, Brush, new Point(560, 30));
            gfx.DrawString("FECHA:", letraFecha, Brush, new Point(487, 60));
            x = rectangulo_1.X + 11;
            y= rectangulo_1.Y + 100;
            //El interlineado varía mucho de campo a campo porque se está intentando imitar por completo una etiqueta
            gfx.DrawString("CLIENTE:", letraGrande, Brush, new Point(x, y));
            gfx.DrawString(etq.Cliente, letraCliente, Brush, new Point(x + 163, y));
            y += 46;
            gfx.DrawString("PRODUCTO:", letraGrande, Brush, new Point(x, y));
            y += 79;
            gfx.DrawString("PMM:", letraGrande, Brush, new Point(x, y));
            gfx.DrawString("ENTREGA:", letraGrande, Brush, new Point(x+288, y));
            y += 52;
            gfx.DrawString("No. PEDIDO CLIENTE:", letraPequena, Brush, new Point(x, y));
            gfx.DrawString("N° DE CAJAS:", letraPequena, Brush, new Point(x+341, y));
            y += 32;
            gfx.DrawString("EMBARQUE:", letraPequena, Brush, new Point(x, y));
            gfx.DrawString("CLAVE:", letraMedia, Brush, new Point(x+290, y));
            y += 64;
            gfx.DrawString("P.NETO:", letraGrande, Brush, new Point(x, y));
            
            //gfx.DrawImage();

            //Campos en el Primer Recuadro
            //Image imageFile = Image.FromFile("PNG.jpg");
           // gfx.DrawImage(newImage, xi, yi, widthi, heighti);
            //gfx.DrawImage(Image imageFile, Rectangle rect_Superior, 25, 25, 100, 100);
            //e.Graphics.DrawImage(imageFile, new Point(22,27));
           

            return e;
        }

    }
}
