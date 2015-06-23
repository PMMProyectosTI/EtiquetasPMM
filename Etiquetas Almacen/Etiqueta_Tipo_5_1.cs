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
    class Etiqueta_Tipo_5_1 : Etiqueta
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

        public Etiqueta_Tipo_5_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 399;
            this.Ancho = 403; 
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
                conn.Close();
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

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_5_1 etq)
        {
            int x = 25, y = 20;

            //Fuente de letra
            Font letraCliente = new Font("Arial", 18);
            Font letraGrande = new Font("Arial", 11, FontStyle.Bold);
            Font letraMuyGrande = new Font("Arial", 20, FontStyle.Bold);
            Font letraCampos = new Font("Arial", 8, FontStyle.Bold);
            Font letraCarton = new Font("Arial", 12, FontStyle.Bold);
            Font letraWeights = new Font("Arial", 10, FontStyle.Bold);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);
     
            
            //Rectángulos 
            Rectangle rectangulo_contorno = new Rectangle(x, y, etq.Ancho, etq.Largo);
            Rectangle rect_superior = new Rectangle(x, y, etq.Ancho, 68);
            Rectangle rect_superior_2 = new Rectangle(x, rect_superior.Y + 68, etq.Ancho, 64);
            Rectangle rect_lateral_izquierdo = new Rectangle(x, rect_superior_2.Y + 64, 95, 231);
            Rectangle rect_central = new Rectangle(rect_lateral_izquierdo.X + 95, rect_superior_2.Y + 64, 155, 231);
            Rectangle rect_lateral_derecho_sup = new Rectangle(rect_central.X + 155, rect_superior_2.Y + 64, 153, 108);
            Rectangle rect_lateral_derecho_inf = new Rectangle(rect_central.X + 155, rect_lateral_derecho_sup.Y + 108, 153, 123);
            Rectangle rect_inferior = new Rectangle(x, rect_lateral_izquierdo.Y + 231, etq.Ancho, 36);

            gfx.DrawRectangle(pluma, rectangulo_contorno);
            gfx.DrawRectangle(pluma, rect_superior);
            gfx.DrawRectangle(pluma, rect_superior_2);
            gfx.DrawRectangle(pluma, rect_lateral_izquierdo);
            gfx.DrawRectangle(pluma, rect_central);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_sup);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_inf);
            gfx.DrawRectangle(pluma, rect_inferior);
            //TERMINAN RECTÁNGULOS

            //LOGO
            //TERMINA LOGO

            
            etq.Cliente = "Braun Oral-B Ireland Ltd.";
            //Cliente
            Size textSize = TextRenderer.MeasureText(etq.Cliente, letraCliente);
            x = (int)getCenterXcoordinate(rect_superior.X + etq.Ancho, textSize.Width, rect_superior.X) + 48;
            gfx.DrawString(etq.Cliente, letraCliente, Brush, new Point(x, rect_superior.Y + 18));
            

            //Campos en Recuadro lateral izquierdo
            x= rect_lateral_izquierdo.X + 1/2;
            y= rect_lateral_izquierdo.Y + 3;

            //El interlineado varía mucho de campo a campo porque se está intentando imitar por completo una etiqueta
            gfx.DrawString("MATERIAL:", letraGrande, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("COLOR:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("CALIPER:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("CUT LENGTH:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("P. O. No.:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("DISP. LOT:", letraGrande, Brush, new Point(x, y));
            y += 25; 
            gfx.DrawString("PROD. LOT:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("DATE:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("NET WEIGHT:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("GROSS WEIGHT:", letraCampos, Brush, new Point(x, y));
            y += 21;
            gfx.DrawString("TARE:", letraCampos, Brush, new Point(x, y));
            //y += 35;
            //gfx.DrawString("", letraCampos, Brush, new Point(x, y));
            //gfx.DrawImage();

            //Campos en el Recuadro Inferior
            x = rect_inferior.X + 75;
            y = rect_inferior.Y + 8; 
            gfx.DrawString("CARTON:", letraCarton, Brush, new Point(x, y));

            //Campos en el Recuadro Lateral Derecho Superior
            int acarreox = rect_lateral_derecho_sup.Width / 2;
            x = rect_lateral_derecho_sup.X + acarreox - 63;
            y = rect_lateral_derecho_sup.Y + 3; 
            gfx.DrawString("GROSS WEIGHT:", letraWeights, Brush, new Point(x,y));

            //Campos en el Recuadro Lateral Derecho Inferior
            x = rect_lateral_derecho_inf.X + acarreox - 50;
            y = rect_lateral_derecho_inf.Y + 3;
            gfx.DrawString("NET WEIGHT:", letraWeights, Brush, new Point(x, y));

            //Campos en el Segundo Recuadro Superior
            x = rect_superior_2.X + 8;
            y = rect_superior_2.Y + 16;
            gfx.DrawString("ITEM:", letraMuyGrande, Brush, new Point(x, y));

            //Campos en el Primer Recuadro
            //Image imageFile = Image.FromFile("PNG.jpg");
           // gfx.DrawImage(newImage, xi, yi, widthi, heighti);
            //gfx.DrawImage(Image imageFile, Rectangle rect_Superior, 25, 25, 100, 100);
            //e.Graphics.DrawImage(imageFile, new Point(22,27));
           

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
