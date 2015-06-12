using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Drawing.Printing;
using System.Drawing;
using System.Windows.Forms;
namespace Etiquetas_Almacen
{
    class Etiqueta_Tipo_3_1 : Etiqueta
    {
        protected  Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        public Etiqueta_Tipo_3_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 311;
            this.Ancho = 751; 
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

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_3_1 etq)
        {
            int x = 25, y = 20;

            //Fuente de letra
            Font letraCliente = new Font("Arial", 22);
            Font letraGrande = new Font("Arial", 16, FontStyle.Bold);
            Font letraCampos = new Font("Arial", 12, FontStyle.Bold);      

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);
            
            //Rectángulos 
            Rectangle rectangulo_contorno = new Rectangle(x, y, etq.Ancho, etq.Largo);
            
            
            /*Rectangle rect_superior = new Rectangle(x, y, etq.Ancho, 70);
            Rectangle rect_superior_2 = new Rectangle(x, rect_superior.Y + 70, etq.Ancho, 92);//c
            Rectangle rect_lateral_izquierdo = new Rectangle(x, rect_superior_2.Y + 92, 175, 379);
            Rectangle rect_central = new Rectangle(rect_lateral_izquierdo.X + 175, rect_superior_2.Y + 92, 373, 379);
            Rectangle rect_lateral_derecho_sup = new Rectangle(rect_central.X + 373, rect_superior_2.Y + 92, 238, 211);
            Rectangle rect_lateral_derecho_inf = new Rectangle(rect_central.X + 373, rect_lateral_derecho_sup.Y + 211, 238, 168);
            Rectangle rect_inferior = new Rectangle(x, rect_lateral_izquierdo.Y + 379, etq.Ancho, 30);

            gfx.DrawRectangle(pluma, rectangulo_contorno);

            //TERMINAN RECTÁNGULOS

            //LOGO
            //TERMINA LOGO


            etq.Cliente = "Braun Kronberg";
            //Cliente
            Size textSize = TextRenderer.MeasureText(etq.Cliente, letraCliente);
            x = (int)getCenterXcoordinate(rect_superior.X + etq.Ancho, textSize.Width, rect_superior.X);
            gfx.DrawString(etq.Cliente, letraCliente, Brush, new Point(x, rect_superior.Y + 18));
            

            //Campos en Recuadro lateral izquierdo
            x= rect_lateral_izquierdo.X + 9;
            y= rect_lateral_izquierdo.Y + 9;

            //El interlineado varía mucho de campo a campo porque se está intentando imitar por completo una etiqueta
            gfx.DrawString("ITEM No.:",letraGrande,Brush, new Point(x,y));
            y += 45;
            gfx.DrawString("MATERIAL:", letraCampos, Brush, new Point(x, y));
            y += 39;
            gfx.DrawString("COLOR:", letraCampos, Brush, new Point(x, y));
            y += 42;
            gfx.DrawString("CALIPER:", letraCampos, Brush, new Point(x, y));
            y += 42;
            gfx.DrawString("CUT LENGTH:", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("P. O. No.:", letraCampos, Brush, new Point(x, y));
            y += 30;
            gfx.DrawString("PRODUCTION REF:", letraCampos, Brush, new Point(x, y));
            y += 27;
            gfx.DrawString("DISPATCH DATE:", letraCampos, Brush, new Point(x, y));
            y += 27;
            gfx.DrawString("GROSS WEIGHT:", letraCampos, Brush, new Point(x, y));
            y += 29;
            gfx.DrawString("TARE:", letraCampos, Brush, new Point(x, y));
            y += 28;
            gfx.DrawString("NET WEIGHT:", letraCampos, Brush, new Point(x, y));
            //y += 35;
            //gfx.DrawString("", letraCampos, Brush, new Point(x, y));
            //gfx.DrawImage();

            //Campos en el Recuadro Inferior
            x = (int)getCenterXcoordinate(rect_inferior.X + etq.Ancho, textSize.Width, rect_inferior.X) + 92;
            gfx.DrawString("CARTON:", letraGrande, Brush, new Point(x, (etq.Largo - rect_inferior.Y)*(1/2)+564));

            //Campos en el Recuadro Lateral Derecho Superior
            int acarreox = rect_lateral_derecho_sup.Width / 2;
            x = rect_lateral_derecho_sup.X + acarreox - 70;
            y = rect_lateral_derecho_sup.Y + 13; 
            gfx.DrawString("GROSS WEIGHT:", letraCampos, Brush, new Point(x,y));

            //Campos en el Recuadro Lateral Derecho Inferior
            x = rect_lateral_derecho_inf.X + acarreox - 57;
            //x = rect_lateral_derecho_inf.X + 50;
            y = rect_lateral_derecho_inf.Y + 29;
            gfx.DrawString("NET WEIGHT:", letraCampos, Brush, new Point(x, y));

            //Campos en el Segundo Recuadro Superior
            textSize = TextRenderer.MeasureText("LOT:", letraGrande);
            x = (int)getCenterXcoordinate(rect_superior_2.X + etq.Ancho, textSize.Width, rect_superior_2.X) + 129;
            gfx.DrawString("LOT:", letraGrande, Brush, new Point(x, rect_superior_2.Y + 27));


            int xi = 40;//int xi = 35;
            int yi = 35;//int yi = 35;
            int widthi = 120;//int widthi = 125;
            int heighti = 34;//int heighti = 40;
            //Campos en el Primer Recuadro
            //Image imageFile = Image.FromFile("PNG.jpg");
           // gfx.DrawImage(newImage, xi, yi, widthi, heighti);
            //gfx.DrawImage(Image imageFile, Rectangle rect_Superior, 25, 25, 100, 100);
            //e.Graphics.DrawImage(imageFile, new Point(22,27));*/
           


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
