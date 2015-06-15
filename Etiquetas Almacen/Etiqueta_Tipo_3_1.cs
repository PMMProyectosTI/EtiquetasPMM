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
    class Etiqueta_Tipo_3_1 : Etiqueta
    {
        protected Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        /*protected int xi = 40;//int xi = 35;
        protected int yi = 35;//int yi = 35;
        protected int widthi = 120;//int widthi = 125;
        protected int heighti = 34;//int heighti = 40;*/

        public Image LogoPMM
        { get; set; }

        public Image ISO9001
        { get; set; }

        public Image ISO14001
        { get; set; }  
        public Etiqueta_Tipo_3_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 308;
            this.Ancho = 751;
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data" + @" Source=" + path.LocalPath + "; Extended Properties='Excel 12.0 Macro;HDR=NO;'";
            this.LogoPMM = Image.FromFile(@"C:/GitHub/EtiquetasPMM/logos/PMM_Logo.png");
            this.ISO14001 = Image.FromFile(@"C:/GitHub/EtiquetasPMM/logos/ISO14001_BN.jpg");
            this.ISO9001 = Image.FromFile(@"C:/GitHub/EtiquetasPMM/logos/ISO9001_BN.jpg");
        }

        public bool obtenerDatos(string cp)
        {
            OleDbCommand get = new OleDbCommand();
            get.Connection = conn;
            get.CommandText = "SELECT F2,F3,F4,F5,F6 FROM [" + hoja + "$] WHERE F1='" + cp + "'";
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
            Font letraGrande = new Font("Arial", 19, FontStyle.Bold);
            Font letraCampos = new Font("Arial", 15, FontStyle.Bold);
            Font letraCampoGrande = new Font("Arial", 16, FontStyle.Bold); //usado para texto "DISPATCH LOT:"
            Font letraCampoChico = new Font("Arial", 10, FontStyle.Bold); //usado para texto "PRODUCTION LOT:"

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);

            //Rectángulos 
            Rectangle rectangulo_contorno = new Rectangle(x, y, etq.Ancho, etq.Largo);
            Rectangle rect_superior = new Rectangle(x, y, etq.Ancho, 95);
            Rectangle rect_superior_2 = new Rectangle(x, rect_superior.Y + rect_superior.Height, etq.Ancho, 37);
            Rectangle rect_lateral_izquierdo = new Rectangle(x, rect_superior_2.Y + rect_superior_2.Height, 372, 176);
            Rectangle rect_lateral_derecho_1 = new Rectangle(rect_lateral_izquierdo.X + rect_lateral_izquierdo.Width, rect_superior_2.Y + rect_superior_2.Height, 379 , 67);
            Rectangle rect_lateral_derecho_2 = new Rectangle(rect_lateral_izquierdo.X + rect_lateral_izquierdo.Width, rect_lateral_derecho_1.Y + rect_lateral_derecho_1.Height, 379, 55);
            Rectangle rect_lateral_derecho_3 = new Rectangle(rect_lateral_izquierdo.X + rect_lateral_izquierdo.Width, rect_lateral_derecho_2.Y + rect_lateral_derecho_2.Height, 379, 54);

            gfx.DrawRectangle(pluma, rectangulo_contorno);
            gfx.DrawRectangle(pluma, rect_superior);
            gfx.DrawRectangle(pluma, rect_superior_2);
            gfx.DrawRectangle(pluma, rect_lateral_izquierdo);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_1);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_2);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_3);
            

            x = rect_superior_2.X + 4;
            y = rect_superior_2.Y + 2;

            gfx.DrawString("ITEM:", letraGrande, Brush, new Point(x, y));
            Size textSize = TextRenderer.MeasureText(etq.ClaveProducto, letraGrande);
            x = (int)getCenterXcoordinate(rect_superior_2.X + etq.Ancho, textSize.Width, rect_superior_2.X);
            gfx.DrawString(etq.ClaveProducto, letraGrande, Brush, new Point(x , y));
            

            //texto en recuadro lateral izquierdo (material, color, caliper, dispatch lot, production lot)
            x = rect_lateral_izquierdo.X + 4;
            y = rect_lateral_izquierdo.Y + 4;
            gfx.DrawString("MATERIAL:", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("COLOR:", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("CALIPER:", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("DISPATCH LOT:", letraCampoGrande, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("PRODUCTION LOT:", letraCampoChico, Brush, new Point(x, y));
            y += 38;

            x = rect_lateral_izquierdo.X + 193;
            y = rect_lateral_izquierdo.Y + 4;
            //estos valores son de prueba, deberán ser reemplazados por etq.Propiedad
            gfx.DrawString("NYLON 6.12", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("INDICATOR", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("0.007 in", letraCampos, Brush, new Point(x, y));
            y += 36;
            gfx.DrawString("12K131", letraCampoGrande, Brush, new Point(x, y));
            y += 37;
            gfx.DrawString("4N6122JJ01", letraCampoChico, Brush, new Point(x, y));
            ////////////////////////////////////////////////////////////////////////////////

            x = rect_lateral_derecho_1.X + 10;
            textSize = TextRenderer.MeasureText("NET WEIGHT:",letraCampos);
            y = (int)getCenterYcoordinate(rect_lateral_derecho_1.Y + rect_lateral_derecho_1.Height, textSize.Height, rect_lateral_derecho_1.Y);
            gfx.DrawString("NET WEIGHT:", letraCampos, Brush, new Point(x, y));

            x = rect_lateral_derecho_2.X + 10;
            textSize = TextRenderer.MeasureText("GROSS WEIGHT:", letraCampos);
            y = (int)getCenterYcoordinate(rect_lateral_derecho_2.Y + rect_lateral_derecho_2.Height, textSize.Height, rect_lateral_derecho_2.Y);
            gfx.DrawString("GROSS WEIGHT:", letraCampos, Brush, new Point(x, y));

            x = rect_lateral_derecho_3.X + 45;
            textSize = TextRenderer.MeasureText("CARTON:", letraCampoGrande);
            y = (int)getCenterYcoordinate(rect_lateral_derecho_3.Y + rect_lateral_derecho_3.Height, textSize.Height, rect_lateral_derecho_3.Y);
            gfx.DrawString("CARTON:", letraCampoGrande, Brush, new Point(x, y));

            

            /////////valroes de prueba
            textSize = TextRenderer.MeasureText("8.16 Kg.", letraCampos);
            y = (int)getCenterYcoordinate(rect_lateral_derecho_1.Y + rect_lateral_derecho_1.Height, textSize.Height, rect_lateral_derecho_1.Y);
            x = (int)getRightXcoordinate(rect_lateral_derecho_1.X + rect_lateral_derecho_1.Width, textSize.Width, rect_lateral_derecho_1.X) - 10;
            gfx.DrawString("8.16 Kg.", letraCampos, Brush, new Point(x, y));

            textSize = TextRenderer.MeasureText("9.22 Kg.", letraCampos);
            y = (int)getCenterYcoordinate(rect_lateral_derecho_2.Y + rect_lateral_derecho_2.Height, textSize.Height, rect_lateral_derecho_2.Y);
            x = (int)getRightXcoordinate(rect_lateral_derecho_2.X + rect_lateral_derecho_2.Width, textSize.Width, rect_lateral_derecho_2.X) - 10;
            gfx.DrawString("9.22 Kg.", letraCampos, Brush, new Point(x, y));

            textSize = TextRenderer.MeasureText("1 : 36", letraCampoGrande);
            y = (int)getCenterYcoordinate(rect_lateral_derecho_3.Y + rect_lateral_derecho_3.Height, textSize.Height, rect_lateral_derecho_3.Y);
            x = rect_lateral_derecho_3.X + 200;
            gfx.DrawString("1 : 36", letraCampoGrande, Brush, new Point(x, y));

            
            //Imágenes
            gfx.DrawImage(etq.LogoPMM, rectangulo_contorno.X + 5, rectangulo_contorno.Y + 2, 356, 77);
            gfx.DrawImage(etq.ISO9001, 580, rectangulo_contorno.Y + 4, rect_superior.Height - 8, rect_superior.Height - 8);
            gfx.DrawImage(etq.ISO14001, 686, rectangulo_contorno.Y + 4, rect_superior.Height - 8, rect_superior.Height - 8);
            
            return e;
        }

        protected double getCenterXcoordinate(int etiquetaX, int textoX, int rectX)//anchura en pixeles del texto y de la etiqueta
        {
            double x = 0;
            x = (0.5 * (etiquetaX + rectX)) - (0.5 * textoX);
            return x;
        }
        protected double getRightXcoordinate(int etiquetaX, int elementX, int rectX)//anchura en pixeles del texto y de la etiqueta
        {
            double x = 0;
            x = rectX + etiquetaX - elementX - 7;
            while (x > etiquetaX - elementX)
            {
                x--;
            }
            return x;//se le quitan 10 pixeles para que no esté pegado a la derecha
        }
        protected double getCenterYcoordinate(int etiquetaY, int textoY, int rectY)//anchura en pixeles del texto y de la etiqueta
        {
            double y = 0;
            y = (0.5 * (etiquetaY + rectY)) - (0.5 * textoY);
            return y;
        }


    }
}
