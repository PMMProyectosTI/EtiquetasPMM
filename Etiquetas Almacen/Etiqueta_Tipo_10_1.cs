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
    class Etiqueta_Tipo_10_1 : Etiqueta
    {
        protected  Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        public Image LogoPMM
        { get; set; }

        public string Direccion
        { get; set; }

        public Etiqueta_Tipo_10_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 532;
            this.Ancho = 786; 
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

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_10_1 etq)
        {
            obtenerDatos(etq.ClaveProducto);
            //Valores de prueba
            etq.Direccion = "Gasoducto #98 Col. Carrillo Puerto 76138, Queretaro, Qro.";

            
            int x = 25, y = 20;

            //Fuente de letra

            Font letraCampos = new Font("Arial", 12, FontStyle.Bold);
            Font letraLoteDespacho = new Font("Arial", 14, FontStyle.Bold);
            Font letraTitulo1 = new Font("Arial", 11, FontStyle.Bold);
            Font letraTitulo2 = new Font("Arial", 9);
            Font letraDireccion = new Font("Arial", 14);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 4);
            
            //Rectángulos 
            Rectangle rectangulo_contorno = new Rectangle(x, y, etq.Ancho, etq.Largo);
            Rectangle rect_superior = new Rectangle(x, y, etq.Ancho, 71);
            Rectangle rect_superior_2 = new Rectangle(x, rect_superior.Y + 71, etq.Ancho, 76);
            Rectangle rect_lateral_izquierdo = new Rectangle(x, rect_superior_2.Y + 76, 224, 358);
            Rectangle rect_central = new Rectangle(rect_lateral_izquierdo.X + 224, rect_superior_2.Y + 76, 340, 358);
            Rectangle rect_lateral_derecho_sup = new Rectangle(rect_central.X + 340, rect_superior_2.Y + 76, 221, 189);
            Rectangle rect_lateral_derecho_inf = new Rectangle(rect_central.X + 340, rect_lateral_derecho_sup.Y + 189, 221, 169);
            Rectangle rect_inferior = new Rectangle(x, rect_lateral_izquierdo.Y + 358, etq.Ancho, 27);

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
            gfx.DrawImage(etq.LogoPMM, rect_superior.X + 3, rect_superior.Y + 3, 265, 60);
            //TERMINA LOGO

            //Cliente
            x = 416;
            y = rect_superior.Y + 3;
            gfx.DrawString("Proveedora Mexicana de Monofilamentos S.A. de C.V.", letraTitulo1, Brush, new Point(x, y));
            gfx.DrawString("Oriente 217 No. 190, Agricola Oriental, 08500, Mexico City, Mexico.", letraTitulo2, Brush, new Point(x+11, y+15));
            gfx.DrawString("Tel. 00 5255 5763 8663 Fax 00 5255 5558 4483", letraTitulo2, Brush, new Point(x+118, y+30));
            gfx.DrawString("pmm@pmm-mex.com", letraTitulo2, Brush, new Point(x+263, y+45));
            
            //Campos en Recuadro lateral izquierdo
            x= rect_lateral_izquierdo.X + 10;
            y= rect_lateral_izquierdo.Y + 25;
            gfx.DrawString("CLAVE DE PRODUCTO:",letraCampos,Brush, new Point(x,y));
            gfx.DrawString(etq.ClaveProducto, letraCampos, Brush, new Point(rect_central.X + 250, y));
            y += 24;
            gfx.DrawString("ORDEN DE COMPRA:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("LOTE DE DESPACHO:", letraLoteDespacho, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("LOTE DE PRODUCCIÓN:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("PRODUCTO:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("MATERIAL:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Material, letraCampos, Brush, new Point(rect_central.X + 250, y));
            y += 24;
            gfx.DrawString("COLOR:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Color, letraCampos, Brush, new Point(rect_central.X + 250, y));
            y += 24;
            gfx.DrawString("CALIBRE:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Calibre, letraCampos, Brush, new Point(rect_central.X + 250, y));
            y += 24;
            gfx.DrawString("LARGO DE CORTE:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("PESO BRUTO:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("TARA:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("PESO NETO:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("FECHA:", letraCampos, Brush, new Point(x, y));
            y += 24;
            gfx.DrawString("EMBARQUE:", letraCampos, Brush, new Point(x, y));

            //Campos en el recuadro superior 2
            x = rect_superior_2.X + 10;
            y = rect_superior_2.Y + 3;
            gfx.DrawString("DIRECCIÓN:", letraDireccion, Brush, new Point(x, y));
            x += 125;
            gfx.DrawString(etq.Cliente, letraDireccion, Brush, new Point(x, y));
            y += 25; ;
            gfx.DrawString(etq.Direccion, letraDireccion, Brush, new Point(x, y));
            
            //Campos en el Recuadro Lateral Derecho Superior
            x = rect_lateral_derecho_sup.X + 55;
            y = rect_lateral_derecho_sup.Y + 27;
            gfx.DrawString("PESO BRUTO:", letraCampos, Brush, new Point(x,y));

            //Campos en el Recuadro Lateral Derecho Inferior
            x = rect_lateral_derecho_inf.X + 60;
            y = rect_lateral_derecho_inf.Y + 27;
            gfx.DrawString("PESO NETO:", letraCampos, Brush, new Point(x, y));

            //Campos en el cuandro inferior
            x = 373;
            y = rect_inferior.Y + 3;
            gfx.DrawString("CARTON:", letraLoteDespacho, Brush, new Point(x, y));


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
