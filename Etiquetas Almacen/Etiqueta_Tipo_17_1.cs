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
    class Etiqueta_Tipo_17_1 :  Etiqueta
    {
         protected  Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        public string Fecha
        { get; set; }

        public Etiqueta_Tipo_17_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 882;
            this.Ancho = 770; 
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

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_17_1 etq)
        {
            obtenerDatos(etq.ClaveProducto);
            //valores de prueba
            string texto = "EB20 FILAMENTS";


            int x = 10, y = 25;

            //Fuente de letra

            Font letraCampos = new Font("Arial", 46, FontStyle.Bold);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 2);


            //Rectángulos             
            Rectangle rect_contorno = new Rectangle(x, y, etq.Ancho, etq.Largo);
            Rectangle rect_1 = new Rectangle(x, rect_contorno.Y, etq.Ancho, 147);
            Rectangle rect_2 = new Rectangle(x, rect_1.Y + 147, etq.Ancho, 147);
            Rectangle rect_3 = new Rectangle(x, rect_2.Y + 147, etq.Ancho, 147);
            Rectangle rect_4 = new Rectangle(x, rect_3.Y + 147, etq.Ancho, 147);
            Rectangle rect_5 = new Rectangle(x, rect_4.Y + 147, etq.Ancho, 147);
            Rectangle rect_6 = new Rectangle(x, rect_5.Y + 147, etq.Ancho, 147);

            gfx.DrawRectangle(pluma, rect_contorno);
            gfx.DrawRectangle(pluma, rect_1);
            gfx.DrawRectangle(pluma, rect_2);
            gfx.DrawRectangle(pluma, rect_3);
            gfx.DrawRectangle(pluma, rect_4);
            gfx.DrawRectangle(pluma, rect_5);
            gfx.DrawRectangle(pluma, rect_6);

            //Campos en Recuadro lateral izquierdo
            x = rect_contorno.X + 100;
            y = rect_contorno.Y + 40;
            for (int i = 0; i < 6; i++)
            {
                gfx.DrawString(texto, letraCampos, Brush, new Point(x, y));
                y += 147;
            }

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
