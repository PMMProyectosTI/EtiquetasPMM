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
    class Etiqueta_Tipo_12_1 : Etiqueta
    {
        protected  Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        protected string hoja = "Hoja de Especificaciones grales";
        protected OleDbConnection conn = new OleDbConnection();

        

        public string Fecha
        { get; set; }

        public Etiqueta_Tipo_12_1(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 268;
            this.Ancho = 523; 
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

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_12_1 etq)
        {
            obtenerDatos(etq.ClaveProducto);
            //valores de prueba
            string numero = "00060";
            int po = 100535;
            string fecha = "20/09/2013";
            string prueba = "0.016 GREEN 454 NYLON 6.12 \"R\" \"X\" Shape";
            int contador = 1;
            int x = 125, y = 25;

            //Fuente de letra

            Font letraCampos = new Font("Arial", 14);
            Font letraNumero = new Font("Arial", 52, FontStyle.Bold);
            Font letraDatos = new Font("Arial", 14, FontStyle.Bold);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);

            
            while (contador < 3 )
            {
                if (contador == 1)
                    {
                        x = 125;
                        y = 25;
                    }
                else
                    { 
                        x = 125;
                        y = 380;
                    }

                //Rectángulos             
                Rectangle rect_superior = new Rectangle(x, y, etq.Ancho, etq.Largo);

                gfx.DrawRectangle(pluma, rect_superior);

                //Campos en Recuadro lateral izquierdo
                x = rect_superior.X + 20;
                y = rect_superior.Y + 45;
                gfx.DrawString("BRM P/N:", letraCampos, Brush, new Point(x, y));
                gfx.DrawString(numero, letraNumero, Brush, new Point(x+115, y-30));
                y = rect_superior.Y + 180;
                gfx.DrawString(prueba, letraDatos, Brush, new Point(x+20, y));
                y = rect_superior.Y + 215;
                gfx.DrawString("P.O.:", letraCampos, Brush, new Point(x, y));
                gfx.DrawString(po.ToString(), letraDatos, Brush, new Point(x+50, y));
                x = rect_superior.X + 300;
                gfx.DrawString("Packaging", letraCampos, Brush, new Point(x, y-10));
                gfx.DrawString("Date:", letraCampos, Brush, new Point(x+45, y+10));
                gfx.DrawString(fecha, letraDatos, Brush, new Point(x + 100, y));

                contador++;
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
