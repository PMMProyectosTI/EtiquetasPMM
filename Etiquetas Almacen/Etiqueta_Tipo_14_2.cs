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
    class Etiqueta_Tipo_14_2 : Etiqueta_Tipo_14_1
    {
       
        public Etiqueta_Tipo_14_2(string cp) : base(cp)
        {
            this.ClaveProducto = cp;
            this.Largo = 512;
            this.Ancho = 772; 
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data" + @" Source=" + path.LocalPath + "; Extended Properties='Excel 12.0 Macro;HDR=NO;'";
            this.LogoPMM = Image.FromFile(@"C:/GitHub/EtiquetasPMM/logos/PMM_Logo.png");
        }

        public PrintPageEventArgs dibujarEtiqueta(PrintPageEventArgs e, Etiqueta_Tipo_14_2 etq)
        {
            obtenerDatos(etq.ClaveProducto);
            //Valores de prueba
            etq.Direccion = "Gasoducto #98 Col. Carrillo Puerto";
            etq.Contenido = "xxxxxxxxxxx";
            etq.Producto = "xxxxxxxxxxxx";
            etq.OrdenCompra = "xxxxxxxxxx";
            etq.PesoBruto = "xxxxxxxxxxxx";
            etq.PesoNeto = "xxxxxxxxxxx";
            etq.Tara = "xxxxxxxxxxxxxxx";
            etq.Fecha = "xxxxxxxxxxx";

            int x = 37, y = 24;

            //Fuente de letra

            Font letraCampos = new Font("Arial", 12, FontStyle.Bold);
            Font letraLoteDespacho = new Font("Arial", 14, FontStyle.Bold);
            Font letraTitulo1 = new Font("Arial", 11, FontStyle.Bold);
            Font letraTitulo2 = new Font("Arial", 11);
            Font letraDireccion = new Font("Arial", 16);

            Graphics gfx = e.Graphics;
            SolidBrush Brush = new SolidBrush(System.Drawing.Color.Black);
            Pen pluma = new Pen(System.Drawing.Color.Black, 3);
            
            //Rectángulos             
            Rectangle rect_superior = new Rectangle(x, y, etq.Ancho, 135);
            Rectangle rect_superior_2 = new Rectangle(x, rect_superior.Y + 135, etq.Ancho, 259);

            Rectangle rectangulo_contorno = new Rectangle(x ,rect_superior_2.Y + 297, etq.Ancho, etq.Largo);
            Rectangle rect_lateral_izquierdo = new Rectangle(x, rectangulo_contorno.Y, 179, 482);
            Rectangle rect_central = new Rectangle(rect_lateral_izquierdo.X + 179, rect_lateral_izquierdo.Y, 355, 482);
            Rectangle rect_lateral_derecho_sup = new Rectangle(rect_central.X + 355, rect_lateral_izquierdo.Y, 238, 215);
            Rectangle rect_lateral_derecho_inf = new Rectangle(rect_central.X + 355, rect_lateral_derecho_sup.Y + 215, 238, 267);
            Rectangle rect_inferior = new Rectangle(x, rect_lateral_izquierdo.Y + 482, etq.Ancho, 30);

            gfx.DrawRectangle(pluma, rectangulo_contorno);
            gfx.DrawRectangle(pluma, rect_superior_2);
            gfx.DrawRectangle(pluma, rect_lateral_izquierdo);
            gfx.DrawRectangle(pluma, rect_central);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_sup);
            gfx.DrawRectangle(pluma, rect_lateral_derecho_inf);
            gfx.DrawRectangle(pluma, rect_inferior);
            //TERMINAN RECTÁNGULOS

            //LOGO
            gfx.DrawImage(etq.LogoPMM, rect_superior.X + 3, rect_superior.Y + 3, 355, 84);
            //TERMINA LOGO

            //PMM
            x = 368;
            y = rect_superior.Y + 3;
            gfx.DrawString("Proveedora Mexicana de Monofilamentos S.A. de C.V.", letraTitulo1, Brush, new Point(x+43, y));
            gfx.DrawString("Oriente 217 No. 190, Agricola Oriental,", letraTitulo2, Brush, new Point(x+ 170, y+=22));
            gfx.DrawString("08500, Mexico City, Mexico.", letraTitulo2, Brush, new Point(x + 242, y+=22));
            gfx.DrawString("Tel. 00 5255 5763 8663", letraTitulo2, Brush, new Point(x+268, y+=22));
            gfx.DrawString("Fax 00 5255 5558 4483", letraTitulo2, Brush, new Point(x + 268, y += 22));
            gfx.DrawString("pmm@pmm-mex.com", letraTitulo2, Brush, new Point(x+275, y+=22));

            //CLIENTE
            x = rect_superior_2.X + 25;
            y = rect_superior_2.Y + 15;
            gfx.DrawString("CONSIGNEE:", letraDireccion, Brush, new Point(x, y));
            gfx.DrawString(etq.Cliente, letraDireccion, Brush, new Point(x+175, y+30));
            gfx.DrawString(etq.Direccion, letraDireccion, Brush, new Point(x + 175, y + 60));
            gfx.DrawString("FREIGHT:", letraDireccion, Brush, new Point(x, y+180));


            //Campos en Recuadro lateral izquierdo
            x = rect_lateral_izquierdo.X + 20;
            y = rect_lateral_izquierdo.Y + 10;
            gfx.DrawString("THIS PALLET", letraCampos, Brush, new Point(x, y));
            y += 18;
            gfx.DrawString("CONTAINS:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Contenido, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("ITEM No.:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.ClaveProducto, letraCampos, Brush, new Point(x+300, y));
            y += 38;
            gfx.DrawString("PRODUCT:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Producto, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("P.O. No.:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.OrdenCompra, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("MATERIAL:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Material, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("COLOR:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Color, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("CALIPER:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Calibre, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("STRAND COUNT:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Corte, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("GROSS WEIGHT", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.PesoBruto, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("TARE:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Tara, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("NET WEIGHT:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.PesoNeto, letraCampos, Brush, new Point(x + 300, y));
            y += 38;
            gfx.DrawString("DATE:", letraCampos, Brush, new Point(x, y));
            gfx.DrawString(etq.Fecha, letraCampos, Brush, new Point(x + 300, y));

            //Campos laterales derecho superior
            x = rect_lateral_derecho_sup.X + 51;
            y = rect_lateral_derecho_sup.Y + 25;
            gfx.DrawString("GROSS WEIGHT:", letraCampos, Brush, new Point(x, y));

            //Campos laterales derecho inferior
            x = rect_lateral_derecho_inf.X + 63;
            y = rect_lateral_derecho_inf.Y + 25;
            gfx.DrawString("NET WEIGHT:", letraCampos, Brush, new Point(x, y));

            //Campos rectangulo inferior
            x = rect_inferior.X + 260;
            y = rect_inferior.Y + 5;
            gfx.DrawString("PALLET:", letraCampos, Brush, new Point(x, y));
            x = rect_inferior.X + 430;
            gfx.DrawString("OF", letraCampos, Brush, new Point(x, y));


            return e;
        }

    }
}
