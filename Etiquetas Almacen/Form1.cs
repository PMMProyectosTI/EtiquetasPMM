using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing.Drawing2D;

namespace Etiquetas_Almacen
{
    public partial class Form1 : Form
    {
        //Declaración de variables globales
        #region variablesGlobales
            int tipo_etiqueta_en_uso=0;
            int variacion_en_uso = 0;
            Etiqueta etq;
            Etiqueta_Tipo_1 etq_1;
            Etiqueta_Tipo_1_2 etq_1_2; //etiqueta tipo 1 - variacion 2
            Etiqueta_Tipo_1_3 etq_1_3; //etiqueta tipo 1 - variacion 3

            Etiqueta_Tipo_2 etq_2;
            PrintDocument printDoc = new PrintDocument();
            PrintPreviewDialog previewdlg = new PrintPreviewDialog();
        #endregion
        

        public Form1()
        {
            InitializeComponent();
            printDoc.PrintPage += new PrintPageEventHandler(printDoc_PrintPage);
        }

        void printDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            switch (etq.TipoEtiqueta)
            {
                case 1:
                    switch (etq.VariacionEtiqueta)
                    { 
                        case 1:
                            e = etq_1.dibujarEtiqueta(e, etq_1);
                            break;
                        case 2:
                            e = etq_1_2.dibujarEtiqueta(e, etq_1_2);
                            break;
                        case 3:
                            e = etq_1_3.dibujarEtiqueta(e, etq_1_3);
                            break;
                    } 
                    break;
                case 2:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_2.dibujarEtiqueta(e, etq_2);
                            break;
                    }
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            etq = new Etiqueta(textBox1.Text);
            if (obtenerTipo())
            {
                if (obtenerVariacion())
                {
                    crearNuevaEtiqueta_tipoX(etq);
                    label1.Text = etq.ClaveProducto;
                    button2_Click(sender, e);
                }
            }
        }
        private bool obtenerTipo()
        {
            if (etq.obtenerTipo(etq.ClaveProducto))
                return true;
            else
                return false;
        }
        private bool obtenerVariacion()
        {
            if (etq.obtenerVariacionEtiqueta(etq.ClaveProducto))
                return true;
            else
                return false;
        }
        private void crearNuevaEtiqueta_tipoX(Etiqueta etq)
        {
            switch (etq.TipoEtiqueta)
            { 
                case 1:
                    switch (etq.VariacionEtiqueta)
                    { 
                        case 1:
                            tipo_etiqueta_en_uso = 1;
                            variacion_en_uso = 1;
                            etq_1 = new Etiqueta_Tipo_1(etq.ClaveProducto);
                            break;
                        case 2:
                            tipo_etiqueta_en_uso = 1;
                            variacion_en_uso = 2;
                            etq_1_2 = new Etiqueta_Tipo_1_2(etq.ClaveProducto);
                            break;
                        case 3:
                            tipo_etiqueta_en_uso = 1;
                            variacion_en_uso = 3;
                            etq_1_3 = new Etiqueta_Tipo_1_3(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 2:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 2;
                            variacion_en_uso = 1;
                            etq_2 = new Etiqueta_Tipo_2(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 3:
                    break;
                case 4:
                    break;
                case 5:
                    break;
                case 6:
                    break;
                case 7:
                    break;
                case 8:
                    break;
                case 9:
                    break;
                case 10:
                    break;
                case 11:
                    break;
                case 12:
                    break;
                case 13:
                    break;
                case 14:
                    break;
                case 15:
                    break;
                case 16:
                    break;
                case 17:
                    break;
                case 18:
                    break;
                case 19:
                    break;
                case 20:
                    break;
                case 21:
                    break;
                case 22:
                    break;

            }
        }

        private void imprimir()
        {
            previewdlg.Document = printDoc;
            previewdlg.Size = new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            previewdlg.PrintPreviewControl.Zoom = 2;
            //TAMAÑO
            switch (etq.TipoEtiqueta)
            { 
                case 1:
                    printDoc.PrinterSettings.DefaultPageSettings.PaperSize.RawKind = 1;
                    break;
                case 2:
                    printDoc.PrinterSettings.DefaultPageSettings.PaperSize.RawKind = 1;
                    break;
                case 3:
                    break;
                case 4:
                    break;
                case 5:
                    break;
            }
            if (previewdlg.ShowDialog() == DialogResult.OK)
            {
                printDoc.Print();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            imprimir();
        }
    }
}
