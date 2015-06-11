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
            Etiqueta etq;
            Etiqueta_Tipo_1 etq_1;
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
            switch (tipo_etiqueta_en_uso)
            { 
                case 1:     
                    e = etq_1.dibujarEtiqueta(e, etq_1);
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
                crearNuevaEtiqueta_tipoX(etq);
                label1.Text = etq_1.ClaveProducto;
                button2_Click(sender, e);
            }
        }
        private bool obtenerTipo()
        {
            if (etq.obtenerTipo(etq.ClaveProducto))
                return true;
            else
                return false;
        }

        private void crearNuevaEtiqueta_tipoX(Etiqueta etq)
        {
            switch (etq.TipoEtiqueta)
            { 
                case 1:
                    tipo_etiqueta_en_uso = 1;
                    etq_1 = new Etiqueta_Tipo_1(etq.ClaveProducto);
                    break;
                case 2:
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
            switch (tipo_etiqueta_en_uso)
            { 
                case 1:
                    printDoc.PrinterSettings.DefaultPageSettings.PaperSize.RawKind = 1;
                    break;
                case 2:
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
