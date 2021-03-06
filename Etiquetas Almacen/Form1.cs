﻿using System;
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
            Etiqueta_Tipo_2_2 etq_2_2; //La dos y la 4 son identicas por lo que no se creo la variación 4
            Etiqueta_Tipo_2_3 etq_2_3;

            Etiqueta_Tipo_3_1 etq_3_1;
            
            Etiqueta_Tipo_4_1 etq_4_1;
            
            Etiqueta_Tipo_5_1 etq_5_1;

            Etiqueta_Tipo_6_1 etq_6_1;

            Etiqueta_Tipo_8_1 etq_8_1; //Vvariacion identica

            Etiqueta_Tipo_10_1 etq_10_1;

            Etiqueta_Tipo_12_1 etq_12_1;

            Etiqueta_Tipo_14_1 etq_14_1;
            Etiqueta_Tipo_14_2 etq_14_2;

            Etiqueta_Tipo_16_1 etq_16_1;

            Etiqueta_Tipo_17_1 etq_17_1;

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
                        case 2:
                            e = etq_2_2.dibujarEtiqueta(e, etq_2_2);
                            break;
                        case 3:
                            e = etq_2_3.dibujarEtiqueta(e, etq_2_3);
                            break;
                    }
                    break;
                case 3:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_3_1.dibujarEtiqueta(e, etq_3_1);
                            break;
                    }
                    break;
                case 4:
                    switch (etq.VariacionEtiqueta)
                    { 
                        case 1:
                            e = etq_4_1.dibujarEtiqueta(e, etq_4_1);
                            break;
                    }
                    break;
                case 5:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_5_1.dibujarEtiqueta(e, etq_5_1);
                            break;
                    }
                    break;
                case 6:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_6_1.dibujarEtiqueta(e, etq_6_1);
                            break;
                    }
                    break;
                case 8:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_8_1.dibujarEtiqueta(e, etq_8_1);
                            break;
                    }
                    break;
                case 10:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_10_1.dibujarEtiqueta(e, etq_10_1);
                            break;
                    }
                    break;
                case 12:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_12_1.dibujarEtiqueta(e, etq_12_1);
                            break;
                    }
                    break;
                case 14:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_14_1.dibujarEtiqueta(e, etq_14_1);
                            break;
                        case 2:
                            e = etq_14_2.dibujarEtiqueta(e, etq_14_2);
                            break;
                    }
                    break;
                case 16:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_16_1.dibujarEtiqueta(e, etq_16_1);
                            break;
                    }
                    break;
                case 17:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            e = etq_17_1.dibujarEtiqueta(e, etq_17_1);
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
            this.Cursor = Cursors.WaitCursor;
            etq = new Etiqueta(textBox1.Text);
            if (obtenerTipo())
            {
                crearNuevaEtiqueta_tipoX(etq);
                label1.Text = etq.ClaveProducto;
                button2_Click(sender, e);
            }
            this.Cursor = Cursors.Default;
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
                        case 2:
                            tipo_etiqueta_en_uso = 2;
                            variacion_en_uso = 2;
                            etq_2_2 = new Etiqueta_Tipo_2_2(etq.ClaveProducto);
                            break;
                        case 3:
                            tipo_etiqueta_en_uso = 2;
                            variacion_en_uso = 3;
                            etq_2_3 = new Etiqueta_Tipo_2_3(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 3:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 3;
                            variacion_en_uso = 1;
                            etq_3_1 = new Etiqueta_Tipo_3_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 4:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 4;
                            variacion_en_uso = 1;
                            etq_4_1 = new Etiqueta_Tipo_4_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 5:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 5;
                            variacion_en_uso = 1;
                            etq_5_1 = new Etiqueta_Tipo_5_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 6:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 6;
                            variacion_en_uso = 1;
                            etq_6_1 = new Etiqueta_Tipo_6_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 7:
                    break;
                case 8:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 8;
                            variacion_en_uso = 1;
                            etq_8_1 = new Etiqueta_Tipo_8_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 9:
                    break;
                case 10:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 10;
                            variacion_en_uso = 1;
                            etq_10_1 = new Etiqueta_Tipo_10_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 11:
                    break;
                case 12:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 12;
                            variacion_en_uso = 1;
                            etq_12_1 = new Etiqueta_Tipo_12_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 13:
                    break;
                case 14:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 14;
                            variacion_en_uso = 1;
                            etq_14_1 = new Etiqueta_Tipo_14_1(etq.ClaveProducto);
                            break;
                        case 2:
                            tipo_etiqueta_en_uso = 14;
                            variacion_en_uso = 2;
                            etq_14_2 = new Etiqueta_Tipo_14_2(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 15:
                    break;
                case 16:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 16;
                            variacion_en_uso = 1;
                            etq_16_1 = new Etiqueta_Tipo_16_1(etq.ClaveProducto);
                            break;
                    }
                    break;
                case 17:
                    switch (etq.VariacionEtiqueta)
                    {
                        case 1:
                            tipo_etiqueta_en_uso = 17;
                            variacion_en_uso = 1;
                            etq_17_1 = new Etiqueta_Tipo_17_1(etq.ClaveProducto);
                            break;
                    }
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
                    printDoc.PrinterSettings.DefaultPageSettings.PaperSize.RawKind = 1;

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
