using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;
namespace Etiquetas_Almacen
{
    class Etiqueta
    {
        Uri path = new Uri("file://C:/hv/EspPrueba.xls");
        string hoja = "Hoja de Especificaciones grales";
        OleDbConnection conn = new OleDbConnection();
        public Etiqueta(string cp)
        {
            this.ClaveProducto = cp;
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data" + @" Source=" + path.LocalPath + "; Extended Properties='Excel 12.0 Macro;HDR=NO;'";
        }

        public string ClaveProducto
        { get; set; }

        public string Cliente
        { get; set; }

        public string Calibre
        { get; set; }

        public string Corte
        { get; set; }

        public string Perfil
        { get; set; }

        public string Color
        { get; set; }

        public int Largo
        { get; set; }

        public int Ancho
        { get; set; }

        public string Material
        { get; set; }

        public bool EsCaja
        { get; set; }

        public bool EsPallet
        { get; set; }

        public int Cantidad
        { get; set;  }

        public int TipoEtiqueta
        { get; set; }

        public int VariacionEtiqueta
        { get; set; }

        public bool obtenerTipo(string cp)
        {
            OleDbCommand get = new OleDbCommand();
            get.Connection = conn;
            get.CommandText = "SELECT F25 FROM [" + hoja + "$] WHERE F1='" + cp + "'";
            try
            {
                conn.Open();
                OleDbDataReader reader = get.ExecuteReader();
                while (reader.Read())
                {
                    this.TipoEtiqueta = int.Parse(reader["F25"].ToString());
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
        public bool obtenerVariacionEtiqueta(string cp)
        {
            OleDbCommand get = new OleDbCommand();
            get.Connection = conn;
            get.CommandText = "SELECT F26 FROM [" + hoja + "$] WHERE F1='" + cp + "'";
            try
            {
                conn.Open();
                OleDbDataReader reader = get.ExecuteReader();
                while (reader.Read())
                {
                    this.VariacionEtiqueta = int.Parse(reader["F26"].ToString());
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
    }
}
