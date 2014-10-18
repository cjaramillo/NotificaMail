using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;              //--->Para Usar Objetos DataSet Y DataView
using System.Windows.Forms;     //--->Para usar el ojeto ComboBox
using System.Data.OleDb;
using System.Data.SqlClient;



namespace NotificaMail
{
    public class Datos
    {
        #region "Variables Datos"
        public OleDbConnection cadcon = new OleDbConnection();
        static public OleDbDataReader odatareader;
        static public OleDbCommand ocomando;
        public OleDbCommandBuilder oComandoB;
        public OleDbDataAdapter oadapter;
        public DataSet odataset;
        public DataSet odatasetAux;
        public DataView ogdataview;
        public Boolean CreaTabla = false;
        public double TotalMes = 0;
        #endregion

        #region "Variables"
        static public string strBase;
        static public string strServidor;
        static public string strMaquina;
        static public bool retencion;
        static public bool impuesto;
        static public bool transforma;
        static public bool CopiaCosto;
        static public bool InvRegular;
        static public bool Pago;
        static public string NumeroOrden;
        static public int idOrden;
        static public int idArte;
        static public int idDisenador;
        static public int idCliente;
        static public int idEquipo;
        static public Double Largo;
        static public Double Ancho;
        static public int NroRetazo;
        static public int idSucursal;
        static public int idPedido;
        static public int idCaso;
        static public bool PuestaPunto;
        static public bool Acceso;
        static public bool FiltroPedido;
        static public bool SelecEquipo;
        static public bool CambiaEmpresa;
        static public bool mSuministros;
        static public bool mProduccion;
        static public bool mTecnico;
        static public bool mImportaciones; // Añadido
        static public string Usuario;
        static public bool MuestraCaso;
        static public bool CreaEgreso;
        static public int idTecnico;
        static public bool ImprimeIbg;
        static public bool CopiaDetCaso;
        static public bool NotaDebito;
        static public bool CreaCompra;
        static public bool QuitaProy;
        static public bool RevisaProy;
        static public bool AsiSolicitud;
        static public int idTipoFactura;
        static public int idContrato;

        #endregion


        static public string strReporte;

        public void Conectar()
        {
            cadcon.ConnectionString = @"Auto Translate=True;User ID=fer;"
                + "Tag with column collation when possible=False;"
                + "Data Source = " + Datos.strServidor + "; Password=05043001;"
                + "Initial Catalog=" + Datos.strBase + ";Use Procedure for Prepare=1;"
                + "Provider=SQLOLEDB.1;Persist Security Info=True;"//Workstation ID=Local;"
                + "Use Encryption for Data=False;Packet Size=4096";
        }


        public string ConectarAdaptador()
        {
            Conectar();
            return cadcon.ConnectionString.ToString();
        }


        public void LlenaLista(CheckedListBox Lista, string mitabla, string member, string valor, string filtro, bool Bloquear)
        {
            Conectar();
            string cadsql;
            try
            {
                if (filtro == "*")
                {
                    cadsql = "select " + member + "," + valor + " from " + mitabla + " order by " + member;
                }
                else
                {
                    cadsql = "select " + member + "," + valor + " from " + mitabla + " where " + filtro + " order by " + member;
                }
                this.odataset = new DataSet();
                this.oadapter = new OleDbDataAdapter(cadsql, this.cadcon);
                this.oadapter.Fill(this.odataset, mitabla);
                this.ogdataview = new DataView(this.odataset.Tables[mitabla], "", member, DataViewRowState.OriginalRows);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            try
            {
                Lista.DataSource = this.ogdataview;
                Lista.DisplayMember = member;
                Lista.ValueMember = valor;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public void LlenaCombo(ComboBox combo, string mitabla, string member, string valor, string filtro, bool Bloquear)
        {
            if (Bloquear == true) combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            Conectar();
            string cadsql;
            try
            {
                if (filtro == "*")
                {
                    cadsql = "select " + member + "," + valor + " from " + mitabla + " order by " + member;
                }
                else
                {
                    cadsql = "select " + member + "," + valor + " from " + mitabla + " where " + filtro + " order by " + member;
                }
                this.odataset = new DataSet();
                this.oadapter = new OleDbDataAdapter(cadsql, this.cadcon);
                this.oadapter.Fill(this.odataset, mitabla);
                this.ogdataview = new DataView(this.odataset.Tables[mitabla], "", member, DataViewRowState.OriginalRows);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            try
            {
                combo.DataSource = this.ogdataview;
                combo.DisplayMember = member;
                combo.ValueMember = valor;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //Modifico la Funcion LLena Combo para mostrar n cantidad de registros Ordenados Descendentemente
        public void LlenaCombo(ComboBox combo, string mitabla, string member, string valor, string filtro, bool Bloquear, string nRegistros)
        {
            if (Bloquear == true) combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            Conectar();
            string cadsql;
            try
            {
                if (filtro == "*")
                {
                    cadsql = "select " + member + "," + valor + " from " + mitabla + " order by " + member + " Desc";
                }
                else
                {
                    cadsql = "select  " + member + "," + valor + " from " + mitabla + " where " + filtro + " order by " + member + " Desc";
                }
                this.odataset = new DataSet();
                this.oadapter = new OleDbDataAdapter(cadsql, this.cadcon);
                this.oadapter.Fill(this.odataset, mitabla);
                this.ogdataview = new DataView(this.odataset.Tables[mitabla], "", "", DataViewRowState.OriginalRows);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            try
            {
                combo.DataSource = this.ogdataview;
                combo.DisplayMember = member;
                combo.ValueMember = valor;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public void EjecutaSql(string cadsql, Boolean bMenConf)
        {
            Conectar();
            try
            {
                ocomando = new OleDbCommand(cadsql, cadcon);
                if (cadcon.State == ConnectionState.Closed) cadcon.Open();
                ocomando.CommandTimeout = 0;
                ocomando.ExecuteNonQuery();
                if (bMenConf == true) MessageBox.Show("Transaccion Finalizada", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                //El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
                
                cadcon.Close();
            }
        }

        public bool EjecutaSqlBool(string cadsql)
        {
            Conectar();
            try
            {
                ocomando = new OleDbCommand(cadsql, cadcon);
                if (cadcon.State == ConnectionState.Closed) cadcon.Open();
                ocomando.ExecuteNonQuery();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                //El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
                cadcon.Close();
            }
        }


        public double EjecutaEscalarF(string cadsql)
        {
            Conectar();
            double descalar = 0;
            try
            {
                ocomando = new OleDbCommand(cadsql, cadcon);
                if (cadcon.State == ConnectionState.Closed) cadcon.Open();
                descalar = ((double)ocomando.ExecuteScalar());
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                //El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
                cadcon.Close();
            }
            return descalar;
        }


        public DateTime EjecutaEscalarDate(string cadsql)
        {
            Conectar();
            DateTime Dateescalar = System.DateTime.Now;
            try
            {
                ocomando = new OleDbCommand(cadsql, cadcon);
                if (cadcon.State == ConnectionState.Closed) cadcon.Open();
                Dateescalar = (DateTime)ocomando.ExecuteScalar();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                //El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
                cadcon.Close();
            }
            return Dateescalar;
        }


        public string EjecutaEscalarStr(string cadsql)
        {
            Conectar();
            string strescalar = "";
            try
            {
                ocomando = new OleDbCommand(cadsql, cadcon);
                if (cadcon.State == ConnectionState.Closed) cadcon.Open();
                strescalar = (string)ocomando.ExecuteScalar();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                //El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
                cadcon.Close();
            }
            return strescalar;
        }


        public int EjecutaEscalar(string cadsql)
        {
            Conectar();
            int iescalar = 0;
            try
            {
                ocomando = new OleDbCommand(cadsql, cadcon);
                if (cadcon.State == ConnectionState.Closed) cadcon.Open();
                iescalar = ((int)ocomando.ExecuteScalar());
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                //El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
                cadcon.Close();
            }
            return iescalar;
        }


        public void LlenaGrid(DataGrid grilla, string mitabla, string comandosql)
        {
            Conectar();
            try
            {
                odataset = new DataSet();
                oadapter = new OleDbDataAdapter(comandosql, cadcon);
                oadapter.Fill(odataset, mitabla);
                //Aumento Ctrl para Verificar Si la Tabla se Creo y poder Crear el View
                //Esto ocurre Con los sp
                if (odataset.Tables.Count != 0)
                {
                    ogdataview = new DataView(odataset.Tables[mitabla], "", "", DataViewRowState.OriginalRows);
                    CreaTabla = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                grilla.DataSource = ogdataview;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        public void LlenaGridAux(DataGrid grilla, string mitabla, string comandosql)
        {
            Conectar();

            try
            {
                oadapter = new OleDbDataAdapter(comandosql, cadcon);
                oadapter.Fill(odatasetAux, mitabla);
                //Aumento Ctrl para Verificar Si la Tabla se Creo y poder Crear el View
                //Esto ocurre Con los sp
                if (odatasetAux.Tables.Count != 0)
                {
                    ogdataview = new DataView(odatasetAux.Tables[mitabla], "", "", DataViewRowState.OriginalRows);
                    CreaTabla = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                grilla.DataSource = ogdataview;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        public void HabilitaControl(Control cn, bool activar)
        {
            string stControl = cn.GetType().ToString();
            if (stControl == "System.Windows.Forms.TextBox")
                cn.Enabled = activar;
            else if (stControl == "System.Windows.Forms.Label")
                cn.Enabled = true;
            else if (stControl == "System.Windows.Forms.ListBox")
                cn.Enabled = activar;
            else if (stControl == "System.Windows.Forms.ComboBox")
                cn.Enabled = activar;
            else if (stControl == "System.Windows.Forms.DateTimePicker")
                cn.Enabled = activar;
            else if (stControl == "System.Windows.Forms.NumericUpDown")
                cn.Enabled = activar;
            else if (stControl == "System.Windows.Forms.DataGrid")
                cn.Enabled = activar;
            //else if ((stControl=="System.Windows.Forms.GroupBox")&& (cn.Name!="grpToolBar"))
            //	cn.Enabled = activar;
        }


        public void HabilitaControles(Form MiForma, bool activar)
        {
            // Habilita o deshabilida controles en forma principal

            foreach (Control cn in MiForma.Controls)
                this.HabilitaControl(cn, activar);
        }


        public void LimpiaControles(Form MiForma)
        {

            foreach (Control cn in MiForma.Controls)
            {
                string stControl = cn.GetType().ToString();
                if (stControl == "System.Windows.Forms.TextBox")
                    cn.Text = "";
            }
        }


        public bool IsNumeric(object Expression)
        {
            bool isNum;
            double retNum;
            isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }


        public void CargaExcel(DataGrid MiGrilla)
        {

            #region "Variables"
            odataset = new DataSet();
            string strArchivo = "";
            OleDbDataAdapter myData;
            OpenFileDialog ofExcel = new OpenFileDialog();
            OleDbConnection cadcon = new OleDbConnection();
            #endregion

            ofExcel.DefaultExt = "xls";
            ofExcel.Filter = "Excel (*.xls)|*.xls";

            if (DialogResult.OK == ofExcel.ShowDialog())
            {
                strArchivo = ofExcel.FileName.ToString();
            }
            else
            {
                MessageBox.Show("Cancelado Por el usuario", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //			
            //---> Conexion
            //
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source=" + strArchivo + ";" +
                "Extended Properties='Excel 8.0;'";
            //Declaro Adaptador...
            myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
            myData.TableMappings.Add("Tabla", "ExcelSube");
            odataset.Clear();

            try
            {

                myData.Fill(odataset);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del archivo " + strArchivo.ToString());
                return;
            }
            MiGrilla.TableStyles.Clear();
            MiGrilla.DataSource = odataset.Tables[0].DefaultView;

        }


        public void MiControl(GroupBox Grupo, Boolean es)
        {
            foreach (Control cn in Grupo.Controls)
            {
                if (cn.GetType().ToString() == "System.Windows.Forms.Button")
                    cn.Enabled = es;
                else if (cn.GetType().ToString() == "System.Windows.Forms.TextBox")
                    cn.Enabled = es;
                else if (cn.GetType().ToString() == "System.Windows.Forms.DateTimePicker")
                    cn.Enabled = es;
            }

        }


        public void MuestraLista(CheckedListBox miLista)
        {
            //int Ancho=miLista.Width;
            if (miLista.Size.Height != 180)
                miLista.Size = new System.Drawing.Size(miLista.Width, 180);
            else
                miLista.Size = new System.Drawing.Size(miLista.Width, 20);

        }


        public string FiltroLista(CheckedListBox miLista)
        {
            string Filtro = "";

            for (int i = 0; i < miLista.Items.Count - 1; i++)
            {

                if (miLista.GetItemChecked(i) == true) Filtro += " And '" + miLista.GetItemText(miLista.Items[i]) + "'";
            }

            return Filtro;
        }

    }
}
