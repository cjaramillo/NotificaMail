using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Net.Mail;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows;
using CrystalDecisions.Shared;


namespace NotificaMail
{
    public partial class Principal : Form
    {

        #region "Variables"

        string path = @"C:\Latinium\Alertas\";
        string log = "Log.txt";
        string mes = "", dia = "", hora="", minuto="", segundo="";
        Datos miClase = new Datos();

        #endregion

        void formateaDiasyMes(){
            if (Int32.Parse(DateTime.Now.Month.ToString()) < 10)
                mes = "0" + DateTime.Now.Month.ToString();
            else
                mes = DateTime.Now.Month.ToString();

            if (Int32.Parse(DateTime.Now.Day.ToString()) < 10)
                dia = "0" + DateTime.Now.Day.ToString();
            else
                dia = DateTime.Now.Day.ToString();

            if (Int32.Parse(DateTime.Now.Hour.ToString()) < 10)
                hora = "0" + DateTime.Now.Hour.ToString();
            else
                hora = DateTime.Now.Hour.ToString();

            if (Int32.Parse(DateTime.Now.Minute.ToString()) < 10)
                minuto = "0" + DateTime.Now.Minute.ToString();
            else
                minuto = DateTime.Now.Minute.ToString();

            if (Int32.Parse(DateTime.Now.Second.ToString()) < 10)
                segundo = "0" + DateTime.Now.Second.ToString();
            else
                segundo = DateTime.Now.Second.ToString();
        }
        
        public Principal()
        {
            InitializeComponent();
        }

        // Cada elemento de este array tiene el número del reporte que se necesita enviar.
        private List<int> args;
        public Principal(List<int> args) 
        {
            this.args = args;
            InitializeComponent();
        }
        

        public void enviarAlertas(int handler)
        {
            string sqlQuery;
            switch (handler)
            {
                case -1:
                    {
                        /*
                         * Ejecuto el script desde el modelo para actualizar los anticipos solicitados desde OC.
                         * PENDIENTE!
                         * Aquí hay una observación. Desde el SP emito un raiserror en caso de error.. esto no debería ser así porque 
                         * el modelo no debería interactuar directamente con la vista.. esto será solucionado cuando se integre el barco y las alertas mail
                         * en una sola aplicación..
                         * 
                         * */
                        sqlQuery =@"exec sp_ActualizaAnticiposOCS";
                        miClase.EjecutaSql(sqlQuery, false);
                    } break;
                case 1:
                    {
                        // Lista de anticipos solicitados desde OC PENDIENTES DE PAGO
                        sqlQuery = @"
                            select COUNT(*) from Compra where idTipoFactura=4 and compra.Numero in 
                                (select detcompra.RefNumero from DetCompra where detcompra.idCompra in 
                                    (select idcompra from Compra where idTipoFactura=26 and usuario='SPLotes') 
                                and DetCompra.RefNumero is not null) 
                            and compra.idCompra not in 
                                (select idCompra from Pago where idCompra in 
                                    (select idcompra from Compra where idTipoFactura=4 and compra.borrar=0 and compra.Numero in 
                                        (select detcompra.RefNumero from DetCompra where detcompra.idCompra in 
                                            (select idcompra from Compra where idTipoFactura=26 and usuario='SPLotes') 
                                        and DetCompra.RefNumero is not null and detcompra.Borrar=0)
                                    ) 
                                and pago.borrar=0)";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Anticipos Pendientes de Pago -- No existe Información");
                        }
                    } break;

                case 2:
                    {
                        // Alertas Produccion
                        miClase.EjecutaSql("exec sp_generaNotificaciones 1", false);
                        if (miClase.EjecutaEscalar("select COUNT(*) from notificaProduccionTEMP") > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Alertas de Producción -- No existe Información");
                        }
                    } break;

                case 3:
                    {
                        //Alertas Entrega.
                        miClase.EjecutaSql("exec sp_generaNotificaciones 2", false);
                        if (miClase.EjecutaEscalar("select COUNT(*) from notificaProduccionTEMP") > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Alertas de Entrega -- No existe Información");
                        }
                    } break;
                case 4:
                    {
                        // Pagos pendientes de facturas normales de proveedores PE con cargo a importaciones pendientes de pago.
                        sqlQuery = @"
                            select COUNT(*)
                            from (
	                            select compra.idCompra
	                            from Compra 
		                            inner join Cliente on compra.idCliente=cliente.idCliente
		                            inner join DetCompra on compra.idCompra=detcompra.idCompra
		                            inner join Articulo on detcompra.idArticulo=Articulo.idArticulo
		                            left outer join Pago on compra.idCompra=pago.idCompra
	                            where compra.idTipoFactura=4 and compra.Borrar=0 and compra.Fecha>='20140101'
		                            and compra.idCliente in (select idCliente from Cliente where Nombre like 'PE %') -- Todos los idCliente que son PE
		                            and detcompra.idArticulo in (select idArticulo from Articulo where Articulo like 'IG-%') -- Todos los artículos asociados a la IG.
	                            group by compra.idCompra, compra.numero, compra.Saldo, compra.Total,compra.Fecha,compra.FechaVencimiento,cliente.Nombre, 
		                            Compra.FechaIngreso,Articulo.Articulo, Cliente.DiasCredito,Compra.DiasCredito
	                            having isnull(SUM(pago.pago),0)<compra.Saldo
                            ) tabla
                        ";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Facturas de proveedores PE con saldos pendientes de pago -- No existe Información");
                        }
                    } break;
                case 5:
                    {
                        // Reporte de Cruce de Anticipos
                        sqlQuery = @"exec sp_reporteCruceAnticipos";
                        miClase.EjecutaSql(sqlQuery, false);
                        sqlQuery = @"select count(*) from tmpReporteCruceAnticipos where ocultar=0";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Reporte de Cruces de Anticipos -- No existe Información");
                        }
                    }break;
                case 6: { 
                        // Reporte de artículos creados durante la última semana.
                        sqlQuery = @"
                            SELECT  COUNT(*)
                            FROM    Articulo LEFT OUTER JOIN
			                            ArticuloSubGrupo ON Articulo.idSubGrupo = ArticuloSubGrupo.idSubGrupo LEFT OUTER JOIN
                                        ArticuloMarca ON Articulo.idMarca = ArticuloMarca.idMarca
                            WHERE   (Articulo.FechaIngreso >= DATEADD(week, - 1, GETDATE()))
                        ";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0) 
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Reporte de Articulos Creados -- No existe Información");
                        }
                    } break;
                case 7: 
                    { 
                        // Reporte de Vencimiento de Productos
                        sqlQuery = @"
                            SELECT	COUNT (*)
                            FROM    ArticuloGrupo INNER JOIN
			                            Articulo INNER JOIN
                                        Compra INNER JOIN
                                        DetCompra ON Compra.idCompra = DetCompra.idCompra ON Articulo.idArticulo = DetCompra.idArticulo ON ArticuloGrupo.idGrupoArticulo = Articulo.idGrupoArticulo
                            WHERE	(Articulo.idGrupoArticulo = 41) AND (Compra.idTipoFactura = 9) AND (Compra.idSubProyecto = 1) 
		                            AND (Compra.NUMERO LIKE 'IBG-%') 
		                            AND (GETDATE() >= DATEADD(month, - 6, DetCompra.Vencimiento)) 
		                            and (DATEADD(MONTH,-3, GETDATE())<detcompra.Vencimiento)
		                            and detcompra.RefCodigo is not null and LEN(LTRIM(detcompra.refcodigo))>0
                        ";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Reporte de Vencimiento de Productos -- No existe Información");
                        }
                    } break;
                case 8: 
                    {
                        /*
                         * Importaciones Liquidadas. (Artículos de tipo IG desactivados.)
                         * Si el reporte se ejecuta un Lunes le resta 3 días, si se ejecuta un domingo resta 2 días; para el resto de días siempre restará un solo día.
                         * 
                         * OJO: Si hacen alguna liquidación un sábado o un domingo por defecto saldrá el día LUNES.
                         * Si generan un reporte el Domingo jala la información desde el Viernes.
                         * Si generan un reporte el Lunes jala la información desde el Viernes
                         * */

                        sqlQuery = @"
                            declare @vNroDias int 
                            set @vNroDias=-1
                            if (datepart(dw,GETDATE())=1)
	                            set @vNroDias=-3
                            if (datepart(dw,GETDATE())=7)
	                            set @vNroDias=-2
                            select	count(ArticuloIGDesactivado.id)
                            from	ArticuloIGDesactivado right outer join (
                                -- Esta subconsulta devuelve los últimos estados registrados en la tabla ArticuloIGDesactivado según la línea de tiempo.
                                select MAX(id) as id, idArticulo
                                from ArticuloIGDesactivado
                                where Fecha>=DATEADD(day,@vNroDias,getdate())
                                group by idArticulo
                            ) filtro on ArticuloIGDesactivado.id=filtro.id
                            where ArticuloIGDesactivado.descontinuado=1
                        ";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0) 
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("Importaciones Liquidadas -- No existe Información");
                        }
                    } break;
                case 9:
                    {
                        // IBG´s creados en las últimas 24 horas
                        sqlQuery = @"
                            declare @vNroDias int 
                            set @vNroDias=-1
                            if (datepart(dw,GETDATE())=1)
	                            set @vNroDias=-3
                            if (datepart(dw,GETDATE())=7)
	                            set @vNroDias=-2
                            Select	COUNT (Compra.idCompra)
                            FROM    Compra left outer join Cliente on compra.idCliente=Cliente.idCliente
                            WHERE   (Compra.idTipoFactura = 9) AND (Compra.Numero LIKE 'IBG-%') AND (Compra.FechaIngreso >= DATEADD(day, @vNroDias, GETDATE())) 
		                            AND Cliente.Nombre like 'PE %'
                        ";
                        if (miClase.EjecutaEscalar(sqlQuery) > 0)
                        {
                            generaReporte(handler);
                            enviaMail(handler);
                        }
                        else
                        {
                            guardaLog("IBG(s) Creados -- No existe Información");
                        }
                    } break;
            }
        }

        
        void guardaLog(string mensaje) 
        {
            System.IO.Directory.CreateDirectory(path);
            formateaDiasyMes();
            if (File.Exists(String.Concat(path, log)))
            {
                using (System.IO.FileStream fs = System.IO.File.OpenWrite(String.Concat(path, log)))
                {
                    fs.Position = fs.Length;
                    mensaje = "\r\n" + DateTime.Now.Year.ToString() + mes + dia +"--"+ hora+":"+minuto+":"+segundo + " " + mensaje;
                    char[] vector = new char[2000];
                    vector = mensaje.ToCharArray();
                    for (byte i = 0; i < mensaje.Length; i++)
                    {
                        fs.WriteByte((byte)vector[i]);
                    }
                }
            }
            else 
            {
                using (System.IO.FileStream fs = System.IO.File.Create(String.Concat(path, log)))
                {
                    string mensajeInicial = "*************************** LOG DE ENVIO ALERTAS ************************";
                    char[] vector = new char[2000];
                    vector=mensajeInicial.ToCharArray();
                    for (byte i = 0; i < mensajeInicial.Length; i++)
                    {
                        fs.WriteByte((byte)vector[i]);
                    }
                }
                using (System.IO.FileStream fs = System.IO.File.OpenWrite(String.Concat(path, log)))
                {
                    fs.Position = fs.Length;
                    mensaje = "\r\n" + DateTime.Now.Year.ToString() + mes + dia + "--" + hora + ":" + minuto + ":" + segundo + " " + mensaje;
                    char[] vector = new char[2000];
                    vector = mensaje.ToCharArray();
                    for (byte i = 0; i < mensaje.Length; i++)
                    {
                        fs.WriteByte((byte)vector[i]);
                    }
                }
            }
        }

        private string nombreReporte = "", nombreArchivo = "";
        
        public void generaReporte(int nroReporte)
        {
            /* nroReporte: 
             * 1= Anticipos Solicitados desde OC que se encuentran pendientes de PAGO
             * 2= Alertas de producción
             * 3= Alertas de entrega
             * 4= Alertas de Facturas normales pendientes de pago.
             * 5= Reporte de Cruces de Anticipos.
             * 6= Reporte de artículos creados durante la última semana
             * 7= Reporte de Vencimiento de productos.
             * 8= Importaciones Liquidadas. (Artículos de tipo IG desactivados.)
             * 9= IBG´s creados en las últimas 24 horas. PI
             */
            nombreReporte = nombreArchivo = "";
            formateaDiasyMes();

            switch (nroReporte) 
            {
                case 1: 
                    {
                        nombreReporte = "AnticiposOCPendientesPago.rpt";
                        nombreArchivo = path+@"AnticiposPendientesPago-"+ DateTime.Now.Year.ToString() + mes + dia+".pdf";
                    } break;
                case 2:
                    {
                        nombreReporte = "Alertas.rpt";
                        nombreArchivo = path + @"AlertasProduccion-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 3:
                    {
                        nombreReporte = "Alertas.rpt";
                        nombreArchivo = path + @"AlertasEntrega-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 4: 
                    {
                        nombreReporte = "FacturasNormalesPendientePago.rpt";
                        nombreArchivo = path + @"FacturasNormalesPendientePago-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 5:
                    {
                        nombreReporte = "ReporteCruceAnticipos.rpt";
                        nombreArchivo = path + @"ReporteCruceAnticipos-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 6:
                    {
                        nombreReporte = "ArticulosCreados.rpt";
                        nombreArchivo = path + @"ArticulosCreados-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 7: 
                    {
                        nombreReporte = "VencimientoProducto.rpt";
                        nombreArchivo = path + @"VencimientoProducto-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 8:
                    {
                        nombreReporte = "ImportacionesLiquidadas.rpt";
                        nombreArchivo = path + @"ImportacionesLiquidadas-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
                case 9:
                    {
                        nombreReporte = "IBGsCreados.rpt";
                        nombreArchivo = path + @"IBGsCreados-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;

            }

            // Se genera el reporte.
            ReportDocument oRpt = new ReportDocument();
            string strReporte = Datos.strReporte + nombreReporte;
            if (!File.Exists(@strReporte))
            {
                MessageBox.Show("Archivo no existe: " + strReporte);
                return;
            }
            try
            {
                oRpt.Load(strReporte);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Carga de Reporte: " + strReporte);
                return;
            }

            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            //Seteo la Informacion para la cadena de conexion de los Reportes
            crConnectionInfo.ServerName = Datos.strServidor;
            crConnectionInfo.DatabaseName = Datos.strBase;
            crConnectionInfo.UserID = "fer";
            crConnectionInfo.Password = "05043001";

            //Declaro los objetos que voy a utilizar
            TableLogOnInfo crTableLogOnInfo;
            Database crDatabase = oRpt.Database;//-->Para la BDD
            Tables crTables = crDatabase.Tables;//-->Para las tablas
            Table crTable;

            //------Me barro las tablas-----
            for (int i = 0; i < crTables.Count; i++)
            {
                crTable = crTables[i];
                crTableLogOnInfo = crTable.LogOnInfo;
                crTableLogOnInfo.ConnectionInfo = crConnectionInfo; //-->Asigno La Informacion de la Conexion
                crTable.ApplyLogOnInfo(crTableLogOnInfo);
            }
            this.crvReportes.ReportSource = oRpt;
            
            // Inicia exportación a pdf.
            oRpt.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
            oRpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
            DiskFileDestinationOptions objDiskOpt = new DiskFileDestinationOptions();

            objDiskOpt.DiskFileName = nombreArchivo;
            oRpt.ExportOptions.DestinationOptions = objDiskOpt;
            oRpt.Export();
        }

        void enviaMail(int tipo) 
        { 
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("m_ortiz@graphicsource.com.ec"); // de (string)

            /*
            De acuerdo al tipo se debe definir el conjunto de destinatarios.
            Para el Release (20140410) se trabaja con un solo conjunto.
            */
            switch (tipo){
                case 1: 
                    {
                        
                        msg.To.Add("m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec, m_burbano@graphicsource.com.ec, r_ponce@graphicsource.com.ec, l_correa@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Anticipos Pendientes de Pago"; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de anticipos pendientes de pago para su revisión.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;

                case 2:
                    {
                        
                        msg.To.Add("m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Alertas Produccion"; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de las Ordenes de Compra que cumplirán su tiempo de producción.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;

                case 3:
                    {
                        
                        msg.To.Add("m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Alertas de Entrega"; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de los pedidos realizados a Proveedores que cumplirán su tiempo de tránsito.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
                case 4:
                    {
                        
                        msg.To.Add("m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec, m_burbano@graphicsource.com.ec, l_correa@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Facturas Normales proveedores PE - Pendientes de Pago "; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de las facturas normales de proveedores del exterior por concepto de importaciones que tienen saldos pendientes de pago.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
                case 5:
                    {
                        
                        msg.To.Add("m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec, m_burbano@graphicsource.com.ec, l_correa@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Reporte de Cruces de Anticipos "; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de los anticipos solicitados, así como también el número de las facturas finales e IG a las cuales se deberán cruzar.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
                case 6: 
                    {
                        
                        msg.To.Add(@"m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec, p_valenzuela@graphicsource.com.ec, r_cevallos@graphicsource.com.ec, t_rivas@graphicsource.com.ec, k_mejia@graphicsource.com.ec, a_rivas@graphicsource.com.ec, j_pacurucu@graphicsource.com.ec,l_gomez@graphicsource.com.ec, m_rodriguez@graphicsource.com.ec, s_marcatoma@graphicsource.com.ec, silvim2006@hotmail.com,  c_bravo@graphicsource.com.ec, logistica@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Reporte de Artículos creados "; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de los artículos creados durante la última semana.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
                case 7: 
                    {
                        
                        msg.To.Add(@"m_ortiz@graphicsource.com.ec, importaciones@graphicsource.com.ec, r_cevallos@graphicsource.com.ec, t_rivas@graphicsource.com.ec, c_dolder@graphicsource.com.ec, k_mejia@graphicsource.com.ec, a_rivas@graphicsource.com.ec, j_pacurucu@graphicsource.com.ec, m_rodriguez@graphicsource.com.ec, x_estevez@graphicsource.com.ec, logistica@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Reporte de Vencimiento de Productos "; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de artículos con su fecha de vencimiento.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
                case 8:
                    {
                        //Importaciones Liquidadas. (Artículos de tipo IG desactivados.)
                        
                        msg.To.Add(@"m_ortiz@graphicsource.com.ec, r_cevallos@graphicsource.com.ec, importaciones@graphicsource.com.ec, x_estevez@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Importaciones Liquidadas "; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de las importaciones que han sido liquidadas en las últimas 24 horas laborables.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
                case 9:
                    {
                        // IBG´s creados en las últimas 24 horas.
                        // vendedores
                        
                        msg.To.Add(@"m_ortiz@graphicsource.com.ec, r_cevallos@graphicsource.com.ec, importaciones@graphicsource.com.ec, x_estevez@graphicsource.com.ec, l_gomez@graphicsource.com.ec, a_rivas@graphicsource.com.ec, p_valenzuela@graphicsource.com.ec, t_rivas@graphicsource.com.ec, j_pacurucu@graphicsource.com.ec, j_guevara@graphicsource.com.ec, k_mejia@graphicsource.com.ec, soporte@graphicsource.com.ec, s_marcatoma@graphicsource.com.ec, silvim2006@hotmail.com, c_bravo@graphicsource.com.ec,  m_chavez@graphicsource.com.ec, j_soria@graphicsource.com.ec, e_davila@graphicsource.com.ec, g_pacheco@graphicsource.com.ec, j_coello@graphicsource.com.ec, f_vinces@graphicsource.com.ec, h_cardenas@graphicsource.com.ec"); // para (string)
                        msg.Bcc.Add("c_jaramillo@graphicsource.com.ec, fernando_defaz@graphicsource.com.ec"); // copia oculta (string separado con comas para varios)
                        
                        //msg.To.Add("c_jaramillo@graphicsource.com.ec"); // para (string)
                        msg.Subject = "Nuevos IBG(s) "; // Asunto (string)
                        msg.Body = "Adjunto se remite un listado en detalle de los ingresos de bodega creados en las últimas 24 horas laborables.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema, por favor no responda al mismo."; // Cuerpo del mensaje (string)
                    } break;
            }

            msg.Priority = MailPriority.High; // Prioridad (propiedad de MailPriority)
            msg.IsBodyHtml = false; // true si es html, false si es txt
            msg.Attachments.Add(new Attachment(nombreArchivo));
            SmtpClient clienteSMTP = new SmtpClient("192.168.1.1"); // El servidor de correo
            try
            {
                clienteSMTP.Send(msg);
                //MessageBox.Show("Mensaje Enviado Favor Revisar", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                guardaLog(msg.Subject+" -- Enviado");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void Principal_Load(object sender, EventArgs e)
        {
            System.IO.Directory.CreateDirectory(path);
            /*
            Datos.strServidor = @"192.168.1.15";
            Datos.strBase = "GraphicSource2007";
            Datos.strReporte = @"\\Servidor\Latinium\Reportes\";
            */
            
            
            Datos.strServidor = @"192.168.1.56";
            Datos.strBase = "GraphicSource2007";
            Datos.strReporte = @"\\CESAR\Latinium\Reportes\";
            

            Datos.strMaquina = miClase.EjecutaEscalarStr("select host_name()");
            Datos.idSucursal = miClase.EjecutaEscalar("Select Top 1 IdSucursal from SucursalGs Where Principal=1");

            /*
             * Desde aquí administro los reportes que el usuario me pidió y los mando a llamar. En caso de que me pida un reporte que no existe lo ignoro.
             * Con esto hago escalable la app en caso de que a futuro me pida más reportes con diferentes días.
             * */
            List<int> list; 
            foreach (int nroRep in args) 
            {
                // Reportes habilitados para envíos de Lunes a Viernes : 1,2,3,4,8,9
                list= new List<int> { 1, 2, 3, 4, 8, 9 };
                if (list.Contains(nroRep) && (int)DateTime.Now.DayOfWeek >= 1 && (int)DateTime.Now.DayOfWeek <= 5)
                {
                    enviarAlertas(nroRep);
                }
                else
                {
                    // Reportes habilitados para envíos sólo los Martes y Jueves : 5
                    list = new List<int> { 5 };
                    if (list.Contains(nroRep) && ((int)DateTime.Now.DayOfWeek == 2 || (int)DateTime.Now.DayOfWeek == 4))
                    {
                        enviarAlertas(nroRep);
                    }
                    else
                    {
                        // Reportes habilitados para envíos sólo los días Lunes: 6,7
                        list = new List<int> { 6, 7 };
                        if (list.Contains(nroRep) && (int)DateTime.Now.DayOfWeek == 1)
                        {
                            enviarAlertas(nroRep);
                        }
                        else 
                        {
                            // Script de actualización de Anticipos solicitados desde OC: Se ejecutará solo lunes, miércoles y viernes. Siempre y cuando se reciba el argumento.
                            list = new List<int> { -1 };
                            if (list.Contains(nroRep) && ((int)DateTime.Now.DayOfWeek == 1 || (int)DateTime.Now.DayOfWeek == 3 || (int)DateTime.Now.DayOfWeek == 5))
                            {
                                enviarAlertas(nroRep);
                            }
                        }
                    }
                }
            }
        }

        private void Principal_Shown(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
