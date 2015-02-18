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


        private void llamaLog(int handler, String sqlQuery, String msg)
        {
            if (miClase.EjecutaEscalar(sqlQuery) > 0)
            {
                generaReporte(handler);
                enviaMail(handler);
            }
            else
            {
                guardaLog(msg);
            }

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
                        llamaLog(handler, sqlQuery, "Anticipos Pendientes de Pago -- No existe Información");
                    } break;

                case 2:
                    {
                        // Alertas Produccion
                        miClase.EjecutaSql("exec sp_generaNotificaciones 1", false);
                        llamaLog(handler, "select COUNT(*) from notificaProduccionTEMP", "Alertas de Producción -- No existe Información");
                    } break;

                case 3:
                    {
                        //Alertas Entrega.
                        miClase.EjecutaSql("exec sp_generaNotificaciones 2", false);
                        llamaLog(handler, "select COUNT(*) from notificaProduccionTEMP", "Alertas de Entrega -- No existe Información");
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
                        llamaLog(handler, sqlQuery, "Facturas de proveedores PE con saldos pendientes de pago -- No existe Información");
                    } break;
                case 5:
                    {
                        // Reporte de Cruce de Anticipos
                        sqlQuery = @"exec sp_reporteCruceAnticipos";
                        miClase.EjecutaSql(sqlQuery, false);
                        sqlQuery = @"select count(*) from tmpReporteCruceAnticipos where ocultar=0";
                        llamaLog(handler, sqlQuery, "Reporte de Cruces de Anticipos -- No existe Información");
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
                        llamaLog(handler, sqlQuery, "Reporte de Articulos Creados -- No existe Información");
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
                        llamaLog(handler, sqlQuery, "Reporte de Vencimiento de Productos -- No existe Información");
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
                        llamaLog(handler, sqlQuery, "Importaciones Liquidadas -- No existe Información");
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
                        llamaLog(handler, sqlQuery, "IBG(s) Creados -- No existe Información");
                    } break;
                case 10:
                    {
                        // 10= Solicitud envio estados de cuenta proveedores
                        if (finMes())
                        {
                            enviaMail(handler);
                        }
                    } break;
                case 11: 
                    {
                        // 11= Reporte de importaciones pendientes de transporte.
                        sqlQuery = @"
                            SELECT		COUNT(Compra.idCompra)
                            FROM        Compra
                            WHERE		(idTipoFactura = 2) AND (idComprobante <> 33) AND (Borrar = 0) AND (Usuario <> 'OrdenLotes') 
			                            AND (FechaEntrega >= FechaRevision) AND (Numero NOT IN
                                                      (SELECT     Numero
                                                        FROM          Compra AS Compra_1
                                                        WHERE      (idTipoFactura = 14)))
                        ";
                        llamaLog(handler, sqlQuery, "Importaciones en espera de transporte -- No existe Información");
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

        private bool finMes()
        {
            // Retorna true si la fecha Actual es el último día del mes (entre lunes y viernes) o a su vez si se trata del último viernes del mes.
            if (DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)-2<=DateTime.Now.Day)
            {
                if (DateTime.Now.DayOfWeek == DayOfWeek.Friday)
                {
                    return true;
                }
                else
                {
                    if (DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month) == DateTime.Now.Day)
                        return true;
                    else
                        return false;
                }
            }
            else
            {
                return false;
            }
        }
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
             * 9= IBG´s creados en las últimas 24 horas.
             * 10= Solicitud envio estados de cuenta proveedores
             * 11= Reporte de importaciones pendientes de transporte.
             * 
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
                case 10:
                    {
                        // Esta opción no contempla un RPT para generar info.
                        nombreReporte = "";
                        nombreArchivo = "";
                    } break;
                case 11:
                    {
                        // Ordenes de compra que se encuentran en espera de transporte.
                        nombreReporte = "OCEsperaTransporte.rpt";
                        nombreArchivo = path + @"OCEsperaTransporte-" + DateTime.Now.Year.ToString() + mes + dia + ".pdf";
                    } break;
            }
            if (nroReporte!=10) // Aquí deben ir especificados las alertas que no tengan asociado un RPT
            {
                // Reportes que hacen referencia a un rpt.
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
        }

        void enviaMail(int tipo) 
        {
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("jefeimportaciones@graphicsource.com.ec"); // de (string)
            msg.Priority = MailPriority.High; // Prioridad (propiedad de MailPriority)
            msg.IsBodyHtml = false; // true si es html, false si es txt
            /*
            De acuerdo al tipo se debe definir el conjunto de destinatarios.
            Para el Release (20140410) se trabaja con un solo conjunto.
            */

            if (miClase.EjecutaEscalar("select count(*) from AlertaMails where Borrar=0 and Alerta"+tipo.ToString()+"=1") > 0)
                msg.To.Add(miClase.resultConcatenado("select mail from AlertaMails where Borrar=0 and Alerta"+tipo.ToString()+"=1 "));
            if (miClase.EjecutaEscalar("select count(*) from AlertaMails where Borrar=0 and Alerta" + tipo.ToString() + "=2") > 0)
                msg.CC.Add(miClase.resultConcatenado("select mail from AlertaMails where Borrar=0 and Alerta" + tipo.ToString() + "=2 "));
            if (miClase.EjecutaEscalar("select count(*) from AlertaMails where Borrar=0 and Alerta" + tipo.ToString() + "=3") > 0)
                msg.Bcc.Add(miClase.resultConcatenado("select mail from AlertaMails where Borrar=0 and Alerta" + tipo.ToString() + "=3 "));

            //msg.To.Add("c_jaramillo@graphicsource.com.ec"); 

            switch (tipo){
                case 1: 
                    {
                        msg.Subject = "Anticipos Pendientes de Pago"; 
                        msg.Body = "Adjunto se remite un listado en detalle de anticipos pendientes de pago para su revisión.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;

                case 2:
                    {
                        msg.Subject = "Alertas Produccion"; 
                        msg.Body = "Adjunto se remite un listado en detalle de las Ordenes de Compra que cumplirán su tiempo de producción.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;

                case 3:
                    {
                        msg.Subject = "Alertas de Entrega"; 
                        msg.Body = "Adjunto se remite un listado en detalle de los pedidos realizados a Proveedores que cumplirán su tiempo de tránsito.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;
                case 4:
                    {
                        msg.Subject = "Facturas Normales proveedores PE - Pendientes de Pago "; 
                        msg.Body = "Adjunto se remite un listado en detalle de las facturas normales de proveedores del exterior por concepto de importaciones que tienen saldos pendientes de pago.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;
                case 5:
                    {
                        msg.Subject = "Reporte de Cruces de Anticipos "; 
                        msg.Body = "Adjunto se remite un listado en detalle de los anticipos solicitados, así como también el número de las facturas finales e IG a las cuales se deberán cruzar.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;
                case 6: 
                    {
                        msg.Subject = "Reporte de Artículos creados "; 
                        msg.Body = "Adjunto se remite un listado en detalle de los artículos creados durante la última semana.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;
                case 7: 
                    {
                        msg.Subject = "Reporte de Vencimiento de Productos "; 
                        msg.Body = "Adjunto se remite un listado en detalle de artículos con su fecha de vencimiento.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                        
                    } break;
                case 8:
                    {
                        msg.Subject = "Importaciones Liquidadas "; 
                        msg.Body = "Adjunto se remite un listado en detalle de las importaciones que han sido liquidadas en las últimas 24 horas laborables.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;
                case 9:
                    {
                        msg.Subject = "IBG - IMPORTACIONES RECIBIDAS "; 
                        msg.Body = "Adjunto se remite un listado en detalle de los ingresos de bodega creados en las últimas 24 horas laborables.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;
                case 10:
                    {
                        msg.IsBodyHtml = true;
                        msg.Subject = "Solicitud Estados de Cuenta"; 
                        String mensaje = @"
                            <html>
                            <meta charset=""UTF-8"">
                             <body style=""background-color: FEFBFB"">
                                <center>
                                <img src=""http://firmasgs.hostinazo.com/firmasGs/logoGSNewGIF.gif"" align=""middle"" alt=""logo_graphicsource""/>    
                                </center>
                                <h3>
                                Estimados proveedores:
                                </h3>
                                Nuestro Departamento de Contabilidad  necesita mensualmente recibir estados de cuenta a la fecha, requisito que será indispensable para proceder con las cancelaciones de las obligaciones pendientes conciliadas con dichos reportes.
                            Favor enviar el requerimiento al correo: m_burbano@graphicsource.com.ec 
                                <h3>
                                Dear suppliers:
                                </h3>
                                Our Accounting Department needs to receive monthly our current billing statements , as a requirement that will be essential to proceed with the cancellation of outstanding obligations.
                            Please send this request to the email : m_burbano@graphicsource.com.ec 
                            <br> <br>
                            </body>
                            </html>
                            ";
                        msg.Body = mensaje;
                    } break;
                case 11:
                    {
                        msg.Subject = "Importaciones en espera de transporte";
                        msg.Body = "Adjunto se remite un listado en detalle de las importaciones que se encuentran en espera de transporte.\r\n \r\nEste mensaje ha sido generado automáticamente por el sistema."; 
                    } break;

            }
            if (tipo!=10) // Solo el 10 no tiene adjunto.
            {
                msg.Attachments.Add(new Attachment(nombreArchivo));
            }
            SmtpClient clienteSMTP = new SmtpClient("192.168.1.1"); // El servidor de correo
            
            try
            {
                if ((msg.To.Count + msg.CC.Count + msg.Bcc.Count)>0)
                {
                    clienteSMTP.Send(msg);
                    if (tipo == 10) // El reporte 10 es un caso especial porque tiene que ser copia oculta entre proveedores y debo llevar un control a quién nomás envié en el LOG.
                        guardaLog(msg.Subject + " -- Enviado a los siguientes destinatarios:\r\nPara:" + msg.To.ToString() + "\r\nCC: " + msg.CC.ToString() + "\r\nBCC: " + msg.Bcc.ToString());
                    else
                        guardaLog(msg.Subject + " -- Enviado");
                }
                else
                {
                    guardaLog(msg.Subject + " -- No hay destinatarios definidos.");
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void Principal_Load(object sender, EventArgs e)
        {
            System.IO.Directory.CreateDirectory(path);

            
            Datos.strServidor = @"192.168.1.15";
            Datos.strBase = "GraphicSource2007";
            Datos.strReporte = @"\\Servidor\Latinium\Reportes\";
            
            /*
            Datos.strServidor = @"192.168.1.34";
            Datos.strBase = "GraphicSource2007";
            Datos.strReporte = @"\\CESAR\Latinium\Reportes\";
            */

            Datos.strMaquina = miClase.EjecutaEscalarStr("select host_name()");
            Datos.idSucursal = miClase.EjecutaEscalar("Select Top 1 IdSucursal from SucursalGs Where Principal=1");

            /*
             * Desde aquí administro los reportes que el usuario me pidió y los mando a llamar. En caso de que me pida un reporte que no existe lo ignoro.
             * Con esto hago escalable la app en caso de que a futuro me pida más reportes con diferentes días.
             * */
            List<int> list; 
            foreach (int nroRep in args) 
            {
                // Reportes habilitados para envíos de Lunes a Viernes : 1,2,3,4,8,9,10,11
                list= new List<int> { 1, 2, 3, 4, 8, 9, 10, 11};
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
