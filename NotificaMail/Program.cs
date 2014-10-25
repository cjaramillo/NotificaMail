using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using System.Windows.Forms;


/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
*/
namespace NotificaMail
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main(String[] args)
        {
            List<int> r;
            // Tengo que recibir parámetros para la ejecución. Si no recibo parámetros no hago nada.!
            if (args[0] == null || args[0].Length == 0)
            {
                Application.Exit();
            }
            else
            {
                r = convertir(args);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Principal(r));
            }
        }

        static List<int> convertir(String[] args)
        {
            List<int> retorno = new List<int>();
            foreach (String arg in args)
            {
                try
                {
                    retorno.Add(Int32.Parse(arg));
                }
                catch (InvalidCastException ice)
                {
                    Console.WriteLine("Se ha encontrado un argumento inválido: " + ice.Message);
                }
            }
            return retorno;
        } 
    }
}
