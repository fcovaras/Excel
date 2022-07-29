using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WinPerGenArchivo
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]

        static void Main(string [] arg)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            Application.Run(new frmGenArchivo(arg));
        }
    }
}
