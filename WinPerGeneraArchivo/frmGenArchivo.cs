using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinPerGenArchivo
{
    public partial class frmGenArchivo : Form
    {
        public string file1 = string.Empty;
        public string file2 = string.Empty;
        public int nReg = 0;
        public string registro = string.Empty;
        public int tipo = -1;
        public string archivo_sal = string.Empty;
        public string delimitador = string.Empty;
        public int nro_lineas = -1;
        public bool bSeparaArchivo = false;
        public int formato_salida = 0;

        public frmGenArchivo(string[] argv)
        {
            file1 = argv[0];
            file2 = argv[1];
            InitializeComponent();
            if (argv.Length > 2)
            {
                archivo_sal = argv[2].ToString();
            }

            label2.Text = file1.ToString();
            label3.Text = file2.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Show();
            LeeEspecificaciones(file1);

        }


        private void LeeEspecificaciones(string sArch)
        {
            int counter = 0;
            const string mtdName = "LeeEspecificaciones";

            try
            {
                using (StreamReader lector = new StreamReader(sArch, Encoding.Default))
                {
                    while (lector.Peek() > -1)
                    {
                        counter++;
                        string linea = lector.ReadLine();
                        if (counter == 1)
                            tipo = int.Parse(linea);
                        else if (counter == 2)  
                        {
                            if (String.IsNullOrEmpty(archivo_sal))
                                archivo_sal = linea;
                        }
                        else if (counter == 3)
                            delimitador = linea;
                        else if (counter == 4)
                            nro_lineas = int.Parse(linea);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, mtdName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }


            if (tipo > 0)
            {
                if (nro_lineas > 0)
                    bSeparaArchivo = true;

                if (tipo == 1)
                    CreateExcel(file2.ToString(), archivo_sal);
                else if (tipo == 3)
                    ExportarAHTML(file2.ToString(), archivo_sal);
                else if (tipo > 3)
                    ExportarATxt(file2.ToString(), archivo_sal);
                else
                    mensaje.Text = "Opción no valida";
            }


        }

        public void CreateExcel(string sArchDatos, string nameFile)
        {
            const string mtdName = "CreateExcel";
            var CodErr = 0;
            var MsgErr = "";

            string path = nameFile;
            string extension = Path.GetExtension(path);
            int index = path.LastIndexOf('\\');
            //string fileName = path.Substring(index, path.Length);
            string fileName = path.Substring(path.LastIndexOf('\\') + 1);
            index = fileName.LastIndexOf('.');
            string FileName = fileName.Substring(0, index);
            string pathGuardar = Path.GetDirectoryName(path);

            Microsoft.Office.Interop.Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

           
            if (extension.ToLower() == "xls")
                formato_salida = (int)Excel.XlFileFormat.xlWorkbookNormal;
            else
                formato_salida = (int)Excel.XlFileFormat.xlWorkbookDefault;



            try
            {
                //var apExcel = new Excel.Application();
                //object opc = Type.Missing;

                int archivo = 1;

                xlApp = new Excel.Application();

                xlApp.DecimalSeparator = ",";
                xlApp.ThousandsSeparator = ".";
                xlApp.UseSystemSeparators = false;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                int row = 0;
                Excel.Range rng = (Excel.Range)xlWorkSheet.Rows[1];
                rng.EntireRow.Font.Bold = true;

                nameFile = pathGuardar + "\\" + FileName + extension;

                foreach (string fila in System.IO.File.ReadAllLines(sArchDatos, Encoding.Default))
                {
                    row++;
                    mensaje.Refresh();
                    if (row > 0)
                    {
                        mensaje.Text = string.Format("Procesando registro: {0}", row);
                        mensaje2.Text = string.Format("Archivo Salida: {0}", nameFile);
                    }
                    if ((row > nro_lineas) && bSeparaArchivo)
                    {
                        nameFile = pathGuardar + "\\" + FileName + "-" + archivo + extension;

                        xlWorkBook.SaveAs(nameFile);
                        xlWorkBook.Close();
                        archivo++;
                        row = 1;
                        xlApp = new Excel.Application();
                        misValue = System.Reflection.Missing.Value;
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    }
                    int col = 0;
                    foreach (var campo in fila.Split(new Char[] { '|' }))
                    {
                        xlWorkSheet.Cells[row, ++col] = campo;
                    }
                }

                if (archivo > 1)
                    nameFile = pathGuardar + "\\" + FileName + "-" + archivo + extension;
                else
                    nameFile = pathGuardar + "\\" + FileName + extension;

                xlWorkBook.SaveAs(nameFile, formato_salida, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);
            }
            catch (Exception exception)
            {
                // var msg = "Error al ejecutar " + mtdName + ": " + exception.Message;

                CodErr = 1;
                MsgErr = exception.Message;
                mensaje.Text = MsgErr;
                MessageBox.Show(MsgErr, mtdName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            finally
            {
                CodErr = 0;
                // MessageBox.Show("Proceso finalizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        //public void ExportAExcel(string sArchDatos, string nameFile)
        //{
        //    const string mtdName = "ExportAExcel";
        //    var CodErr = 0;
        //    var MsgErr = "";

        //    string path = nameFile;
        //    string extension = Path.GetExtension(path);
        //    int index = path.LastIndexOf('\\');
        //    //string fileName = path.Substring(index, path.Length);
        //    string fileName = path.Substring(path.LastIndexOf('\\') + 1);
        //    index = fileName.LastIndexOf('.');
        //    string FileName = fileName.Substring(0, index);
        //    string pathGuardar = Path.GetDirectoryName(path);
            

        //    try
        //    {
        //        var apExcel = new Excel.Application();
        //        object opc = Type.Missing;

        //        int archivo = 1;

        //        var libro = apExcel.Workbooks.Add(opc);
        //        var hoja1 = new Excel.Worksheet();
        //        hoja1 = (Excel.Worksheet)libro.Sheets.Item["Hoja1"];
        //        hoja1.Activate();


        //        int row = 0;
        //        Excel.Range rng = (Excel.Range)hoja1.Rows[1];
        //        rng.EntireRow.Font.Bold = true;

        //        nameFile = pathGuardar + "\\" + FileName + extension;

        //        foreach (string fila in System.IO.File.ReadAllLines(sArchDatos, Encoding.Default))
        //        {
        //            row++;
        //            mensaje.Refresh();
        //            if (row > 0)
        //            {
        //                mensaje.Text = string.Format("Procesando registro: {0}", row);
        //                mensaje2.Text = string.Format("Archivo Salida: {0}", nameFile);
        //            }
        //            if ((row > nro_lineas) && bSeparaArchivo)
        //            {



        //                nameFile = pathGuardar + "\\" + FileName + "-" + archivo + extension;


        //                libro.SaveAs(nameFile);
        //                libro.Close();
        //                archivo++;
        //                row = 1;
        //                apExcel = new Excel.Application();
        //                opc = Type.Missing;
        //                libro = apExcel.Workbooks.Add(opc);
        //                hoja1 = new Excel.Worksheet();
        //                hoja1 = (Excel.Worksheet)libro.Sheets.Item["Hoja1"];
        //                hoja1.Activate();

        //            }
        //            int col = 0;
        //            foreach (var campo in fila.Split(new Char[] { '|' }))
        //            {
        //                hoja1.Cells[row, ++col] = campo;
        //            }
        //        }

        //        if (archivo > 1)
        //            nameFile = pathGuardar + "\\" + FileName + "-" + archivo + extension;
        //        else
        //            nameFile = pathGuardar + "\\" + FileName + extension;

        //        libro.SaveAs(nameFile);
        //        libro.Close();
        //    }
        //    catch (Exception exception)
        //    {
        //       // var msg = "Error al ejecutar " + mtdName + ": " + exception.Message;
               
        //        CodErr = 1;
        //        MsgErr = exception.Message;
        //        mensaje.Text = MsgErr;
        //        MessageBox.Show(MsgErr, mtdName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        Application.Exit();
        //    }
        //    finally
        //    {
                
        //        MessageBox.Show("Proceso finalizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        Application.Exit();
        //    }
        //}

        public void ExportarATxt(string sArchDatos, string nameFile)
        {
            const string mtdName = "ExportarATxt";
            var CodErr = 0;
            var MsgErr = "";
            try
            {
                
                System.IO.StreamWriter sw = new StreamWriter(nameFile);
                int row = 0;
                foreach (string fila in System.IO.File.ReadAllLines(sArchDatos))
                {
                    row++;
                    mensaje.Text = string.Format("Procesando registro {0}:", row);
                    registro = fila.ToString().Replace(delimitador, "\t");
                    sw.WriteLine(registro);
                    
                }
                sw.Close();
            }
            catch (Exception e)
            {

                mensaje.Text="Exception: " + e.Message;
                Application.Exit();
            }
            finally
            {
                mensaje.Text ="Finalizado.";
                Application.Exit();
            }
        }

        public void ExportarAHTML(string sArchDatos, string nameFile)
        {
            const string mtdName = "ExportarAHTML";
            var CodErr = 0;
            var MsgErr = "";
            string path = nameFile;
            string extension = Path.GetExtension(path);
            string fileName = path.Substring(path.LastIndexOf(((char)92)) + 1);
            int index = fileName.LastIndexOf('.');
            string FileName = fileName.Substring(0, index);
            string pathGuardar = Path.GetDirectoryName(path);
            try
            {
                System.IO.StreamWriter sw = new StreamWriter(nameFile);
                var sBufferEnc = "<html> <head> </head> <body> <table width=\"75%\" border=\"1\" <tr>";
                int row = 0;
                sw.WriteLine(sBufferEnc);
                mensaje2.Text = string.Format("Archivo: {0}", FileName);
                foreach (string fila in System.IO.File.ReadAllLines(sArchDatos))
                {
                    row++;
                    mensaje.Text = string.Format("Procesando registro: {0}", row);

                    if (row > 0)
                    {
                       // int col = 0;
                        foreach (var campo in fila.Split(new Char[] { '|' }))
                        {
                            var sBufferDet = campo.ToString();
                            sBufferDet = "<td>" + sBufferDet + "</td>";
                            sw.WriteLine(sBufferDet);
                        }
                        sw.WriteLine("</tr>");
                    }

                }
                sw.Close();
            }
 
            catch (Exception exception)
            {
                var msg = "Error al ejecutar " + mtdName + ": " + exception.Message;
                //ActionLog.Error(msg, exception);
                CodErr = 1;
                MsgErr = exception.Message;
                mensaje.Text = MsgErr;
                Application.Exit();
            }
            finally
            {
                mensaje.Text = "Proceso finalizado.";
                Application.Exit();
            }
        }
        }
    }


