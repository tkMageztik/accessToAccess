using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Text;

namespace BP.AtNinjas.GetOldTrackingData
{
    public class GetOldTrackingData
    {
        public string AccessConnection { get; set; }
        private string RawReportPath { get; set; }
        private string TemplatePath { get; set; }

        public GetOldTrackingData()
        {
            Execute();
        }

        private void Execute()
        {
            try
            {
                if (ConfigurationManager.ConnectionStrings["AccessConnection"] == null)
                {
                    throw new ArgumentNullException("AccessConnection", "El AccessConnection no existe en el archivo de configuración");
                }
                else
                {
                    AccessConnection = ConfigurationManager.ConnectionStrings["AccessConnection"].ToString();
                }

                if (ConfigurationManager.AppSettings["RawReportPath"] == null)
                {
                    throw new ArgumentNullException("RawReportPath", "El RawReportPath no existe en el archivo de configuración");
                }
                else
                {
                    RawReportPath = ConfigurationManager.AppSettings["RawReportPath"].ToString();
                }

                if (ConfigurationManager.AppSettings["TemplatePath"] == null)
                {
                    throw new ArgumentNullException("TemplatePath", "El TemplatePath no existe en el archivo de configuración");
                }
                else
                {
                    TemplatePath = ConfigurationManager.AppSettings["TemplatePath"].ToString();
                }

                //Comex();
                //Tesoreria();
                //Leasing();
                //Canje();
                //BaseDeDatos();
                AlianzasComerciales();
                //Custodia();
                ////Garantias();
                //GestionDeSoluciones();

            }
            catch (ArgumentNullException exc)
            {
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
            }
            catch (Exception exc)
            {

            }

        }

        //TODO: change name for just call returnDatabasepath or similar
        private string ValidateConfig(string keyName)
        {
            string databaseFilePath = ConfigurationManager.AppSettings[keyName];

            if (databaseFilePath == null) { throw new ArgumentNullException(keyName, "El " + keyName + " no existe en el archivo de configuración"); }

            if (!File.Exists(databaseFilePath))
            { throw new FileNotFoundException("La base de datos " + keyName + " no existe en la ruta indicada", databaseFilePath); }

            return databaseFilePath;
        }


        private void Comex()
        {
            string databasePath = ValidateConfig("Comex");
            //ejecutar 6 am
            string date = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));

            string sql = "SELECT * FROM COMEX WHERE FORMAT(HORA_INGRESO,'yyyyMM') = '" + date + "'";
            GenerateRawReport("RAW_COMEX_", GetData(databasePath, sql), "COMEX");

            sql = "SELECT * FROM AUTORIZACION WHERE FORMAT(HORA_INGRESO,'yyyyMM') <= '" + date + "'";
            GenerateRawReport("RAW_COMEX_", GetData(databasePath, sql), "AUTORIZACION");
        }

        private void Tesoreria()
        {
            string databasePath = ValidateConfig("Tesoreria");
            string date = string.Format("{0:yyyyMMdd}", DateTime.Now);

            string sql = "SELECT * FROM PROCESAMIENTO WHERE FORMAT(HORA_INGRESO,'yyyyMMdd') <= '" + date + "'";
            GenerateRawReport("RAW_TESORERIA_", GetData(databasePath, sql), "PROCESAMIENTO");

            sql = "SELECT * FROM AUTORIZACION WHERE FORMAT(HORA_INGRESO,'yyyyMMdd') <= '" + date + "'";
            GenerateRawReport("RAW_TESORERIA_", GetData(databasePath, sql), "AUTORIZACION");
        }

        private void Leasing()
        {
            string databasePath = ValidateConfig("Leasing");
            string date = string.Format("{0:yyyyMMdd}", DateTime.Now);

            string sql = "SELECT * FROM PROCESAMIENTO WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_LEASING_", GetData(databasePath, sql), "PROCESAMIENTO");

            sql = "SELECT * FROM AUTORIZACION WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_LEASING_", GetData(databasePath, sql), "AUTORIZACION");
        }

        private void Canje()
        {
            //Tiene pass
            string databasePath = ValidateConfig("Canje");
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));

            //FECHA_HORA > CONTIENE ?
            string sql = "SELECT * FROM ALIANZAS WHERE FECHA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_CANJE_", GetData(databasePath, sql), "?");
        }

        private void BaseDeDatos()
        {
            string databasePath = ValidateConfig("BaseDeDatos");
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));
            //FECHA_HORA > CONTIENE ?
            string sql = "SELECT * FROM Base WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("Reporte_BaseDeDatos_", GetData(databasePath, sql), "BASE");
        }

        private void AlianzasComerciales()
        {
            string databasePath = ValidateConfig("AlianzasComerciales");
            string date = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));

            string sql = "SELECT * FROM ALIANZAS WHERE FORMAT(FECHA,'yyyyMM') >= '" + date + "'";
            //GenerateRawReport("Reporte_Alianzas_" + date, GetData(databasePath, sql), "RAW_ALIANZAS");
            GenerateRawReport("Reporte_Alianzas_", GetData(databasePath, sql), "BASE");
        }

        private void Custodia()
        {
            string databasePath = ValidateConfig("Custodia");
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));

            //FECHA_HORA > CONTIENE ?
            string sql = "SELECT * FROM TRF WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_CUSTODIA_", GetData(databasePath, sql), "?");

            sql = "SELECT * FROM TRANSACCIONAL WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_CUSTODIA_", GetData(databasePath, sql), "?");

            sql = "SELECT * FROM REQUERIMIENTOS WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_CUSTODIA_", GetData(databasePath, sql), "?");
        }

        //CONTROL DE OPERACIONES / MESA DE CONSULTA OPERATIVA
        private void Garantias()
        {

        }
        private void GestionDeSoluciones()
        {
            string databasePath = ValidateConfig("GestionDeSoluciones");
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));

            string sql = "SELECT * FROM Base1 WHERE FECHA_HORA LIKE '*" + date + "*'";
            GenerateRawReport("RAW_GS_", GetData(databasePath, sql), "(?)");
        }

        private DataTable GetData(string databasePath, string sql)
        {
            string cns = ConfigurationManager.ConnectionStrings["AccessConnection"].ToString().Replace("[databasePath]", databasePath);

            OleDbConnection cn = new OleDbConnection(cns);

            try
            {
                //revisar el uso de Dapper.
                //cn.Open();
                using (cn.OpenAsync())
                {
                    using (OleDbCommand cmd = new OleDbCommand(sql, cn))
                    {
                        //Util.Util.LogProceso("paso x aqui en ExcecuteNonQuery Visado1 formato: ");
                        var dt = new DataTable();
                        dt.Load(cmd.ExecuteReader());
                        return dt;
                    }
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        private void GenerateRawReport(string reportName, DataTable dt, string sheet)
        {
            try
            {
                string fileName = reportName + string.Format("{0:yyyyMM}", DateTime.Now) + ".xlsx";

                FileInfo excelFile = new FileInfo(Path.Combine(this.RawReportPath, fileName));

                File.Copy(Path.Combine(this.TemplatePath, reportName + "_Template"), excelFile.FullName, true);

                using (var excelPackage = new ExcelPackage(excelFile))
                {
                    //var orderedProperties = (from property in typeof(HomologatedTrackingBE).GetProperties()
                    //                         where Attribute.IsDefined(property, typeof(DisplayAttribute))
                    //                         orderby ((DisplayAttribute)property
                    //                                  .GetCustomAttributes(typeof(DisplayAttribute), false)
                    //                                  .Single()).Order
                    //                         select property);

                    //ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("VISADO");
                    //excelWorksheet.Cells["A1"].LoadFromCollection<HomologatedTrackingBE>(bl.Visa, true, OfficeOpenXml.Table.TableStyles.Light1,
                    //    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public, orderedProperties.ToArray());

                    //excelWorksheet.Column(3).Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";
                    //excelWorksheet.Column(11).Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";
                    //excelWorksheet.Column(12).Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";
                    //excelWorksheet.Column(19).Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";
                    //excelWorksheet.Cells["A:U"].AutoFitColumns();
                    ExcelWorksheet excelWorksheet;

                    if (excelPackage.Workbook.Worksheets[sheet] == null)
                    {
                        excelWorksheet = excelPackage.Workbook.Worksheets.Add(sheet);
                    }
                    else
                    {
                        //excelWorksheet = excelPackage.Workbook.Worksheets[sheet];
                        excelPackage.Workbook.Worksheets.Delete(sheet);
                        excelWorksheet = excelPackage.Workbook.Worksheets.Add(sheet);
                    }

                    ExcelRangeBase excelRangeBase = excelWorksheet.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Light2);

                    if(reportName == "Reporte_Alianzas_")
                    {
                        AlianzasComerciales(excelPackage, excelRangeBase, excelWorksheet);
                    }

                    excelPackage.Save();
                }
            }

            catch (Exception exc)
            {
                //Util.Util.LogProceso(exc.Message);
                throw exc;
            }

        }

        private void AlianzasComerciales(ExcelPackage excelPackage, ExcelRangeBase excelRangeBase, ExcelWorksheet excelWorksheet)
        {
            int nextRow = excelRangeBase.End.Row + 1;
            int column = excelRangeBase.End.Column;

            //excelWorksheet.Cells[nextRow, 4, nextRow, column].Formula = "SUM()";

            //columna 4 == columna E
            excelWorksheet.Cells[nextRow, 5, nextRow, column].Formula = "SUM(E2:E" + excelRangeBase.End.Row + ")";

            //excelWorksheet.Cells[nextRow, 4, nextRow, column].Formula = "SUM(E2:E10)";
            excelWorksheet.Cells[nextRow, 5, nextRow, column].Calculate();

            excelWorksheet.Column(2).Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";

            //TODO: take a look in future for performance decrease
            excelWorksheet.Cells[1, 1, excelRangeBase.End.Row, column].AutoFitColumns();

            //TODO: el idioma... 
            excelPackage.Workbook.Worksheets["MATRIZ"].Cells["J1"].Value = DateTime.Now.ToString("MMMM", CultureInfo.GetCultureInfo("es-PE"));

            //Empieza desde la columna E osea al total de columnas hay que restarle 4, para que empiece desde la 5
            for (int i = 0; i <= column - 4; i++)
            {
                //excelPackage.Workbook.Worksheets["MATRIZ"].SetValue(2 + i, 10, excelWorksheet.Cells[nextRow, 5 + i].Value);
                excelPackage.Workbook.Worksheets["MATRIZ"].Cells[2 + i, 10].Value = excelWorksheet.Cells[nextRow, 5 + i].Value;
            }

            //string fileName = "RAW_COMEX_" + string.Format("{0:yyyyMM}", DateTime.Now) + ".xlsx";

            //FileInfo excelFile = new FileInfo(Path.Combine(ConfigurationManager.AppSettings["RawReportPath"], fileName));

            //try
            //{
            //    if (excelFile.Exists)
            //    {
            //        excelFile.Delete();
            //    }
            //    excelPackage.SaveAs(excelFile);
            //}
            //catch (IOException)
            //{
            //    fileName = fileName.Replace(".xlsx", "_" + String.Format("{0:HHmmss}", DateTime.Now) + ".xlsx");

            //    //Util.Util.LogProceso("El archivo está en uso o hubo algún otro problema que no permite sobreescribir el archivo, se ha generado otro archivo de nombre " + fileName);

            //    excelFile = new FileInfo(Path.Combine(ConfigurationManager.AppSettings["RawReportPath"], fileName));
            //    excelPackage.SaveAs(excelFile);
            //}
        }

    }
}
