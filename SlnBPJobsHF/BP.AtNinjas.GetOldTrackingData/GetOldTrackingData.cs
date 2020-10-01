using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using To.AtNinjas.Util;

namespace BP.AtNinjas.GetOldTrackingData
{
    public class GetOldTrackingData
    {
        public string AccessConnection { get; set; }
        public string AccessConnectionCanje { get; set; }

        private string RawReportPath { get; set; }
        private string TemplatePath { get; set; }
        private string BankReportPath { get; set; }

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

                if (ConfigurationManager.ConnectionStrings["AccessConnectionCanje"] == null)
                {
                    throw new ArgumentNullException("AccessConnectionCanje", "El AccessConnectionCanje no existe en el archivo de configuración");
                }
                else
                {
                    AccessConnectionCanje = ConfigurationManager.ConnectionStrings["AccessConnectionCanje"].ToString();
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

                if (ConfigurationManager.AppSettings["BankReportPath"] == null)
                {
                    throw new ArgumentNullException("BankReportPath", "El BankReportPath no existe en el archivo de configuración");
                }
                else
                {
                    BankReportPath = ConfigurationManager.AppSettings["BankReportPath"].ToString();
                }

                Util.LogProceso("Inicio el proceso");


                //The problem with the LIKE operator is that the Access database engine(ACE, Jet, whatever) uses 
                //    different wildcard characters depending on so - called ANSI Query Mode.Presumably you are 
                //    using ANSI-92 Query Mode by implication of using SqlOleDb (but would you have known for sure ?)

                Comex();
                Tesoreria();
                Leasing();
                Canje();
                BaseDeDatos();
                AlianzasComerciales();
                Custodia();
                //Garantias();
                GestionDeSoluciones();

                Util.LogProceso("Terminó el proceso");

            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }

        }

        //TODO: change name for just call returnDatabasepath or similar
        private string ValidateConfig(string keyName)
        {
            string databaseFilePath = ConfigurationManager.AppSettings[keyName];

            if (databaseFilePath == null) { throw new ArgumentNullException(keyName, "El " + keyName + " no existe en el archivo de configuración"); }

            if (!File.Exists(databaseFilePath))
            { throw new FileNotFoundException("La base de datos " + keyName + " no existe en la ruta indicada o no se cuenta con acceso", databaseFilePath); }
            //else
            //{
            //    File.Copy(databaseFilePath, Path.Combine(RawReportPath, Path.GetFileName(databaseFilePath)));
            //}

            return databaseFilePath;
            //return Path.Combine(RawReportPath, Path.GetFileName(databaseFilePath));
        }

        private void Comex()
        {
            try
            {
                string databasePath = ValidateConfig("Comex");
                //ejecutar 6 am
                string date = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));

                string sql = "SELECT * FROM COMEX WHERE FORMAT(HORA_INGRESO,'yyyyMM') = '" + date + "'";
                GenerateRawReport("RAW_COMEX_", GetData(databasePath, sql), "COMEX");

                sql = "SELECT * FROM AUTORIZACION WHERE FORMAT(HORA_INGRESO,'yyyyMM') <= '" + date + "'";
                GenerateRawReport("RAW_COMEX_", GetData(databasePath, sql), "AUTORIZACION");

                Util.LogProceso("Terminó Comex");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private void Tesoreria()
        {
            try
            {
                string databasePath = ValidateConfig("Tesoreria");
                string date = string.Format("{0:yyyyMM}", DateTime.Now);

                string sql = "SELECT * FROM PROCESAMIENTO WHERE FORMAT(HORA_INGRESO,'yyyyMM') >= '" + date + "'";
                GenerateRawReport("RAW_TESORERIA_", GetData(databasePath, sql), "PROCESAMIENTO");

                sql = "SELECT * FROM AUTORIZACION WHERE FORMAT(HORA_INGRESO,'yyyyMM') >= '" + date + "'";
                GenerateRawReport("RAW_TESORERIA_", GetData(databasePath, sql), "AUTORIZACION");

                Util.LogProceso("Terminó Tesoreria");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private void Leasing()
        {
            try
            {
                string databasePath = ValidateConfig("Leasing");
                string date = string.Format("{0:/MM/yyyy}", DateTime.Now);

                string sql = "SELECT * FROM PROCESAMIENTO WHERE FECHA_HORA LIKE '%" + date + "%'";
                GenerateRawReport("RAW_LEASING_", GetData(databasePath, sql), "PROCESAMIENTO");

                sql = "SELECT * FROM AUTORIZACION WHERE FECHA_HORA LIKE '%" + date + "%'";
                GenerateRawReport("RAW_LEASING_", GetData(databasePath, sql), "AUTORIZACION");

                Util.LogProceso("Terminó Leasing");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private void Canje()
        {
            try
            {
                //Tiene pass
                string databasePath = ValidateConfig("Canje");
                string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));

                //FECHA_HORA > CONTIENE ?
                string sql = "SELECT * FROM ALIANZAS WHERE FECHA LIKE '%" + date + "%'";
                GenerateRawReport("Reporte_Canje_", GetData(databasePath, sql), "BASE");

                Util.LogProceso("Terminó Canje");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private void BaseDeDatos()
        {
            try
            {
                string databasePath = ValidateConfig("BaseDeDatos");
                string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));
                //FECHA_HORA > CONTIENE ?
                string sql = "SELECT * FROM Base WHERE FECHA_HORA LIKE '%" + date + "%' AND ESTADO IN ('ATENDIDO','OBSERVADO')";
                GenerateRawReport("Reporte_BaseDeDatos_", GetData(databasePath, sql), "BASE");

                Util.LogProceso("Terminó BaseDeDatos");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private void AlianzasComerciales()
        {
            try
            {
                string databasePath = ValidateConfig("AlianzasComerciales");
                string date = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));

                string sql = "SELECT * FROM ALIANZAS WHERE FORMAT(FECHA,'yyyyMM') >= '" + date + "'";
                //GenerateRawReport("Reporte_Alianzas_" + date, GetData(databasePath, sql), "RAW_ALIANZAS");

                DataTable dt = GetData(databasePath, sql);

                GenerateRawReport("Reporte_Alianzas_", dt, "BASE");
                GenerateBankReport("BO_SCO_Reporte_Alianzas_Comerciales.xlsx", dt, "Alianzas_Comerciales");

                Util.LogProceso("Terminó AlianzasComerciales");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private void Custodia()
        {
            try
            {
                string databasePath = ValidateConfig("Custodia");
                string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));

                //FECHA_HORA > CONTIENE ?
                string sql = "SELECT * FROM TRF WHERE FECHA_HORA LIKE '%" + date + "%'";
                GenerateRawReport("RAW_CUSTODIA_", GetData(databasePath, sql), "TRF");

                sql = "SELECT * FROM TRANSACCIONAL WHERE FECHA_HORA LIKE '%" + date + "%'";
                GenerateRawReport("RAW_CUSTODIA_", GetData(databasePath, sql), "TRANSACCIONAL");

                sql = "SELECT * FROM REQUERIMIENTOS WHERE FECHA_HORA LIKE '%" + date + "%'";
                GenerateRawReport("RAW_CUSTODIA_", GetData(databasePath, sql), "REQUERIMIENTOS");

                Util.LogProceso("Terminó Custodia");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        //CONTROL DE OPERACIONES / MESA DE CONSULTA OPERATIVA
        private void Garantias()
        {

        }
        private void GestionDeSoluciones()
        {
            try
            {
                string databasePath = ValidateConfig("GestionDeSoluciones");
                string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));

                string sql = "SELECT * FROM Base1 WHERE FECHA_HORA LIKE '%" + date + "%'";
                GenerateRawReport("RAW_GS_", GetData(databasePath, sql), "BASE");

                Util.LogProceso("Terminó GestionDeSoluciones");
            }
            catch (ArgumentNullException exc)
            {
                Util.LogProceso(exc.Message);
                //Falta algún valor del app config.
            }
            catch (NullReferenceException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (FileNotFoundException exc)
            {
                //Falta algún valor del app config.
                Util.LogProceso(exc.Message);
            }
            catch (Exception exc)
            {
                Util.LogProceso(exc.Message);
                Util.LogProceso(exc.InnerException.Message);
            }
        }

        private DataTable GetData(string databasePath, string sql)
        {
            string cns = AccessConnection.Replace("[databasePath]", databasePath);

            if (databasePath.Contains("CANJE"))
            {
                cns = AccessConnectionCanje.Replace("[databasePath]", databasePath);
            }

            OleDbConnection cn = new OleDbConnection(cns);

            try
            {
                //revisar el uso de Dapper.
                cn.Open();
                //using (cn.OpenAsync())
                //{
                using (OleDbCommand cmd = new OleDbCommand(sql, cn))
                {
                    //Util.Util.LogProceso("paso x aqui en ExcecuteNonQuery Visado1 formato: ");
                    var dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                    return dt;
                }
                //}
            }
            catch (Exception exc)
            {
                throw exc;
            }
            finally { cn.Close(); }
        }

        private void GenerateBankReport(string reportName, DataTable dt, string sheet)
        {
            try
            {
                string folderName = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));

                folderName = Path.Combine(this.BankReportPath, folderName);

                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }

                FileInfo excelFile = new FileInfo(Path.Combine(folderName, reportName));

                //File.Copy(fileName, "BO_SCO_Repor");
                ExcelWorksheet excelWorksheet;

                using (var excelPackage = new ExcelPackage(excelFile))
                {
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

                    ExcelRangeBase excelRangeBase = null;

                    excelRangeBase = excelWorksheet.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Light2);

                    excelPackage.Save();
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
                string fileName = reportName + string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1)) + ".xlsx";

                FileInfo excelFile = new FileInfo(Path.Combine(this.RawReportPath, fileName));

                if (reportName == "Reporte_Alianzas_" || reportName == "Reporte_BaseDeDatos_" || reportName == "Reporte_Canje_")
                {
                    File.Copy(Path.Combine(this.TemplatePath, reportName + "Template.xlsx"), excelFile.FullName, true);
                }

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

                    ExcelRangeBase excelRangeBase = null;

                    //if (reportName != "Reporte_BaseDeDatos_")
                    //{
                    excelRangeBase = excelWorksheet.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Light2);
                    //}
                    //else
                    //{
                    //    excelRangeBase = excelWorksheet.Cells["C1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Light2);


                    //}

                    if (reportName == "Reporte_Alianzas_")
                    {
                        AlianzasComerciales(excelPackage, excelRangeBase, excelWorksheet);
                    }
                    else if (reportName == "Reporte_BaseDeDatos_")
                    {
                        BaseDeDatos(excelPackage, excelRangeBase, excelWorksheet);
                    }
                    else if(reportName == "Reporte_Canje_") {

                        Canje(excelPackage, excelRangeBase, excelWorksheet);
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
            excelPackage.Workbook.Worksheets["MATRIZ"].Cells["J1"].Value = DateTime.Now.AddDays(-1).ToString("MMMM", CultureInfo.GetCultureInfo("es-PE"));

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

        private void BaseDeDatos(ExcelPackage excelPackage, ExcelRangeBase excelRangeBase, ExcelWorksheet excelWorksheet)
        {
            //int nextRow = excelRangeBase.End.Row + 1;
            //TODO: el idioma... 
            excelPackage.Workbook.Worksheets["MATRIZ"].Cells["K1"].Value = DateTime.Now.AddDays(-1).ToString("MMMM", CultureInfo.GetCultureInfo("es-PE"));

            excelWorksheet.Cells[1, excelRangeBase.End.Column + 1].Value = "CONCATENADO";
            excelWorksheet.Cells[1, excelRangeBase.End.Column + 2].Value = "CODIGO";

            excelWorksheet.Cells[2, excelRangeBase.End.Column + 1, excelRangeBase.End.Row + 1, excelRangeBase.End.Column + 1].Formula = "G2&H2&I2";
            excelWorksheet.Cells[2, excelRangeBase.End.Column + 1, excelRangeBase.End.Row + 1, excelRangeBase.End.Column + +1].Calculate();

            //El separador de parámetros de fórmula es ',' por que se usa la notación en inglés, igualmente con las formulas.
            excelWorksheet.Cells[2, excelRangeBase.End.Column + 2, excelRangeBase.End.Row + 1, excelRangeBase.End.Column + 2].Formula = "VLOOKUP(U2,CONCATENADO!$D$2:$E$304,2,0)";
            //excelWorksheet.Cells[2, excelRangeBase.End.Column + 2, excelRangeBase.End.Row + 1, excelRangeBase.End.Column + 2].FormulaR1C1 = "VLOOKUP(U2,CONCATENADO!$D$2:$E$304,2,0)";
            excelWorksheet.Cells[2, excelRangeBase.End.Column + 2, excelRangeBase.End.Row + 1, excelRangeBase.End.Column + 2].Calculate();


            //TODO: take a look in future for performance decrease
            excelRangeBase.AutoFitColumns();

            excelPackage.Save();

            DataTable sheetData = new DataTable();
            try
            {
                using (OleDbConnection conn = this.returnConnection(""))
                {
                    conn.Open();
                    // retrieve the data using data adapter
                    OleDbDataAdapter sheetAdapter =
                        new OleDbDataAdapter(
                            "SELECT  CODIGO,CANT_OPERACIONES FROM (SELECT CODIGO,SUM(CANT_OPERACIONES) AS CANT_OPERACIONES " +
                            "FROM [BASE$] GROUP BY CODIGO) A " +
                            "RIGHT JOIN [MATRIZ$] B " +
                            "ON A.CODIGO = B.JEDOX"
                            , conn);

                    //select * from [MATRIZ$]
                    //select sum(CANT_OPERACIONES) AS T,CODIGO from[BASE$] group by CODIGO


                    sheetAdapter.Fill(sheetData);
                }
            }
            catch (Exception exc)
            {

            }

            excelPackage.Workbook.Worksheets["MATRIZ"].Cells[sheetData.Rows.Count + 3, 10].Value = "TOTAL:";

            for (int i = 0; i <= sheetData.Rows.Count - 1; i++)
            {
                //if(sheetData.Rows[i]["CANT_OPERACIONES"] == DBNull.Value ? 0: Convert.ToDecimal(sheetData.Rows[i]["CANT_OPERACIONES"])
                //    9

                excelPackage.Workbook.Worksheets["MATRIZ"].Cells["K" + (i + 2)].Value =
                    sheetData.Rows[i]["CANT_OPERACIONES"] == DBNull.Value ? 0 : Convert.ToDecimal(sheetData.Rows[i]["CANT_OPERACIONES"]);
                //sheetData.Rows[i]["CANT_OPERACIONES"];
            }

            excelPackage.Workbook.Worksheets["MATRIZ"].Cells[sheetData.Rows.Count + 3, 11].Formula = "SUM(K2:K" + (sheetData.Rows.Count + 1) + ")";
            excelPackage.Workbook.Worksheets["MATRIZ"].Cells[sheetData.Rows.Count + 3, 11].Calculate();

            //excelWorksheet.Cells[nextRow, 4, nextRow, column].Formula = "SUM(E2:E10)";
            //excelWorksheet.Cells[nextRow, 5, nextRow, column].Calculate();

            excelPackage.Save();
            //return sheetData;

            //TOTAL	18344	18344
        }

        private void Canje(ExcelPackage excelPackage, ExcelRangeBase excelRangeBase, ExcelWorksheet excelWorksheet)
        {
            int nextRow = excelRangeBase.End.Row + 1;
            int column = excelRangeBase.End.Column;

            excelPackage.Workbook.Worksheets["MATRIZ"].Cells["AB1"].Value = DateTime.Now.AddDays(-1).ToString("MMMM", CultureInfo.GetCultureInfo("es-PE"));

            excelWorksheet.Cells[nextRow, 5, nextRow, column].Formula = "SUM(E2:E" + excelRangeBase.End.Row + ")";

            //excelWorksheet.Cells[nextRow, 4, nextRow, column].Formula = "SUM(E2:E10)";
            excelWorksheet.Cells[nextRow, 5, nextRow, column].Calculate();

            excelWorksheet.Cells[1, 1, excelRangeBase.End.Row, column].AutoFitColumns();

            //Empieza desde la columna E osea al total de columnas hay que restarle 4, para que empiece desde la 5
            for (int i = 0; i <= column - 5; i++)
            {
                //excelPackage.Workbook.Worksheets["MATRIZ"].SetValue(2 + i, 10, excelWorksheet.Cells[nextRow, 5 + i].Value);
                excelPackage.Workbook.Worksheets["MATRIZ"].Cells[2 + i, 28].Value = excelWorksheet.Cells[nextRow, 5 + i].Value;
            }

            excelPackage.Workbook.Worksheets["MATRIZ"].Cells[nextRow + 3,27].Value = "TOTAL";
            excelPackage.Workbook.Worksheets["MATRIZ"].Cells[nextRow + 3, 28].Formula = "SUM(AB2:AB4)";

        }


        private OleDbConnection returnConnection(string fileName)
        {
            return new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuarios\juarui\Desktop\Documentación RPA's\borrar\Reporte_BaseDeDatos_202009.xlsx;Extended Properties='Excel 12.0 Xml; HDR = YES'");
            return new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + "; Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 8.0;\"");

        }
    }
}
