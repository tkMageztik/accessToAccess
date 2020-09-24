using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace BP.AtNinjas.GetOldTrackingData
{
    public class GetOldTrackingData
    {
        public GetOldTrackingData()
        {
            Execute();
        }

        private void Execute()
        {
            Comex();
            //Tesoreria();
            //Leasing();
            //Canje();
            //BaseDeDatos();
            //AlianzasComerciales();
            //Custodia();
            //Garantias();
            //GestionDeSoluciones();
        }

        private void Comex()
        {
            string databasePath = ConfigurationManager.AppSettings["Comex"];
            //ejecutar 6 am
            string date = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));
            try
            {
                string sql = "SELECT * FROM COMEX WHERE FORMAT(HORA_INGRESO,'yyyyMM') = '" + date + "'";
                GenerateRawReport(GetData(databasePath, sql), "COMEX");

                sql = "SELECT * FROM AUTORIZACION WHERE FORMAT(HORA_INGRESO,'yyyyMM') <= '" + date + "'";
                GenerateRawReport(GetData(databasePath, sql), "AUTORIZACION");
            }

            catch (Exception exc)
            {

            }
        }

        private void Tesoreria()
        {
            string date = string.Format("{0:yyyyMMdd}", DateTime.Now);
            try
            {
                string sql = "SELECT * FROM PROCESAMIENTO WHERE FORMAT(HORA_INGRESO,'yyyyMMdd') <= '" + date + "'";
                GenerateRawReport(GetData("TESORERIA", sql), "PROCESAMIENTO(?)");

                sql = "SELECT * FROM AUTORIZACION WHERE FORMAT(HORA_INGRESO,'yyyyMMdd') <= '" + date + "'";
                GenerateRawReport(GetData("TESORERIA", sql), "AUTORIZACION(?)");
            }

            catch (Exception exc)
            {

            }
        }

        private void Leasing()
        {
            string date = string.Format("{0:yyyyMMdd}", DateTime.Now);
            try
            {
                string sql = "SELECT * FROM PROCESAMIENTO WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD_LEASING", sql), "PROCESAMIENTO(?)");

                sql = "SELECT * FROM AUTORIZACION WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD_LEASING", sql), "AUTORIZACION(?)");
            }

            catch (Exception exc)
            {

            }
        }

        private void Canje()
        {
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));
            try
            {
                //FECHA_HORA > CONTIENE ?
                string sql = "SELECT * FROM ALIANZAS WHERE FECHA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("CANJE", sql), "?");
            }
            catch (Exception exc)
            {

            }
        }

        private void BaseDeDatos()
        {
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));
            try
            {
                //FECHA_HORA > CONTIENE ?
                string sql = "SELECT * FROM Base WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD", sql), "?");
            }

            catch (Exception exc)
            {

            }
        }

        private void AlianzasComerciales()
        {
            string date = string.Format("{0:yyyyMM}", DateTime.Now.AddDays(-1));
            try
            {
                string sql = "SELECT * FROM ALIANZAS WHERE FORMAT(FECHA,'yyyyMM') <= '" + date + "'";
                GenerateRawReport(GetData("ALIANZAS", sql), "ALIANZAS(?)");
            }

            catch (Exception exc)
            {

            }
        }

        private void Custodia()
        {
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));
            try
            {
                //FECHA_HORA > CONTIENE ?
                string sql = "SELECT * FROM TRF WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD_CUSTODIA", sql), "?");

                sql = "SELECT * FROM TRANSACCIONAL WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD_CUSTODIA", sql), "?");

                sql = "SELECT * FROM REQUERIMIENTOS WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD_CUSTODIA", sql), "?");
            }

            catch (Exception exc)
            {

            }
        }

        //CONTROL DE OPERACIONES / MESA DE CONSULTA OPERATIVA
        private void Garantias()
        {

        }
        private void GestionDeSoluciones()
        {
            string date = string.Format("{0:/MM/yyyy}", DateTime.Now.AddDays(-1));
            try
            {
                string sql = "SELECT * FROM Base1 WHERE FECHA_HORA LIKE '*" + date + "*'";
                GenerateRawReport(GetData("BD_GS", sql), "(?)");
            }

            catch (Exception exc)
            {

            }
        }

        private DataTable GetData(string databasePath, string sql)
        {
            OleDbConnection cn = new OleDbConnection(ConfigurationManager.ConnectionStrings["AccessConnection"].ToString().Replace("[databasePath]", databasePath));

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

        private void GenerateRawReport(DataTable dt, string sheet)
        {
            try
            {
                string fileName = "RAW_COMEX_" + string.Format("{0:yyyyMM}", DateTime.Now) + ".xlsx";

                FileInfo excelFile = new FileInfo(Path.Combine(ConfigurationManager.AppSettings["RawReportPath"], fileName));


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

                    excelWorksheet.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Light2);

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

                    excelPackage.Save();
                }
            }

            catch (Exception exc)
            {
                //Util.Util.LogProceso(exc.Message);
                throw exc;
            }

        }

    }
}
