using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPA_GenerateReportExcel_Health
{
    class Program
    {
        static void Main(string[] args)
        {
            var conn = new SqlConnection(Configulation.Db_r4ad01.ToString().Trim());
            conn.Open();
            try
            {
                using (var command = new SqlCommand("SpGetDataPolicyByPolicyNonMotor", conn))
                {
                    command.CommandText = @"SELECT
                                              [run_no] as [ลำดับ], 
                                              [pol_pre] as [subcalss], 
                                              [tr_date] as [วันที่ยอดขาย], 
                                              [pol_no] as [เลขกรมธรรม์], 
                                              CASE WHEN[endos_type] = 'NULL' THEN NULL ELSE[endos_type] END as [ประเภทสลักหลัง],
                                              [endos_no] as [เลขสลักหลัง], 
                                              [ins_name] as [ชื่อผู้เอาประกัน], 
                                              [start_date] as [วันที่เริ่มคุ้มครอง],
                                              [end_date] as [วันที่สิ้นสุด], 
                                              [planname] as [PlanName], 
                                              [sum_insured] as [ทุนความคุ้มครองฉุกเฉิน], 
                                              [status] as [สถานะกรมธรรม์], 
                                              CASE WHEN[cancel_date] = 'NULL' THEN NULL ELSE[cancel_date] END as [วันที่ยกเลิกกรมธรรม์]
                                            FROM IPA_Temp_GetData
                                            WHERE pol_pre_type = 'health'";

                    command.CommandTimeout = 15;
                    command.CommandType = CommandType.Text;
                    SqlDataReader dr = command.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    var resultGen = GenerateExcel(dt);
                    if (resultGen)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            //Send Mail FTP Sucessful
                        }
                        else
                        {
                            //Send Mail Not Found
                        }
                    }
                    else
                    {
                        //Send Mail Gen File Excel Fail
                    }

                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static bool GenerateExcel(DataTable dt)
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                var workbook = new ExcelFile();
                ExcelWorksheet worksheet = workbook.Worksheets.Add("Sheet 1");
                worksheet.InsertDataTable(dt,
                        new InsertDataTableOptions()
                        {
                            ColumnHeaders = true
                        });

                workbook.Save("DataExcel.xlsx");

                var workbook_last = ExcelFile.Load("DataExcel.xlsx");

                var worksheet_last = workbook_last.Worksheets.ActiveWorksheet;

                int columnCount = worksheet_last.CalculateMaxUsedColumns();
                for (int i = 0; i < columnCount; i++)
                {
                    worksheet_last.Columns[i].AutoFit(1, worksheet_last.Rows[1], worksheet_last.Rows[worksheet_last.Rows.Count - 1]);
                }

                workbook_last.Save(Configulation.PathFile.ToString().Trim() + "Row_Column_AutoFit.xlsx");

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        class Configulation
        {
            public static string Db_r4ad01 = ConfigurationManager.AppSettings["Db_r4ad01"].ToString();
            public static string PathFile = ConfigurationManager.AppSettings["PathFile"].ToString();
        };
    }
}
