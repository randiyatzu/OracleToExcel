using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OracleToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Console.Title = String.Format("OracleToExcel (version {0})", version);

            if (!(args.Length == 3))
            {
                showMessageAndExit(
@"Error: Incorrect number of parameter

FORMAT:
OracleToExcel.exe userid/password@db ""excel output path"" ""sql file path | table.column""

EXAMPLE1 (sql file):
    OracleToExcel.exe hr/xxxx@orcl ""D:\excelfilename"" ""D:\sql.txt""

    文字檔模式不支援保留空白列欄、指定欄位值


EXAMPLE2 (table.column):
    OracleToExcel.exe hr/xxxx@orcl ""D:\excelfilename"" ""select 2, 1, 1, sql, assign_list from excel_tab""

    select欄位1: 保留空白列
    select欄位2: 保留空白欄
    select欄位3: 0|1 (不顯示表頭|顯示表頭)
    select欄位4: 資料內容的SQL
    select欄位5: 指定欄位值
"
);
            }

            string[] scs = splitConnectString(fixSql(args[0]));

            string userId = scs[0];
            string password = scs[1];
            string dbServer = scs[2];

            string excelFilePath = args[1];
            string sqlFrom = args[2];
            string outExcelFilePath;
            string createDateTime = DateTime.Now.ToString("yyyy-MM-dd(HH-mm-ss)");

            // 設定欄位為公式的指定開頭字串
            string formulaPattern = @"^<FORMULA>=";

            // 儲存需要調整的欄位寬度
            int[] columnWidth;

            // 上方保留空白列
            int emptyRow = 0;

            // 左方保留空白行
            int emptyCol = 0;

            // 顯示表頭
            int showTitle = 1;

            try
            {
                string oradb = String.Format("Data Source={0};User Id={1};Password={2};", dbServer, userId, password);

                Console.WriteLine(String.Format("Connecting to Oracle Database {0}...", dbServer));

                OracleConnection conn = new OracleConnection(oradb);
                conn.Open();

                // 已經有此檔名,需另取一個檔名
                if (File.Exists(String.Format("{0}.xlsx", excelFilePath)))
                    outExcelFilePath = String.Format("{0}-{1}.xlsx", excelFilePath, createDateTime);
                else
                    outExcelFilePath = String.Format("{0}.xlsx", excelFilePath);

                // 檢查Excel產生檔案的目錄存在否
                if (!Directory.Exists(Path.GetDirectoryName(outExcelFilePath)))
                {
                    showMessageAndExit(
                        String.Format(@"""{0}"" not exist", Path.GetDirectoryName(outExcelFilePath)));
                }

                OracleCommand cmd;
                OracleDataReader dr;
                string rootSql;
                string sql, assignValue = null;

                if (File.Exists(sqlFrom))
                {
                    sql = System.IO.File.ReadAllText(sqlFrom);
                }
                else
                {
                    rootSql = fixSql(sqlFrom);

                    if (!isSelectStatement(rootSql))
                        showMessageAndExit("Error1: Only select statements are allowed");

                    cmd = new OracleCommand(rootSql, conn);
                    cmd.CommandType = CommandType.Text;

                    dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        dr.Read();
                        emptyRow = dr.GetInt32(0);
                        emptyCol = dr.GetInt32(1);
                        showTitle = dr.GetInt32(2);
                        sql = dr.GetString(3);

                        // 2018.7.24 增加
                        // 指定值欄位
                        if (dr.FieldCount == 5)
                            // 2018.10.22 version:7.0.0.0 增加
                            // 檢查是否為null
                            if (!dr.IsDBNull(4))
                                assignValue = dr.GetString(4);
                    }
                    else
                        sql = "sql error";
                }


                sql = fixSql(sql);
                if (!isSelectStatement(sql))
                    showMessageAndExit("Error2: Only select statements are allowed");

                cmd = new OracleCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                dr = cmd.ExecuteReader();

                // 資料總筆數
                DataTable dt = new DataTable();
                dt.Load(dr);
                int rowCount = dt.Rows.Count;

                dr = cmd.ExecuteReader();

                // Create the file using the FileInfo object
                var file = new FileInfo(outExcelFilePath);

                // Create the package and make sure you wrap it in a using statement
                using (var package = new ExcelPackage(file))
                {
                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(
                        createDateTime);

                    // 欄位寬度,預設為0
                    columnWidth = new int[dr.FieldCount];
                    for (var i = 0; i < columnWidth.Length; i += 1)
                        columnWidth[i] = 0;

                    int line = 0;

                    // 2018.6.21 增加
                    // 預留敘述空白列,最後AutoFit之後再做
                    line = line + emptyRow;


                    // 表頭
                    // 2018.7.24 增加 抬頭顯示
                    if (showTitle != 0)
                    {
                        line++;
                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            int col = emptyCol + i + 1;

                            worksheet.Cells[line, col].Value = dr.GetName(i);
                            worksheet.Cells[line, col].Style.Font.Bold = true;

                            // 2018.2.12 增加
                            // 斷行判斷
                            if (dr.GetName(i).IndexOf('\n') != -1)
                            {
                                if (calBreakLineMaxWidth(dr.GetName(i)) > columnWidth[i])
                                    columnWidth[i] = calBreakLineMaxWidth(dr.GetName(i)) + 2; // 因設定title為粗體, 所以欄寬再+2
                                worksheet.Cells[line, col].Style.WrapText = true;
                            }
                        }
                    }

                    // 資料
                    int rec = 0;
                    while (dr.Read())
                    {
                        line++;
                        rec++;
                        Console.WriteLine("Export: {0}/{1}", rec, rowCount);

                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            int col = emptyCol + i + 1;

                            if (dr.GetProviderSpecificFieldType(i) == typeof(Oracle.ManagedDataAccess.Types.OracleDate))
                            {
                                if (!dr.IsDBNull(i))
                                {
                                    worksheet.Cells[line, col].Value = dr.GetDateTime(i);
                                    worksheet.Cells[line, col].Style.Numberformat.Format = "yyyy/mm/dd";
                                }
                            }
                            else if (dr.GetProviderSpecificFieldType(i) == typeof(Oracle.ManagedDataAccess.Types.OracleDecimal))
                            {
                                if (!dr.IsDBNull(i))
                                    worksheet.Cells[line, col].Value = dr.GetDouble(i);
                            }
                            else
                            {
                                if (!dr.IsDBNull(i))
                                {
                                    // 2018.10.22 version:7.0.0.0 增加
                                    // 檢查欄位值是否為公式
                                    if (Regex.IsMatch(dr.GetString(i), formulaPattern, RegexOptions.IgnoreCase))
                                    {
                                        string formula = Regex.Replace(dr.GetString(i), formulaPattern, "", RegexOptions.IgnoreCase);
                                        worksheet.Cells[line, col].Formula = formula;
                                    }
                                    else
                                    {
                                        worksheet.Cells[line, col].Value = dr.GetString(i);

                                        // 斷行判斷
                                        if (dr.GetString(i).IndexOf('\n') != -1)
                                        {
                                            if (calBreakLineMaxWidth(dr.GetString(i)) > columnWidth[i])
                                                columnWidth[i] = calBreakLineMaxWidth(dr.GetString(i));
                                            worksheet.Cells[line, col].Style.WrapText = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Auto Fit
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        int col = emptyCol + i + 1;

                        if (columnWidth[i] != 0)
                            worksheet.Column(col).Width = columnWidth[i];
                        else
                            worksheet.Column(col).AutoFit();
                    }

                    // 2018.7.24 增加
                    // 指定值欄位
                    if (assignValue != null)
                    {
                        // 斷行分每一列資料
                        string[] lines = assignValue.Split(
                            new[] { "\r\n", "\r", "\n" },
                            StringSplitOptions.None
                        );

                        int lineNo = 0;
                        foreach (var element in lines)
                        {
                            lineNo++;

                            string[] assignValues = analyzeAssignValues(element);

                            if (assignValues != null)
                            {
                                if (assignValues[2].ToUpper() == "N") // 數值
                                {
                                    // 2018.8.17 數值第4欄位為空白會出現 "Exception: 輸入字串格式不正確。" 錯誤, 針對空白或null不處理
                                    // 數值必去除頭尾空白
                                    if (!String.IsNullOrEmpty(assignValues[3].Trim()))
                                    {
                                        // 2018.8.17 數值轉換錯誤增加訊息提示
                                        try
                                        {
                                            double num = Convert.ToDouble(assignValues[3]);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Assign: {0}/{1} {2} '{3}' {4}", lineNo, lines.Length, "failed", element, ex.Message);
                                            continue;
                                        }

                                        worksheet.Cells[
                                            Int32.Parse(assignValues[0]),
                                            Int32.Parse(assignValues[1])].Value = Convert.ToDouble(assignValues[3]);
                                    }
                                }
                                else // == 'S' 字串
                                {
                                    // 2018.10.22 version:7.0.0.0 增加
                                    // 檢查欄位值是否為公式
                                    if (Regex.IsMatch(assignValues[3], formulaPattern, RegexOptions.IgnoreCase))
                                    {
                                        string formula = Regex.Replace(assignValues[3], formulaPattern, "", RegexOptions.IgnoreCase);
                                        worksheet.Cells[Int32.Parse(assignValues[0]), Int32.Parse(assignValues[1])].Formula = formula;
                                    }
                                    else
                                    {
                                        // 取代斷行
                                        string s = Regex.Replace(assignValues[3], "<ENTER>", "\r\n", RegexOptions.IgnoreCase);

                                        worksheet.Cells[Int32.Parse(assignValues[0]), Int32.Parse(assignValues[1])].Value = s;
                                        if (s.IndexOf('\n') != -1) // 斷行判斷
                                            worksheet.Cells[Int32.Parse(assignValues[0]), Int32.Parse(assignValues[1])].Style.WrapText = true;
                                    }
                                }
                                // 2018.8.17 成功增加訊息提示
                                Console.WriteLine("Assign: {0}/{1} {2}", lineNo, lines.Length, "Ok");
                            }
                            else
                            {
                                // 2018.8.17 格式錯誤訊息提示
                                Console.WriteLine("Assign: {0}/{1} {2} '{3}'", lineNo, lines.Length, "failed", element);
                            }
                        }
                    }
                    package.Save();
                }
                conn.Close();
                conn.Dispose();

                showMessageAndExit("Finished");
            }
            catch (Exception ex)
            {
                showMessageAndExit(String.Format("OracleException: {0}", ex.Message));
            }
        }

        // 2018.7.24 增加
        // 解析指定值欄位
        public static string[] analyzeAssignValues(string val)
        {
            string[] reArr = new string[4];

            // 取位置與類別設定
            string positionPattern = @"^[\d]+,[\d]+,[S|N|s|n],";
            Match m = Regex.Match(val, positionPattern);

            if (m.Success)
            {
                string[] splitStr = Regex.Split(m.Value, ",");
                reArr[0] = splitStr[0];
                reArr[1] = splitStr[1];
                reArr[2] = splitStr[2];
                reArr[3] = val.Substring(m.Length);

                return reArr;
            }
            else
            {
                return null;
            }
        }

        // 去除開頭空白,結尾斷行、空白、分號
        public static string fixSql(string sql)
        {
            string replaceStr;

            string patternTail = "[\r\n \t;]+$";
            string patternHead = "^[ ]+";
            string replacement = "";
            replaceStr = Regex.Replace(sql, patternTail, replacement);
            replaceStr = Regex.Replace(replaceStr, patternHead, replacement);
            return replaceStr;
        }

        // 是否為select、with開頭
        public static bool isSelectStatement(string sql)
        {
            string pattern = "^select|^with";

            return Regex.IsMatch(sql, pattern, RegexOptions.IgnoreCase);
        }

        // 解析連線資料庫字串
        public static string[] splitConnectString(string connectString)
        {
            string[] reConn = new string[3];

            // var connSplit = connectString.Split(new Char[] { '\\', '/', (char)64 }); // @ = (char)64
            // 2019.6.6 發現部份電腦無法用char split,修改為使用string
            var connSplit = connectString.Split(new string[] { "\\", "/", "@" }, StringSplitOptions.None); // @ = (char)64

            for (int i = 0; i < connSplit.Length; i++)
                reConn[i] = connSplit[i];
            return reConn;
        }

        // 計算斷行的文字最長的那一行長度
        public static int calBreakLineMaxWidth(string txt)
        {
            int max = 0;
            string[] split = txt.Split(new Char[] { '\n' });
            for (int i = 0; i < split.Length; i++)
                if (Encoding.Default.GetBytes(split[i]).Length > max)
                    max = Encoding.Default.GetBytes(split[i]).Length;
            return max;
        }

        public static void showMessageAndExit(string msg)
        {
            Console.WriteLine(msg);
            Console.WriteLine();
            Console.WriteLine("Press Any Key To Exit");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
