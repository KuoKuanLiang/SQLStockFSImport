using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace FSImport
{
    public class Func
    {
        
        /// <summary>
        /// 選擇excel檔案(可多選)
        /// </summary>
        /// <returns></returns>
        public List<string> GetFile()
        {
            List<string> list = new List<string>();

            OpenFileDialog files = new OpenFileDialog();
            //開啟多選功能
            files.Multiselect = true;
            //開啟視窗，回傳值為ok
            if (files.ShowDialog() == DialogResult.OK)
            {
                //取得路徑+名稱+副檔名
                foreach (string str in files.FileNames)
                {
                    list.Add(str);
                }
            }
            return list;
        }

        /// <summary>
        /// 取得excel內的資料
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public List<XlsData> GetExcelData(string filepath)
        {
            List<XlsData> list = new List<XlsData>();

            Excel.Application app = new Excel.Application();//excel應用程式
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range rdata = null;
            try
            {
                wb = app.Workbooks.Open(filepath);//應用程式裡的活頁簿
                ws = wb.Worksheets[1];//活頁簿裡的第一個工作表
                rdata = ws.UsedRange;//取得所有資料

                //取得已用範圍的最後一個儲存格
                Excel.Range count = ws.UsedRange.SpecialCells( Excel.XlCellType.xlCellTypeLastCell);
                int rownum = count.Row;//共多少橫列
                //int colnum = count.Column;//共多少直欄(19)

                //放到list中
                for (int i = 2; i <= rownum; i++)//欄位名稱那列不用取得，從2開始
                {
                    //excel欄名：A、D、E、H、K 、N 、P 、S
                    // 對應編號：1、4、5、8、11、14、16、19
                    list.Add(new XlsData
                    {
                        //-----------------------------------------------------
                        //cusID = rdata.Cells[i, 1].Text.Trim(),   //外包出貨單號
                        //prdID = rdata.Cells[i, 4].Text.Trim(),   //品牌代號
                        //prdName = rdata.Cells[i, 5].Text.Trim(), //品牌名稱
                        //prdSer = rdata.Cells[i, 8].Text.Trim(),  //型號
                        //cusDate = rdata.Cells[i, 11].Text.Trim(),//廠商點收日期
                        //qty = rdata.Cells[i, 14].Text.Trim(),    //數量
                        //oemID = rdata.Cells[i, 16].Text.Trim(),  //代工回修單號
                        //proID = rdata.Cells[i, 19].Text.Trim()   //回廠出貨單號
                        //------------------------------------------------------
                        cusID = rdata.Cells[i, 1].Text.Trim(),   //外包出貨單號
                        prdID = rdata.Cells[i, 7].Text.Trim(),   //品牌代號
                        prdName = rdata.Cells[i, 8].Text.Trim(), //品牌名稱
                        prdSer = rdata.Cells[i, 11].Text.Trim(),  //型號
                        cusDate = rdata.Cells[i, 14].Text.Trim(),//廠商點收日期
                        qty = rdata.Cells[i, 17].Text.Trim(),    //數量
                        oemID = rdata.Cells[i, 19].Text.Trim(),  //代工回修單號
                        proID = rdata.Cells[i, 23].Text.Trim()   //回廠出貨單號

                    });
                    //Console.WriteLine(xlRange.Cells[i, j].Text);
                }
            }
            catch (Exception ex)
            {
                //開啟時產生奇怪的問題，可能是檔案不存在，或路徑錯誤
                GC.Collect();//強行銷燬
            }
            finally 
            {
                //關閉excel
                ws = null;
                wb.Close();
                wb = null;
                rdata = null;
                app.Quit();
                app = null;
            }
          
            //回傳結果
            return list;
        }

        /// <summary>
        /// 匯出整理後excel
        /// </summary>
        /// <param name="prdList"></param>
        /// <param name="cusList"></param>
        /// <param name="proList"></param>
        public void ExportExcelData(List<PrdData> prdList, List<CusData> cusList, List<ProData> proList)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;// 使對象可見。

            // 創建一個新的空工作簿並將其添加到屬性 Workbooks 返回的集合中。
            Excel.Workbook wb = app.Workbooks.Add();

            Excel.Worksheet ws1 = app.ActiveSheet;//wb建立時會存在一個預設工作表
            Excel.Worksheet ws2 = wb.Worksheets.Add();//增加第2個工作表
            Excel.Worksheet ws3 = wb.Worksheets.Add();//增加第3個工作表
            try 
            {
                //第1工作表添加料品資料
                int row1 = 1;
                foreach (var s in prdList)
                {
                    ws1.Cells[row1, "A"] = s.state;
                    ws1.Cells[row1, "B"] = s.prdID;
                    ws1.Cells[row1, "C"] = s.prdName;
                    ws1.Cells[row1, "D"] = s.prdSer;
                    row1++;
                }

                //第2工作表添加客戶來料資料
                int row2 = 1;
                foreach (var s in cusList)
                {
                    ws2.Cells[row2, "A"] = s.state;
                    ws2.Cells[row2, "B"] = s.cusID;
                    ws2.Cells[row2, "C"] = s.oemID;
                    ws2.Cells[row2, "D"] = s.prdID;
                    ws2.Cells[row2, "E"] = s.prdName;
                    ws2.Cells[row2, "F"] = s.prdSer;
                    ws2.Cells[row2, "G"] = s.date;
                    ws2.Cells[row2, "H"] = s.qty;
                    row2++;
                }

                //第3工作表添加出貨資料
                int row3 = 1;
                foreach (var s in proList)
                {
                    ws3.Cells[row3, "A"] = s.state;
                    ws3.Cells[row3, "B"] = s.proID;
                    ws3.Cells[row3, "C"] = s.oemID;
                    ws3.Cells[row3, "D"] = s.prdID;
                    ws3.Cells[row3, "E"] = s.prdName;
                    ws3.Cells[row3, "F"] = s.prdSer;
                    ws3.Cells[row3, "G"] = s.date;
                    ws3.Cells[row3, "H"] = s.qty;
                    row3++;
                }

                //自動調整欄寬
                for (int i = 1; i <= 8; i++)//最多道第8(H)欄
                {
                    ws1.Columns[i].AutoFit();
                    ws2.Columns[i].AutoFit();
                    ws3.Columns[i].AutoFit();
                }
                //儲存檔案
                wb.SaveAs(@"C:\Users\user\Desktop\result_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss"), Excel.XlSaveAsAccessMode.xlNoChange);

            }
            catch (Exception ex)
            {

            }
            finally
            {
                wb.Close(false);
                app.Workbooks.Close();
                app.Quit();

                //刪除 Windows工作管理員中的Excel.exe 進程。
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app.Workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                app = null;
                wb = null;
                ws1 = null;
                ws2 = null;
                ws3 = null;
                //aRange = null;
                //呼叫垃圾回收
                GC.Collect();//強行銷燬 
            }
        }


        /// <summary>
        /// 取得日期
        /// </summary>
        /// <param name="orgStr"></param>
        /// <returns></returns>
        public string GetDate(string str1)
        {
            // 移除非數字、空白
            string str2 = Regex.Replace(str1, "[^0-9]", "");
            //只取前8碼
            string str3 = str2.Substring(0, 8);
            //回廠出貨單、外包出貨單、代工回修單，去除前面英文，後面8碼都是日期。
            return str3;
        }


        /// <summary>
        /// 處理成資料庫認得的日期格式 (年-月-日)
        /// </summary>
        /// <param name="str1"></param>
        /// <returns></returns>
        public string HandleDate(string str1)
        {
            // 取得前四碼
            string str2 = str1.Substring(0,4);
            // 取得中間兩碼
            string str3 = str1.Substring(4, 2);
            // 取得後面兩碼
            string str4 = str1.Substring(6);
            // 格式：年-月-日
            string str5 = str2 + "-" + str3 + "-" + str4;
            //回傳值
            return str5;
        }


        /// <summary>
        /// 取得所有excel檔案的內容，並分配在三個list中。
        /// </summary>
        /// <param name="dataGridView1"></param>
        /// <returns></returns>
        public Tuple<List<PrdData>, List<CusData>, List<ProData>> CreatDataList(DataGridView dataGridView1)
        {
            //----------------------------------------------------------------------------------------

            //取得列表中，所有excel檔案的內容。

            List<XlsData> list = new List<XlsData>();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                list.AddRange( GetExcelData(dataGridView1.Rows[i].Cells[0].Value.ToString()));
            }

            //移除list中重複的項目。
            list = list.Distinct().ToList();

            //----------------------------------------------------------------------------------------

            //將所有資料，分配給三個list。

            List<PrdData> prdlist = new List<PrdData>();//料品
            List<CusData> cuslist = new List<CusData>();//客戶來料
            List<ProData> prolist = new List<ProData>();//出貨

            foreach (XlsData data in list)
            {
                //料品
                prdlist.Add(new PrdData { prdID = data.prdID, prdName = data.prdName, prdSer = data.prdSer });
                //客戶來料
                if (data.cusID != null && data.cusID != "")
                {
                    // 日期要處理過
                    cuslist.Add(new CusData { cusID = data.cusID, oemID = data.oemID, prdID = data.prdID, prdName = data.prdName, prdSer = data.prdSer, date = GetDate(data.cusDate), qty = data.qty });
                }
                //出貨
                if (data.proID != null && data.proID != "")
                {
                    // 日期要處理過
                    prolist.Add(new ProData { proID = data.proID, oemID = data.oemID, prdID = data.prdID, prdName = data.prdName, prdSer = data.prdSer, date = GetDate(data.proID), qty = data.qty });
                }
            }

            //移除list中重複的項目。
            prdlist = prdlist.Distinct().ToList();
            cuslist = cuslist.Distinct().ToList();
            prolist = prolist.Distinct().ToList();

            //----------------------------------------------------------------------------------------

            return Tuple.Create(prdlist, cuslist, prolist);
        }


        /// <summary>
        /// 判斷料品「狀態」應為何?(重複、新、異常)
        /// </summary>
        /// <param name="prdlist"></param>
        /// <returns></returns>
        public List<PrdData> JudgePrdList(List<PrdData> prdlist , SqlTool sqltool)
        {
            foreach (PrdData s in prdlist)
            {
                //是否重複 (重點:prdname 等於)
                int result = sqltool.SqlJudge("USE [sqlstock_vn01] SELECT top 1 1 FROM product WHERE prdno = '" + s.prdID + "' AND prdno_ser = '" + s.prdSer + "' AND prdname = '" + s.prdName + "'");

                //是否異常 (重點:prdname 不等於)
                int result2 = sqltool.SqlJudge("USE [sqlstock_vn01] SELECT top 1 1 FROM product WHERE prdno = '" + s.prdID + "' AND prdno_ser = '" + s.prdSer + "' AND prdname <> '" + s.prdName + "'");

                if (result == 0 && result2 == 0)
                {
                    s.state = "新";
                }
                else if (result == 1 && result2 == 0)
                {
                    s.state = "重複";
                }
                else if (result == 0 && result2 == 1)
                {
                    s.state = "異常";
                }
                else if (result == 1 && result2 == 1)
                {
                    s.state = "異常";
                }
            }
            return prdlist;
        }


        /// <summary>
        /// 判斷客戶來料「狀態」應為何?(重複、新、異常)
        /// </summary>
        /// <param name="prdlist"></param>
        /// <returns></returns>
        public List<CusData> JudgeCusList(List<CusData> cuslist , SqlTool sqltool)
        {
            foreach (CusData s in cuslist)
            {
                // 外包出貨單號是否存在 (1：存在、0：不存在)
                int result = sqltool.SqlJudge("USE [sqlstock_vn01] SELECT top 1 1 FROM custominmaterialmain WHERE oem_cus_voucherno = '" + s.cusID + "'");

                // 代工回修單號是否存在 (1：存在、0：不存在)
                int result2 = sqltool.SqlJudge("USE [sqlstock_vn01] SELECT top 1 1 FROM custominmaterialdetail WHERE oem_voucherno = '" + s.oemID + "'");

                if (result == 1)
                {
                    if (result2 == 1)
                    {
                        s.state = "重複";
                    }
                    else
                    {
                        //外包出貨單，不可重複使用。
                        s.state = "異常";
                    }
                }
                else
                {
                    if (result2 == 1)
                    {
                        //代工回修單號，在同功能下不可重複使用。
                        s.state = "異常";
                    }
                    else
                    {
                        s.state = "新";
                    }
                }
            }
            return cuslist;
        }


        /// <summary>
        /// 判斷客戶來料「狀態」應為何?(重複、新、異常)
        /// </summary>
        /// <param name="prdlist"></param>
        /// <returns></returns>
        public List<ProData> JudgeProList(List<ProData> prolist, SqlTool sqltool)
        {
            foreach (ProData s in prolist)
            {
                // 回廠出貨單號是否存在 (1：存在、0：不存在)
                int result = sqltool.SqlJudge("USE [sqlstock_vn01] SELECT top 1 1 FROM productoutmain WHERE oem_pro_voucherno = '" + s.proID + "'");

                // 代工回修單號是否存在 (1：存在、0：不存在)
                int result2 = sqltool.SqlJudge("USE [sqlstock_vn01] SELECT top 1 1 FROM productoutdetail WHERE oem_voucherno = '" + s.oemID + "'");

                if (result == 1)
                {
                    if (result2 == 1)
                    {
                        s.state = "重複";
                    }
                    else
                    {
                        //外包出貨單，不可重複使用。
                        s.state = "異常";
                    }
                }
                else
                {
                    if (result2 == 1)
                    {
                        //代工回修單號，在同功能下不可重複使用。
                        s.state = "異常";
                    }
                    else
                    {
                        s.state = "新";
                    }
                }
            }
            return prolist;
        }


        /// <summary>
        /// 料品資料新增
        /// </summary>
        /// <param name="prdlist"></param>
        /// <param name="errorCatch"></param>
        /// <returns></returns>
        public Tuple<List<PrdData>, string> InsertPrdList(List<PrdData> prdlist, string errorCatch, SqlTool sqltool)
        {
            //處理料品，新、重複資料。
            foreach (PrdData s in prdlist)
            {
                // 查詢此料品是否存在，並回傳id。
                string sqlStr1 = "USE [sqlstock_vn01] SELECT product_id, prdno_id FROM product WHERE prdno = '" + s.prdID + "' AND prdno_ser = '" + s.prdSer + "' AND prdname = '" + s.prdName + "'";

                List<string> list = new List<string>();
                list.AddRange(sqltool.SqlSelect(sqlStr1));

                //有值，表示存在。
                if (list.Count > 0)
                {
                    s.product_id = list[0].ToString(); //料品主鍵id
                    s.prdno_id = list[1].ToString(); //料號id
                }
                //沒有值，表示是新資料。
                else
                {
                    //----------------------------------------------------------------------

                    //處理所需的值。

                    // 查詢product_id的最大值。
                    string sqlStr2 = "USE [sqlstockmessage_vn01] SELECT [key_value] FROM [sqlstockmessage_vn01].[dbo].[setupid] where [key_name] = 'PRODUCT'";// "SELECT MAX(product_id) FROM product";
                    List<string> list2 = new List<string>();
                    list2.AddRange(sqltool.SqlSelect(sqlStr2));
                    s.product_id = (Convert.ToInt32(list2[0].ToString()) + 1).ToString(); // 料品主鍵id：最大值+1

                    // (相同品牌代號，會有相同的prdno_id)
                    // 查詢此料品代號是否存在，並回傳prdno_id 。
                    string sqlStr3 = "USE [sqlstock_vn01] SELECT top 1 prdno_id FROM product WHERE prdno = '" + s.prdID + "'";
                    List<string> list3 = new List<string>();
                    list3.AddRange(sqltool.SqlSelect(sqlStr3));
                    if (list.Count > 0) // 有找到，表示有使用過此品牌代號
                    {
                        s.prdno_id = list3[0].ToString();
                    }
                    else //沒找到，表示：1.以前沒使用過。2.以前沒使用過，但此次新增中可能有重複的存在。
                    {
                        //搜尋列表的s.prdID，看相同品牌代號的資料，是否有s.prdno_id存在
                        PrdData temp = prdlist.Find(x=>x.prdID == s.prdID && x.prdno_id != null);

                        if (temp != null)//資料庫沒使用過、但此次新增是第二此以上的使用到。
                        {
                            s.prdno_id = temp.prdno_id;
                        }
                        else //資料庫沒使用過、此次新增也是第一次使用到
                        {
                            // 查詢prdno_id的最大值。(!!!是從sqlstockmessage_vn01資料庫中setupid資料表取得!!!)
                            string sqlStr4 = "USE [sqlstock_vn01] SELECT MAX(prdno_id) FROM product";

                            List<string> list4 = new List<string>();
                            list4.AddRange(sqltool.SqlSelect(sqlStr4));

                            s.prdno_id = (Convert.ToInt32(list4[0].ToString()) + 1).ToString();//料號id：最大值+1
                        }
                    }

                    //----------------------------------------------------------------------

                    // 新增料品資料。

                    string sqlStr5 = "USE [sqlstock_vn01] INSERT INTO product" +
                                    "(" +
                                    "product_id" +
                                    ",prdno_id" +
                                    ",prdno" +
                                    ",prdno_ser" +
                                    ",prdname" +
                                    ",spec" +
                                    ",descr" +
                                    ",machine_prd" +
                                    ",unit" +
                                    ",acurrency" +
                                    ",exchange_rate" +
                                    ",single_material" +
                                    ",final_purchase" +
                                    ",def_warehouse_id" +
                                    ",prd_status" +
                                    ")" +
                                    "VALUES" +
                                    "(" + s.product_id +
                                    ", " + s.prdno_id +
                                    ", '" + s.prdID + "' " +
                                    ", '" + s.prdSer + "' " +
                                    ", '" + s.prdName + "' " +
                                    ", " + " '' " +
                                    ", " + " '自動匯入' " +
                                    ", " + " 0 " +
                                    ", " + " 'pcs' " +
                                    ", " + " 'VND' " +
                                    ", " + " 3600.00 " +
                                    ", " + " 1 " +
                                    ", " + " 0 " +
                                    ", " + " 81 " +
                                    ", " + " 1 " +
                                    ")";
                    int result = sqltool.SqlInsert(sqlStr5);
                    if (result == 0)//失敗
                    {
                        //紀錄錯誤
                        errorCatch = errorCatch + "錯誤";
                        //避免讓客戶來料、出貨有料品的id可以使用。
                        s.product_id = null;
                        s.prdno_id = null;
                        //MessageBox.Show("新增「料品」資料，有錯誤。");
                    }
                    else //更新setupid資料表中PRODUCT資料的最大值，不然sqlstock使用會出錯。
                    {
                        string sqlStr6 = "USE [sqlstockmessage_vn01] UPDATE [dbo].[setupid] SET [key_value] = "+ s.product_id + " WHERE [key_name] = 'PRODUCT'";
                        int result2 = sqltool.SqlInsert(sqlStr6);
                        if (result2 == 0)//失敗
                        {
                            //紀錄錯誤
                            errorCatch = errorCatch + "錯誤";
                            //避免讓客戶來料、出貨有料品的id可以使用。
                            s.product_id = null;
                            s.prdno_id = null;
                            //MessageBox.Show("修改「[sqlstockmessage_vn01].[setupid]資料，有錯誤。");
                        }
                    }
                }
            }
            return Tuple.Create(prdlist, errorCatch);
        }


        /// <summary>
        /// 客戶來料資料新增
        /// </summary>
        /// <param name="cuslist"></param>
        /// <param name="errorCatch"></param>
        /// <returns></returns>
        public Tuple<List<CusData>, string> InsertCusList(List<CusData> cuslist, string errorCatch, SqlTool sqltool)
        {
            foreach (CusData s in cuslist)
            {
                //查詢 外包出貨單，是否存在，並會回傳 custominmaterialmain_id。(1：存在、null：不存在)
                string sqlStr1 = "USE [sqlstock_vn01] SELECT custominmaterialmain_id FROM custominmaterialmain WHERE oem_cus_voucherno = '" + s.cusID + "'";

                List<string> list = new List<string>();
                list.AddRange(sqltool.SqlSelect(sqlStr1));

                string custominmaterialmain_id = null; //客戶來料主鍵

                //存在：主檔已經建立過。
                if (list.Count > 0)
                {
                    custominmaterialmain_id = list[0].ToString();
                }
                //不存在：進行主檔的添加。
                else if (list.Count == 0)
                {
                    //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    //處理所需的值。

                    // 查詢出 主鍵id的最大值、最後的來料單號。
                    string sqlStr2 = "USE [sqlstock_vn01] SELECT TOP 1 custominmaterialmain_id , diff_voucherno FROM custominmaterialmain ORDER BY custominmaterialmain_id DESC";
                    
                    List<string> list2 = new List<string>();
                    list2.AddRange(sqltool.SqlSelect(sqlStr2));

                    string searchdate = ""; //來料單號，去除最後流水3碼
                    string diff_voucherno = ""; //來料單號 
                    string today = DateTime.Now.ToString("yyyyMMdd");//今日日期
                    if (list2.Count == 0)//沒有資料
                    {
                        custominmaterialmain_id = "1";
                        diff_voucherno = "FI" + today + "001";
                        searchdate = "FI" + today;
                    }
                    else
                    {
                        custominmaterialmain_id = (Convert.ToInt32(list2[0].ToString()) + 1).ToString(); // 客戶來料主鍵 + 1

                        string lastDate = GetDate(list2[1].ToString());//最後一筆資料的日期
                        // 今日的001來料單號，已經產生。
                        if (lastDate.Equals(today))
                        {
                            int num = Convert.ToInt32(list2[1].Substring(10)) + 1;

                            diff_voucherno = "FI" + lastDate + num.ToString("D3");//不足三位數時，前面的位數以0補足。
                            searchdate = "FI" + lastDate;
                        }
                        // 今日的001來料單號，還未產生。
                        else
                        {
                            diff_voucherno = "FI" + today + "001";
                            searchdate = "FI" + today;
                        }
                    }

                    string diff_date = HandleDate(s.date);//來料日期

                    //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    //新增客戶來料主檔

                    if (s.product_id != null)
                    {
                        //主檔的新增
                        string sqlStr3 = "USE [sqlstock_vn01] INSERT INTO custominmaterialmain" +
                                        "(" +
                                        "custominmaterialmain_id" +
                                        ",diff_voucherno" +
                                        ",searchdate" +
                                        ",diff_date" +
                                        ",custom_id" +
                                        ",warehouse_id" +
                                        ",stuffno" +
                                        ",stuffname" +
                                        ",memodescr" +
                                        ",isadjust" +
                                        ",isclosesheet" +
                                        ",oem_cus_voucherno" +
                                        ")" +
                                        "VALUES" +
                                        "(" + custominmaterialmain_id +
                                        ", '" + diff_voucherno + "' " +
                                        ", '" + searchdate + "' " +
                                        ", '" + diff_date + "' " +
                                        ", " + " 244 " +
                                        ", " + " 81 " +
                                        ", " + " 'A0-001' " +
                                        ", " + " 'auto' " +
                                        ", " + " '自動匯入' " +
                                        ", " + " 0 " +
                                        ", " + " 0 " +
                                        ", '" + s.cusID + "' " +
                                        ")";
                        int result = sqltool.SqlInsert(sqlStr3);
                        if (result == 0)//失敗
                        {
                            //紀錄錯誤
                            errorCatch = errorCatch + "錯誤";
                            //避免明細表新增時，有值可以新增。
                            custominmaterialmain_id = null;
                            //MessageBox.Show("新增「客戶來料單」資料，有錯誤");
                        }
                    }
                }

                // 確保主檔沒新增時，明細表也不會新增。
                if (custominmaterialmain_id != null) 
                {
                    //明細檔的添加
                    string sqlStr4 = "USE [sqlstock_vn01] INSERT INTO custominmaterialdetail" +
                                     "(" +
                                     "custominmaterialmain_id" +
                                     ",product_id" +
                                     ",qty" +
                                     ",warehouse_id" +
                                     ",inventory_style_id" +
                                     ",memodescr" +
                                     ",adflag" +
                                     ",printtimes" +
                                     ",oem_voucherno" +
                                     ")" +
                                     " VALUES " +
                                     "(" + custominmaterialmain_id +
                                     ", " + s.product_id +
                                     ", " + s.qty +
                                     ", " + " 81 " +
                                     ", " + " 182 " +
                                     ", " + " '自動轉入' " +
                                     ", " + " 1 " +
                                     ", " + " 0 " +
                                     ", '" + s.oemID + "' " +
                                     ")";

                    int result2 = sqltool.SqlInsert(sqlStr4);
                    if (result2 == 0)//失敗
                    {
                        //紀錄錯誤
                        errorCatch = errorCatch + "錯誤";
                        //MessageBox.Show("新增「客戶來料單」，有錯誤");
                    }
                }
            }
            return Tuple.Create(cuslist, errorCatch);
        }


        /// <summary>
        /// 出貨資料新增
        /// </summary>
        /// <param name="cuslist"></param>
        /// <param name="errorCatch"></param>
        /// <returns></returns>
        public Tuple<List<ProData>, string> InsertProList(List<ProData> prolist, string errorCatch, SqlTool sqltool)
        {
            foreach (ProData s in prolist)
            {
                //查詢 外包出貨單，是否存在，並會回傳 productoutmain_id。(1：存在、null：不存在)
                string sqlStr1 = "USE [sqlstock_vn01] SELECT productoutmain_id FROM productoutmain WHERE oem_pro_voucherno = '" + s.proID + "'";

                List<string> list = new List<string>();
                list.AddRange(sqltool.SqlSelect(sqlStr1));

                // 出貨主鍵
                string productoutmain_id = null;

                //存在：主檔已經建立過，不用添加。
                if (list.Count > 0)
                {
                    productoutmain_id = list[0].ToString();//出貨主鍵
                }
                //不存在：主檔添加。
                else if (list.Count == 0)
                {
                    //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    //處理所需的值。

                    // 查詢出主鍵的最大值、最後的來料單號
                    string sqlStr2 = "USE [sqlstock_vn01] SELECT TOP 1 productoutmain_id , diff_voucherno FROM productoutmain ORDER BY productoutmain_id DESC";

                    List<string> list2 = new List<string>();
                    list2.AddRange(sqltool.SqlSelect(sqlStr2));

                    string searchdate = ""; //出貨單號，去除最後流水3碼
                    string diff_voucherno = ""; //出貨單號
                    string today = DateTime.Now.ToString("yyyyMMdd");//今日日期
                    if (list2.Count == 0)//沒有資料
                    {
                        productoutmain_id = "1";
                        diff_voucherno = "F" + today + "001";
                        searchdate = "F" + today;
                    }
                    else
                    {
                        productoutmain_id = (Convert.ToInt32(list2[0].ToString()) + 1).ToString();// 出貨主鍵

                        string lastDate = GetDate(list2[1].ToString());//最後一筆資料的日期
                        // 今日的001出貨單號，已經產生。
                        if (lastDate.Equals(today))
                        {
                            int num = Convert.ToInt32(list2[1].Substring(9)) + 1;

                            diff_voucherno = "F" + lastDate + num.ToString("D3");//不足三位數時，前面的位數以0補足。
                            searchdate = "F" + lastDate;
                        }
                        // 今日的出貨單號，還未產生。
                        else
                        {
                            diff_voucherno = "F" + today + "001";
                            searchdate = "F" + today;
                        }
                    }

                    string outgoods_date = HandleDate(s.date); //出貨日期

                    //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    // 確保不發生，主檔有新增，但明細檔沒有新增。
                    if (s.product_id != null)
                    {
                        //主檔的新增
                        string sqlStr3 = "USE [sqlstock_vn01] INSERT INTO productoutmain" +
                                        "(" +
                                        "productoutmain_id" +
                                        ",outgoods_date" +
                                        ",diff_voucherno" +
                                        ",custom_id" +
                                        ",custno" +
                                        ",cust_nickname" +
                                        ",promoterno" +
                                        ",promotername" +
                                        ",tax_kind" +
                                        ",tax_rate" +
                                        ",isCloseSheet" +
                                        ",memodescr" +
                                        ",warehouse_id" +
                                        ",warehouseno" +
                                        ",warehousename" +
                                        ",diff_classify" +
                                        ",transfercash_key" +
                                        ",tran_summons" +
                                        ",tran_summons_key" +
                                        ",search_date" +
                                        ",total_dis_rate" +
                                        ",bef_discount_total" +
                                        ",isorderin" +
                                        ",ispurch" +
                                        ",WAREMODE" +
                                        ",ATTFILES" +
                                        ",oem_pro_voucherno" +
                                        ")" +
                                        "VALUES" +
                                        "(" + productoutmain_id +
                                        ", '" + outgoods_date + "' " +
                                        ", '" + diff_voucherno + "' " +
                                        ", " + " 244 " +
                                        ", " + " '001' " +
                                        ", " + " 'VISION' " +
                                        ", " + " 'A0-001' " +
                                        ", " + " 'auto' " +
                                        ", " + " 1 " +
                                        ", " + " 0 " +
                                        ", " + " 0 " +
                                        ", " + " '自動匯入' " +
                                        ", " + " 74 " +
                                        ", " + " 'P' " +
                                        ", " + " '成品倉' " +
                                        ", " + " 1 " +
                                        ", " + " 1 " +
                                        ", " + " 1 " +
                                        ", " + " 1 " +
                                        ", '" + searchdate + "' " +
                                        ", " + " 1 " +
                                        ", " + " 0 " +
                                        ", " + " 0 " +
                                        ", " + " 0 " +
                                        ", " + " 0 " +
                                        ", " + " '' " +
                                        ", '" + s.proID + "' " +
                                        ")";

                        int result = sqltool.SqlInsert(sqlStr3);
                        if (result == 0)//失敗
                        {
                            //紀錄錯誤
                            errorCatch = errorCatch + "錯誤";
                            //避免明細表增加時，有值可以新增。
                            productoutmain_id = null;
                            //MessageBox.Show("新增「出貨單」，有錯誤");
                        }
                    }
                }

                // 確保主檔沒新增時，明細表也不會新增。
                if (productoutmain_id != null)
                {
                    //明細檔的新增
                    string sqlStr4 = "USE [sqlstock_vn01] INSERT INTO productoutdetail " +
                                    "(" +
                                    "productoutmain_id" +
                                    ",product_id" +
                                    ",prdno" +
                                    ",prdno_ser" +
                                    ",prdname" +
                                    ",unit" +
                                    ",qty" +
                                    ",warehouse_id" +
                                    ",memodescr" +
                                    ",inventory_style_id" +
                                    ",adflag" +
                                    ",oem_voucherno" +
                                    ")" +
                                    "VALUES" +
                                    "(" + productoutmain_id +
                                    ", " + s.product_id +
                                    ", '" + s.prdID + "' " +
                                    ", '" + s.prdSer + "' " +
                                    ", '" + s.prdName + "' " +
                                    ", " + " 'pcs' " +
                                    ", " + " 12 " +
                                    ", " + " 74 " +
                                    ", " + " '自動匯入' " +
                                    ", " + " 182 " +
                                    ", " + " 2 " +
                                    ", '" + s.oemID + "' " +
                                    ")";

                    int result2 = sqltool.SqlInsert(sqlStr4);
                    if (result2 == 0)//失敗
                    {
                        //紀錄錯誤
                        errorCatch = errorCatch + "錯誤";
                        //MessageBox.Show("新增「出貨單」，有錯誤");
                    }
                }
            }
            return Tuple.Create(prolist, errorCatch);
        }

    }
}
