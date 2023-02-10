using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FSImport
{
    public partial class Form1 : Form
    {
        SqlTool sqltool = new SqlTool();

        Func func = new Func();

        List<PrdData> prdlist = new List<PrdData>();
        List<CusData> cuslist = new List<CusData>();
        List<ProData> prolist = new List<ProData>();

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Form1開啟時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            // 資料庫連接開啟
            sqltool.Sqlstart();
            //Console.WriteLine("123");

            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }

        /// <summary>
        /// Form1關閉時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
            //messageBoxCS.AppendFormat("{0} = {1}", "CloseReason", e.CloseReason);
            //messageBoxCS.AppendLine();
            //messageBoxCS.AppendFormat("{0} = {1}", "Cancel", e.Cancel);
            //messageBoxCS.AppendLine();
            //MessageBox.Show(messageBoxCS.ToString(), "FormClosing Event");

            //交易回滾
            sqltool.TraRollback();

            // 資料庫連接開閉
            sqltool.SqlClose();
        }

        /// <summary>
        /// 選擇excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            // 按鈕不可使用
            button1.Enabled = false;//確保不會再跳出一個選擇視窗

            //開啟選擇視窗，回傳取得檔案的路徑+名稱
            List<string> list = func.GetFile();

            //將回傳值填到表格中
            foreach (string s in list)
            {
                int num = 0;
                //表格大於1筆資料的時候，再做比對判斷。
                if (dataGridView1.Rows.Count != 0)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        //判斷有沒有重複的檔案
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString() == s)
                        {
                            //有重複，不取得
                            num = 1;
                            break;
                        }
                    }
                }
                if(num == 0)
                {
                    //新增一空白列
                    dataGridView1.Rows.Add();

                    //取得共多少列
                    int rowCount = dataGridView1.Rows.Count;

                    //要填入列的內容
                    DataGridViewRow rs = dataGridView1.Rows[rowCount - 1];

                    //填入資料
                    rs.Cells[0].Value = s;
                }
            }

            // 按鈕可使用
            button1.Enabled = true;//繼續選擇excel
            button2.Enabled = true;//開始進行檢查與預覽
        }


        /// <summary>
        /// 檢查並預覽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // 按鈕不可使用
            // 確保執行期間，不會因按其他按鈕而有問題。
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;

            //陣列清空，避免重複按的時候，會使用到之前的資料。
            prdlist.Clear(); //料品
            cuslist.Clear(); //客戶來料
            prolist.Clear(); //出貨

            //表格清空，避免重複按的時候，會使用到之前的資料。
            dataGridView2.Rows.Clear(); //料品
            dataGridView3.Rows.Clear(); //客戶來料
            dataGridView4.Rows.Clear(); //出貨

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：開啟excel，取得資料，分配到三個list");

            // 取得所有excel檔案的內容，並分配在三個list中。
            var returnList = func.CreatDataList(dataGridView1);
            prdlist = returnList.Item1.ToList(); //料品
            cuslist = returnList.Item2.ToList(); //客戶來料
            prolist = returnList.Item3.ToList(); //出貨

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：開啟excel，取得資料，分配到三個list");

            //--------------------------------------------------------------

            //判斷「狀態」應為何?(重複、新、異常)

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：判斷「狀態」應為何?(重複、新、異常)");

            prdlist = func.JudgePrdList(prdlist,sqltool).ToList(); //料品
            cuslist = func.JudgeCusList(cuslist, sqltool).ToList(); //客戶來料
            prolist = func.JudgeProList(prolist, sqltool).ToList(); //出貨

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：判斷「狀態」應為何?(重複、新、異常)");

            //--------------------------------------------------------------

            //在表格顯示資料

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：在表格顯示資料");

            //料品
            foreach (PrdData s in prdlist)
            {
                dataGridView2.Rows.Add();//新增一空白列

                int rowCount = dataGridView2.Rows.Count;//取得共多少列

                DataGridViewRow rs = dataGridView2.Rows[rowCount - 1];//要填入列的內容

                rs.Cells[0].Value = s.state;//狀態
                rs.Cells[1].Value = s.prdID;//品牌代號
                rs.Cells[2].Value = s.prdName;//品牌名稱
                rs.Cells[3].Value = s.prdSer;//型號
            }

            //客戶來料
            foreach (CusData s in cuslist)
            {
                dataGridView3.Rows.Add();//新增一空白列

                int rowCount = dataGridView3.Rows.Count;//取得共多少列

                DataGridViewRow rs = dataGridView3.Rows[rowCount - 1];//要填入列的內容

                rs.Cells[0].Value = s.state;//狀態
                rs.Cells[1].Value = s.cusID;//外包出貨單號
                rs.Cells[2].Value = s.oemID;//代工回修單號
                rs.Cells[3].Value = s.prdID;//品牌代號
                rs.Cells[4].Value = s.prdName;//品牌名稱
                rs.Cells[5].Value = s.prdSer;//型號
                rs.Cells[6].Value = s.date;//點收日期
                rs.Cells[7].Value = s.qty;//數量
            }

            //出貨
            foreach (ProData s in prolist)
            {
                dataGridView4.Rows.Add();//新增一空白列

                int rowCount = dataGridView4.Rows.Count;//取得共多少列

                DataGridViewRow rs = dataGridView4.Rows[rowCount - 1];//要填入列的內容

                rs.Cells[0].Value = s.state;//狀態
                rs.Cells[1].Value = s.proID;//回廠出貨單號
                rs.Cells[2].Value = s.oemID;//代工回修單號
                rs.Cells[3].Value = s.prdID;//品牌代號
                rs.Cells[4].Value = s.prdName;//品牌名稱
                rs.Cells[5].Value = s.prdSer;//型號
                rs.Cells[6].Value = s.date;//點收日期
                rs.Cells[7].Value = s.qty;//數量
            }
            // 可能以上要使用委派

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：在表格顯示資料");

            //--------------------------------------------------------------

            //確保沒有異常資料，才能使用匯入按鈕。
            bool prdBool = prdlist.Exists(x => x.state == "異常");
            bool cusBool = cuslist.Exists(x => x.state == "異常");
            bool proBool = prolist.Exists(x => x.state == "異常");
            if( prdBool == false && cusBool == false && proBool == false)
            {
                button4.Enabled = true;
            }

            // 可使用按鈕
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
        }

        /// <summary>
        /// 將預覽內容，匯出到excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            //建立excel，並將預覽內容資料放入。
            func.ExportExcelData(prdlist,cuslist,prolist);
        }

        /// <summary>
        /// 將預覽內容，匯入到sqlstock使用的資料庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            // 匯入完，應該要清空列表資料，並且解除鎖表。

            // 沒有顯示異常，才能添加。
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;

            //-------------------------------------------------------------------------

            //開始進行匯入

            //錯誤紀錄
            string errorCatch = "";

            //刪除狀態是「重複」的資料。(料品不用刪除)
            cuslist.RemoveAll(x => x.state == "重複"); //客戶來料
            prolist.RemoveAll(x => x.state == "重複"); //出貨

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：料品資料匯入交易");

            //料品
            var returnPrd = func.InsertPrdList(prdlist, errorCatch, sqltool);
            prdlist = returnPrd.Item1.ToList(); //料品
            errorCatch = returnPrd.Item2.ToString(); //錯誤紀錄

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：料品資料匯入交易");

            // 使客戶來料、出貨，取得有使用的料品主鍵id。
            foreach (CusData s in cuslist)
            {
                s.product_id = prdlist.Find(x => x.prdID == s.prdID && x.prdName == s.prdName && x.prdSer == s.prdSer).product_id;
            }
            foreach (ProData s in prolist)
            {
                s.product_id = prdlist.Find(x => x.prdID == s.prdID && x.prdName == s.prdName && x.prdSer == s.prdSer).product_id;
            }

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：客戶來料資料匯入交易");

            //客戶來料
            var returnCus = func.InsertCusList(cuslist, errorCatch, sqltool);
            cuslist = returnCus.Item1.ToList(); //料品
            errorCatch = returnCus.Item2.ToString(); //錯誤紀錄

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：客戶來料資料匯入交易");

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：出貨資料匯入交易");

            //出貨
            var returnPro = func.InsertProList(prolist, errorCatch, sqltool);
            prolist = returnPro.Item1.ToList(); //料品
            errorCatch = returnPro.Item2.ToString(); //錯誤紀錄

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：出貨資料匯入交易");

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 開始：確認或回滾交易");

            string lastState = "";
            //Insert沒有錯誤，確認交易
            if (errorCatch == "")
            {
                sqltool.TraCommit();
                lastState = "完成";
            }
            //Insert有錯誤，回滾交易
            else
            {
                sqltool.TraRollback();
                lastState = "失敗";
            }

            Console.WriteLine(DateTime.Now.ToString("HH-mm-ss-fff") + " 結束：確認或回滾交易");

            // 顯示此次匯入結果
            for (int i = 0; i < dataGridView2.Rows.Count; i++) //料品
            {
                dataGridView2.Rows[i].Cells[0].Value = lastState;
            }
            for (int i = 0; i < dataGridView3.Rows.Count; i++) //客戶
            {
                dataGridView3.Rows[i].Cells[0].Value = lastState;
            }
            for (int i = 0; i < dataGridView4.Rows.Count; i++) //出貨
            {
                dataGridView4.Rows[i].Cells[0].Value = lastState;
            }

            //-------------------------------------------------------------------------

            //斷開連接
            sqltool.SqlClose();

            //重新連接，主要是為了要重新建立SqlTransaction
            sqltool.Sqlstart();

            //可以選擇excel
            button1.Enabled = true;

            //陣列清空，避免重複按的時候，會使用到之前的資料。
            prdlist.Clear();
            cuslist.Clear();
            prolist.Clear();

            //表格清空，讓人員可以繼續選擇excel檔
            dataGridView1.Rows.Clear();

        }

    }
}
