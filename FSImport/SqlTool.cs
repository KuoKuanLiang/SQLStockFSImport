using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;


using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;

using System.Windows.Forms;
//using Dapper;
//using MySql.Data.MySqlClient;

namespace FSImport
{
    public class SqlTool
    {
        SqlConnection conn = new SqlConnection();

        SqlCommand cmd = new SqlCommand();

        SqlTransaction transaction = null;

        //SqlDataReader reader = null;

        public void Sqlstart()
        {
            try
            {
                //資料庫連接字串
                string connectionString = "Server=192.168.1.76;Database=sqlstock_vn01;User ID=sa;Password=123"; //"Server=192.168.0.9;Database=sqlstock_v01;User ID=sa;Password=123"
                //設定連接字串                                                                                                
                conn.ConnectionString = connectionString;
                //開啟連接(避免重複開啟，先註解)
                conn.Open();

                cmd = conn.CreateCommand();
                transaction = conn.BeginTransaction();

                //將transaction分配給sqlcommand物件
                cmd.Transaction = transaction;
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料庫連線錯誤：" + ex);
            }
        }

        /// <summary>
        /// 開啟資料庫連線
        /// </summary>
        /// <returns></returns>
        public SqlConnection SqlOpen()
        {
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            return conn;
        }

        /// <summary>
        /// 關閉資料庫連線
        /// </summary>
        public SqlConnection SqlClose()
        {
            //確認是開啟的狀態
            if (conn.State == ConnectionState.Open)
            {
                //關閉資料庫連接
                conn.Close();
            }
            return conn;
        }

        /// <summary>
        /// 認可資料庫交易
        /// </summary>
        public void TraCommit()
        {
            try
            {
                transaction.Commit();
                //(就算insert、select中，有異常的，還是會確認交易)
            }
            catch (Exception ex)
            {
                Console.WriteLine("TraCommit，錯誤訊息：" + ex);
                TraRollback();
            }
        }

        /// <summary>
        /// 回復資料庫交易，改成變動前。
        /// </summary>
        public void TraRollback()
        {
            try
            {
                //回復交易，執行這行時，會回復在交易內所有SQL所更動的內容。
                transaction.Rollback();
            }
            catch (Exception ex)
            {
                Console.WriteLine("TraRollback，錯誤訊息：" + ex);
            }
        }


        /// <summary>
        /// 特殊查詢，傳回結果
        /// </summary>
        /// <param name="strFrom"></param>
        /// <param name="strWhere"></param>
        /// <returns></returns>
        public int SqlJudge(string sqlStr)
        {
            int result = 0;
            //string sqlStr = "SELECT top 1 1 FROM " + strFrom + " WHERE " + strWhere;
            try
            {
                //確保是連接狀態
                SqlOpen();
                cmd.CommandText = sqlStr;
                SqlDataReader reader = cmd.ExecuteReader();
                //如果有回傳值
                if (reader.HasRows)
                {
                    result = 1;
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SqlJudge，錯誤訊息：" + ex);
            }
            return result;
        }


        /// <summary>
        /// 查詢，傳回結果
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public List<string> SqlSelect(string sqlStr)
        {
            List<string> result = new List<string>();
            try
            {
                //確保是連接狀態
                SqlOpen();
                cmd.CommandText = sqlStr;
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())//橫列
                    {
                        for (int j = 0; j < reader.FieldCount; j++)//直欄
                        {
                            result.Add(reader[j].ToString());
                            //Console.WriteLine(dataReader[j].ToString());
                        }
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SqlSelect，錯誤訊息： " + ex);
            }

            return result;
        }

        /// <summary>
        /// 執行新增，並回傳受影響列數
        /// </summary>
        /// <param name="sqlStr"></param>
        /// <returns></returns>
        public int SqlInsert(string sqlStr)
        {
            int result = 0;
            try 
            {
                //確保是連接狀態
                SqlOpen();
                cmd.CommandText = sqlStr;
                result = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                
                Console.WriteLine("SqlInsert，錯誤訊息：" + ex);
            }
            return result;
        }

    }
}
