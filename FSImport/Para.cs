using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FSImport
{
    /// <summary>
    /// 從excel取得的資料
    /// </summary>
    public class XlsData : IEquatable<XlsData> 
    {
        /// <summary>
        /// 外包出貨單號
        /// </summary>
        public string cusID;
        /// <summary>
        /// 品牌代號
        /// </summary>
        public string prdID;
        /// <summary>
        /// 品牌名稱
        /// </summary>
        public string prdName;
        /// <summary>
        /// 型號
        /// </summary>
        public string prdSer;
        /// <summary>
        /// 廠商點收日期
        /// </summary>
        public string cusDate;
        /// <summary>
        /// 數量
        /// </summary>
        public string qty;
        /// <summary>
        /// 代工回修單號
        /// </summary>
        public string oemID;
        /// <summary>
        /// 回廠出貨單號
        /// </summary>
        public string proID;

        //---------------------------------------------------------------

        //檢查有沒有重複項

        public bool Equals(XlsData other)
        {
            //檢查比較對像是否為null。
            if (Object.ReferenceEquals(other, null)) return false;

            //檢查比較對像是否引用了相同的數據。
            if (Object.ReferenceEquals(this, other)) return true;

            //檢查XlsData的屬性是否相等。
            return cusID.Equals(other.cusID) &&
                   prdID.Equals(other.prdID) &&
                   prdName.Equals(other.prdName) &&
                   prdSer.Equals(other.prdSer) &&
                   cusDate.Equals(other.cusDate) &&
                   qty.Equals(other.qty) &&
                   oemID.Equals(other.oemID) &&
                   proID.Equals(other.proID);
        }

        // 如果 Equals() 對一對對象返回 true
        // 那麼 GetHashCode() 必須為這些對象返回相同的值。

        public override int GetHashCode()
        {
            //獲取對應屬性的哈希碼。
            int hashcusID = cusID.GetHashCode();
            int hashprdID = prdID.GetHashCode();
            int hashprdName = prdName.GetHashCode();
            int hashprdSer = prdSer.GetHashCode();
            int hashcusDate = cusDate.GetHashCode();
            int hashqty = qty.GetHashCode();
            int hashoemID = oemID.GetHashCode();
            int hashproID = proID.GetHashCode();

            //計算XlsData的哈希碼。
            return hashcusID ^ hashprdID ^ hashprdName ^ hashprdSer ^ hashcusDate ^ hashqty ^ hashoemID ^ hashproID;
        }

        //---------------------------------------------------------------
    }

    /// <summary>
    /// 放入客戶來料單的資料
    /// </summary>
    public class CusData : IEquatable<CusData>
    {
        /// <summary>
        /// 狀態
        /// </summary>
        public string state;
        /// <summary>
        /// 外包出貨單號
        /// </summary>
        public string cusID;
        /// <summary>
        /// 代工回修單號
        /// </summary>
        public string oemID;
        /// <summary>
        /// 品牌代號
        /// </summary>
        public string prdID;
        /// <summary>
        /// 品牌名稱
        /// </summary>
        public string prdName;
        /// <summary>
        /// 型號
        /// </summary>
        public string prdSer;
        /// <summary>
        /// 點收日期
        /// </summary>
        public string date;
        /// <summary>
        /// 數量
        /// </summary>
        public string qty;
        ///// <summary>
        ///// 客戶來料主檔主鍵
        ///// </summary>
        //public string cusMain_id;
        /// <summary>
        /// 料品主鍵
        /// </summary>
        public string product_id;

        //---------------------------------------------------------------

        //檢查有沒有重複項

        public bool Equals(CusData other)
        {
            //檢查比較對像是否為null。
            if (Object.ReferenceEquals(other, null)) return false;

            //檢查比較對像是否引用了相同的數據。
            if (Object.ReferenceEquals(this, other)) return true;

            //檢查CusData的屬性是否相等。
            return cusID.Equals(other.cusID) &&
                   oemID.Equals(other.oemID) &&
                   prdID.Equals(other.prdID) &&
                   prdName.Equals(other.prdName) &&
                   prdSer.Equals(other.prdSer) &&
                   date.Equals(other.date) &&
                   qty.Equals(other.qty);
        }

        // 如果 Equals() 對一對對象返回 true
        // 那麼 GetHashCode() 必須為這些對象返回相同的值。

        public override int GetHashCode()
        {
            //獲取對應屬性的哈希碼。
            int hashcusID = cusID.GetHashCode();
            int hashoemID = oemID.GetHashCode();
            int hashprdID = prdID.GetHashCode();
            int hashprdName = prdName.GetHashCode();
            int hashprdSer = prdSer.GetHashCode();
            int hashdate = date.GetHashCode();
            int hashqty = qty.GetHashCode();

            //計算CusData的哈希碼。
            return hashcusID ^ hashoemID ^ hashprdID ^ hashprdName ^ hashprdSer ^ hashdate ^ hashqty ;
        }

        //---------------------------------------------------------------
    }

    /// <summary>
    /// 放入出貨單的資料
    /// </summary>
    public class ProData : IEquatable<ProData>
    {
        /// <summary>
        /// 狀態
        /// </summary>
        public string state;
        /// <summary>
        /// 回廠出貨單號
        /// </summary>
        public string proID;
        /// <summary>
        /// 代工回修單號
        /// </summary>
        public string oemID;
        /// <summary>
        /// 品牌代號
        /// </summary>
        public string prdID;
        /// <summary>
        /// 品牌名稱
        /// </summary>
        public string prdName;
        /// <summary>
        /// 型號
        /// </summary>
        public string prdSer;
        /// <summary>
        /// 出貨日期
        /// </summary>
        public string date;
        /// <summary>
        /// 數量
        /// </summary>
        public string qty;
        ///// <summary>
        ///// 出貨主檔主鍵
        ///// </summary>
        //public string proMain_id;
        /// <summary>
        /// 料品主鍵
        /// </summary>
        public string product_id;

        //---------------------------------------------------------------

        //檢查有沒有重複項

        public bool Equals(ProData other)
        {
            //檢查比較對像是否為null。
            if (Object.ReferenceEquals(other, null)) return false;

            //檢查比較對像是否引用了相同的數據。
            if (Object.ReferenceEquals(this, other)) return true;

            //檢查ProData的屬性是否相等。
            return proID.Equals(other.proID) &&
                   oemID.Equals(other.oemID) &&
                   prdID.Equals(other.prdID) &&
                   prdName.Equals(other.prdName) &&
                   prdSer.Equals(other.prdSer) &&
                   date.Equals(other.date) &&
                   qty.Equals(other.qty);
        }

        // 如果 Equals() 對一對對象返回 true
        // 那麼 GetHashCode() 必須為這些對象返回相同的值。

        public override int GetHashCode()
        {
            //獲取對應屬性的哈希碼。
            int hashproID = proID.GetHashCode();
            int hashoemID = oemID.GetHashCode();
            int hashprdID = prdID.GetHashCode();
            int hashprdName = prdName.GetHashCode();
            int hashprdSer = prdSer.GetHashCode();
            int hashdate = date.GetHashCode();
            int hashqty = qty.GetHashCode();

            //計算ProData的哈希碼。
            return hashproID ^ hashoemID ^ hashprdID ^ hashprdName ^ hashprdSer ^ hashdate ^ hashqty;
        }

        //---------------------------------------------------------------
    }

    /// <summary>
    /// 放入料號基本資料的資料
    /// </summary>
    public class PrdData : IEquatable<PrdData>
    {
        /// <summary>
        /// 狀態
        /// </summary>
        public string state;
        /// <summary>
        /// 品牌代號
        /// </summary>
        public string prdID;
        /// <summary>
        /// 品牌名稱
        /// </summary>
        public string prdName;
        /// <summary>
        /// 型號
        /// </summary>
        public string prdSer;
        /// <summary>
        /// 料品主鍵
        /// </summary>
        public string product_id;
        /// <summary>
        /// 料號(prdno)對應id
        /// </summary>
        public string prdno_id;

        //---------------------------------------------------------------

        //檢查有沒有重複項

        public bool Equals(PrdData other)
        {
            //檢查比較對像是否為null。
            if (Object.ReferenceEquals(other, null)) return false;

            //檢查比較對像是否引用了相同的數據。
            if (Object.ReferenceEquals(this, other)) return true;

            //檢查PrdData的屬性是否相等。
            return prdID.Equals(other.prdID) &&
                   prdName.Equals(other.prdName) &&
                   prdSer.Equals(other.prdSer) ;
        }

        // 如果 Equals() 對一對對象返回 true
        // 那麼 GetHashCode() 必須為這些對象返回相同的值。

        public override int GetHashCode()
        {
            //獲取對應屬性的哈希碼。
            int hashprdID = prdID.GetHashCode();
            int hashprdName = prdName.GetHashCode();
            int hashprdSer = prdSer.GetHashCode();

            //計算PrdData的哈希碼。
            return hashprdID ^ hashprdName ^ hashprdSer;
        }

        //---------------------------------------------------------------
    }
}
