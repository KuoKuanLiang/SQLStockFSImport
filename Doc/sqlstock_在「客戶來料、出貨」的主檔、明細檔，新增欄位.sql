/****** 對custominmaterialmain、custominmaterialdetail、productoutmain、productoutdetail資料表，新增SQLStockFSImport自動匯入系統會使用到的新欄位  ******/
USE [sqlstock_vn01]
GO

alter table custominmaterialmain add oem_cus_voucherno char(15);
alter table custominmaterialdetail add oem_voucherno char(18);
alter table productoutmain add oem_pro_voucherno char(15);
alter table productoutdetail add oem_voucherno char(18);