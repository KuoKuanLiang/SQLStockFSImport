/****** ��custominmaterialmain�Bcustominmaterialdetail�Bproductoutmain�Bproductoutdetail��ƪ�A�s�WSQLStockFSImport�۰ʶפJ�t�η|�ϥΨ쪺�s���  ******/
USE [sqlstock_vn01]
GO

alter table custominmaterialmain add oem_cus_voucherno char(15);
alter table custominmaterialdetail add oem_voucherno char(18);
alter table productoutmain add oem_pro_voucherno char(15);
alter table productoutdetail add oem_voucherno char(18);