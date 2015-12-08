USE [DonationGoodTV]
GO

/****** Object:  StoredProcedure [dbo].[getDonateRecord]    Script Date: 11/28/2013 15:01:33 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


/****** 物件:  預存程序 dbo.getDonateRecord    指令碼日期: 2013/5/7 下午 15:56:38 ******/
/***********************************************************************************************************
建立者：Tanya Wu
建立日期：2013/09/02
功能說明：取得線上徵信錄明細
傳入參數： 
    @NAME	奉獻姓名/團體(捐款人or收據抬頭)
	@YYYY 	奉獻年度
	@MM 	        奉獻月份
傳出結果：Table 徵信錄明細	
修改紀錄：
Date      Content
-----------------------------------------------------------------------------------------------------------
***********************************************************************************************************/
CREATE PROCEDURE [dbo].[getDonateRecord]
(@NAME NVARCHAR(100),
 @YYYY VARCHAR(4),
 @MM   VARCHAR(2))
AS

	SELECT
		  M.Invoice_No as [收據號碼]
	      ,M.Invoice_Title as [奉獻姓名/團體]
	      ,Replace(Convert(Varchar(12),M.Donate_Amt,1),'.00','') as [金額]
	      ,CONVERT(varchar(100), M.Donate_Date, 111) as [日期]
    FROM DONATE M
    LEFT JOIN DONOR D
    ON M.Donor_Id = D.Donor_Id
    WHERE D.Report_Type = '刊登' AND M.Donate_Amt > 0 AND M.Donate_Amt < 500000
      AND DATEPART(Year,M.Donate_Date) = @YYYY
      AND DATEPART(Month,M.Donate_Date) = @MM
      AND M.Invoice_Title LIKE '%'+@NAME+'%'
      AND M.Invoice_Title NOT LIKE '%主知名%' AND D.Donor_Name NOT LIKE '%主知名%'
      AND M.Invoice_Title NOT LIKE '%有志者%' AND D.Donor_Name NOT LIKE '%有志者%'
      AND M.Invoice_Title NOT LIKE '%無名氏%' AND D.Donor_Name NOT LIKE '%無名氏%'
    ORDER BY M.Donate_Date


GO


