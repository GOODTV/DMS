USE [VDWDOM]
GO
/****** Object:  StoredProcedure [dbo].[getDonateRecord]    Script Date: 06/13/2013 17:35:07 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** 物件:  預存程序 dbo.getDonateRecord    指令碼日期: 2013/5/7 下午 15:56:38 ******/
/***********************************************************************************************************
建立者：Tanya Wu
建立日期：2013/05/07
功能說明：取得線上徵信錄明細
傳入參數： 
    @NAME	奉獻姓名/團體
	@YYYY 	奉獻年度
	@MM 	奉獻月份
傳出結果：Table 徵信錄明細	
修改紀錄：
Date      Content
-----------------------------------------------------------------------------------------------------------
***********************************************************************************************************/
ALTER PROCEDURE [dbo].[getDonateRecord]
(@NAME VARCHAR(40),
 @YYYY VARCHAR(4),
 @MM   VARCHAR(2))
AS

	SELECT
		  R.REVNO as [收據號碼]
	      ,P.NAME as [奉獻姓名/團體]
	      ,Replace(Convert(Varchar(12),CONVERT(money,R.REVAMT),1),'.00','') as [金額]
	      ,SUBSTRING(R.REVDT,1,4)+'/'+SUBSTRING(R.REVDT,5,2)+'/'+SUBSTRING(R.REVDT,7,2) as [日期]
    FROM RECE R
    LEFT JOIN PERSON P
    ON R.PID = P.PID
    WHERE P.NEAT = 0 AND R.REVAMT > 0 AND R.REVAMT < 500000
      AND R.REVDT LIKE @YYYY+@MM+'%' 	
      AND P.NAME LIKE '%'+@NAME+'%'
      AND P.NAME NOT LIKE '%主知名%'
      AND P.NAME NOT LIKE '%有志者%'
      AND P.NAME NOT LIKE '%無名氏%'
    ORDER BY R.REVDT
