USE [DonationGoodTV]
GO

/****** Object:  StoredProcedure [dbo].[getDonateRecord]    Script Date: 11/28/2013 15:01:33 ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER OFF
GO


/****** ����:  �w�s�{�� dbo.getDonateRecord    ���O�X���: 2013/5/7 �U�� 15:56:38 ******/
/***********************************************************************************************************
�إߪ̡GTanya Wu
�إߤ���G2013/09/02
�\�໡���G���o�u�W�x�H������
�ǤJ�ѼơG 
    @NAME	�^�m�m�W/����(���ڤHor���ک��Y)
	@YYYY 	�^�m�~��
	@MM 	        �^�m���
�ǥX���G�GTable �x�H������	
�ק�����G
Date      Content
-----------------------------------------------------------------------------------------------------------
***********************************************************************************************************/
CREATE PROCEDURE [dbo].[getDonateRecord]
(@NAME NVARCHAR(100),
 @YYYY VARCHAR(4),
 @MM   VARCHAR(2))
AS

	SELECT
		  M.Invoice_No as [���ڸ��X]
	      ,M.Invoice_Title as [�^�m�m�W/����]
	      ,Replace(Convert(Varchar(12),M.Donate_Amt,1),'.00','') as [���B]
	      ,CONVERT(varchar(100), M.Donate_Date, 111) as [���]
    FROM DONATE M
    LEFT JOIN DONOR D
    ON M.Donor_Id = D.Donor_Id
    WHERE D.Report_Type = '�Z�n' AND M.Donate_Amt > 0 AND M.Donate_Amt < 500000
      AND DATEPART(Year,M.Donate_Date) = @YYYY
      AND DATEPART(Month,M.Donate_Date) = @MM
      AND M.Invoice_Title LIKE '%'+@NAME+'%'
      AND M.Invoice_Title NOT LIKE '%�D���W%' AND D.Donor_Name NOT LIKE '%�D���W%'
      AND M.Invoice_Title NOT LIKE '%���Ӫ�%' AND D.Donor_Name NOT LIKE '%���Ӫ�%'
      AND M.Invoice_Title NOT LIKE '%�L�W��%' AND D.Donor_Name NOT LIKE '%�L�W��%'
    ORDER BY M.Donate_Date


GO


