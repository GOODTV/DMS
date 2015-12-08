using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Threading;

public partial class Ajax : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string Type = Request.Form["Type"];
        string result = "";
        switch (Type)
        {
            case "1":  //以 CityID 取得其鄉鎮列表
                result = GetAreaList();
                break;
            case "2":  //同個案戶籍地址
                result = GetHouseholdAddress();
                break;
            case "3":  //新聞
                result = GetNewsDetail();
                break;
            case "4":  //檢核Email是否存在
                result = EmailIsExist();
                break;
            case "5":  //檢查選取捐款人統計中哪個報表
                DonorStat_Qry();
                break;
            case "6":  //檢查選取會員統計中哪個報表
                MemberStat_Qry();
                break;
            case "7":  //每月固定捐款後重整lbl
                result = LoadPledgeData();
                break;
            case "8":  //檢查相同通訊地址
                result = GetAddress();
                break;
            case "9":  //檢查相同收據地址
                result = GetInvoiceAddress();
                break;
            case "10":  //檢核匯入台銀回覆檔檔名是否存在
                result = PledgeBatchFileNameIsExist();
                break;
            case "11":  //登入Email帳號取得捐款人編號
                result = GetDonorIdFromEmail();
                break;
            case "12":  //取得群組類別資料
                result = GetGroupClass();
                break;
            case "13":  //取得群組代表資料
                result = GetGroupItem();
                break;
            case "14":  //儲存捐款人的群組代表資料
                result = SaveGroupItem();
                break;
            case "15":  //刪除捐款人的群組代表資料
                result = DeleteGroupItem();
                break;
            case "16":  //取得群組代表詳細資料
                result = GetGroupItemDetail();
                break;
            case "17":  //儲存捐款人的群組代表的成員資料
                if (SaveGroupItem() == "Y")
                {
                    result = GetGroupItemMember();
                }
                else
                {
                    result = "";
                }
                break;
            case "18":  //取得捐款人的群組代表的成員資料
                result = GetGroupItemMember();
                break;
            case "19":  //查詢群組代表的相關資料
                result = GetGroupItemData();
                break;
            case "20":  //修改IEPay的付款方式
                result = UpdateIEPaySetPaytype();
                break;
            case "21":  //檢核捐款人收據寄送方式是否已存在
                result = InvoiceTypeIsExist();
                break;
        }
        Response.Write(result);
    }

    //修改IEPay的付款方式
    private string UpdateIEPaySetPaytype()
    {
        string strRet = "";
        string strOrderid = Request.Form["orderid"];
        string strPaytype = Request.Form["paytype"];

        string strSql = " update DONATE_IEPAY set paytype = @paytype where orderid = @orderid ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("orderid", strOrderid);
        dict.Add("paytype", strPaytype);
        int cnt = NpoDB.ExecuteSQLS(strSql, dict);

        if (cnt > 0)
        {
            strRet = "Y";
        }
        return strRet;
    }

    //查詢群組代表的相關資料
    public string GetGroupItemData()
    {

        string strRet = "";
        string strSQLWhere = "";
        string strGroupClass = Request.Form["GroupClass"];
        string strGroupItemName = Request.Form["GroupItemName"];
        string strGroupItemUid = Request.Form["GroupItemUid"];
        string strSupplement = Request.Form["Supplement"];
        string strDonorID = Request.Form["DonorID"];

        string strSql = @"
                   select ROW_NUMBER() OVER(ORDER BY gi.uid) as [序號] ,gi.uid,
                   gc.GroupClassName as 群組類別,
                   gi.GroupItemName as 群組代表,
                   gi.Supplement as 備註
                   from GroupItem gi
                   left join GroupClass gc on gi.GroupClassUid=gc.uid
                   where gi.DeleteDate is null
                  ";
        
        if (strGroupClass != "")
        {
            strSql += "and gi.GroupClassUID=@GroupClassUID ";
            strSQLWhere += "群組類別：" + strGroupClass;
        }
        if (strGroupItemName.Trim() != "")
        {
            strSql += "and gi.GroupItemName like @GroupItemName ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "群組代表：" + strGroupItemName.Trim();
        }
        if (strSupplement.Trim() != "")
        {
            strSql += "and gi.Supplement like @Supplement ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "備註：" + strSupplement.Trim();
        }
        if (strDonorID != "")
        {
            strSql += "and gi.uid in (SELECT GroupItemUid FROM GroupMapping where DeleteDate is null and DonorID = @DonorID)  ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "捐款人編號：" + strDonorID;
        }
        if (strGroupItemUid != "")
        {
            strSql += "and gi.uid = @GroupItemUid ";
            strSQLWhere += (strSQLWhere == "" ? "" : "&nbsp;,&nbsp;") + "群組代表編號：" + strGroupItemUid;
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupClassUID", strGroupClass);
        dict.Add("GroupItemName", "%" + strGroupItemName.Trim() + "%");
        dict.Add("Supplement", "%" + strSupplement.Trim() + "%");
        dict.Add("DonorID", strDonorID);
        dict.Add("GroupItemUid", strGroupItemUid);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        int intRowsCount = dt.Rows.Count;

        if (intRowsCount > 0 && intRowsCount <= 100)
        {

            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap><SPAN>序號</SPAN></TH>
                            <TH noWrap><SPAN>群組類別</SPAN></TH>
                            <TH noWrap><SPAN>群組代表</SPAN></TH>
                            <TH noWrap><SPAN>備註</SPAN></TH>
                            <TH noWrap><SPAN>成員列表 [捐款人編號(捐款人名稱), ...]</SPAN></TH>
                        </TR>";

            foreach (DataRow dr in dt.Rows)
            {

                string strItemUid = dr["uid"].ToString().Trim();

                string strMember = "";
                string strSql2 = @"
					select Donor_Id,Donor_Name,Sex,Title,Donor_Type,Cellular_Phone,Tel_Office,Tel_Office_Loc,Tel_Office_Ext,Email
					,(Case When D.IsAbroad = 'N' then B.Name + C.Name + D.[Address] Else D.[Address] End) as [Address] 
					,Invoice_Type,Remark
                    from GroupMapping GM1 
                    join Donor D on GM1.DonorId = D.Donor_Id and D.DeleteDate is null
			        left Join CODECITY As B on D.City = B.ZipCode 
				    left Join CODECITY As C on D.Area = C.ZipCode
	    			where GM1.DeleteDate is null and GM1.GroupItemUid = @ItemUid
                    order by Donor_Id
                        ";

                Dictionary<string, object> dict2 = new Dictionary<string, object>();
                dict2.Add("ItemUid", strItemUid);
                DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);
                if (dt2.Rows.Count > 0)
                {

                    DataRow dr2;
                    strMember = "";
                    for (int j = 0; j < dt2.Rows.Count; j++)
                    {
                        dr2 = dt2.Rows[j];
                        string strMemberTitle = "";
                        string strTel = dr2["Tel_Office_Loc"].ToString().Trim() == "" ? "" : "(" + dr2["Tel_Office_Loc"].ToString().Trim() + ")";
                        strTel += dr2["Tel_Office"].ToString().Trim();
                        strTel += strTel + dr2["Tel_Office_Ext"].ToString().Trim() == "" ? "" : "#" + dr2["Tel_Office_Ext"].ToString().Trim();
                        strMemberTitle += "性別：" + dr2["Sex"].ToString().Trim() + "\n";
                        strMemberTitle += "稱謂：" + dr2["Title"].ToString().Trim() + "\n";
                        strMemberTitle += "身分別：" + dr2["Donor_Type"].ToString().Trim() + "\n";
                        strMemberTitle += "手機：" + dr2["Cellular_Phone"].ToString().Trim() + "\n";
                        strMemberTitle += "電話(日)：" + strTel + "\n";
                        strMemberTitle += "E-Mail：" + dr2["Email"].ToString().Trim() + "\n";
                        strMemberTitle += "通訊地址：" + dr2["Address"].ToString().Trim() + "\n";
                        strMemberTitle += "收據開立：" + dr2["Invoice_Type"].ToString().Trim() + "\n";
                        strMemberTitle += "捐款人備註：\n" + dr2["Remark"].ToString().Trim() + "\n";
                        strMember += "<SPAN onclick=\"window.event.cancelBubble=true;window.open('" +
                            Util.RedirectByTime("/DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=" + dr2["Donor_Id"].ToString().Trim()) +
                            "','_self','')\" title='" + strMemberTitle + "'><font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "</font>(" + dr2["Donor_Name"].ToString().Trim() + ")</SPAN>";


                    }

                }

                //strBody += "<TR style=\"cursor: pointer; cursor: hand;\" onclick =\"window.open('" + Util.RedirectByTime("GroupItem_Edit.aspx", "GroupItemUID=" + strItemUid) + "','_self','')\">";
                strBody += "<TR id='" + strItemUid + "'>";
                strBody += "<TD noWrap align='right'><SPAN>" + dr["序號"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["群組類別"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["群組代表"].ToString() + "</SPAN></TD>";
                strBody += "<TD><SPAN>" + dr["備註"].ToString() + "</SPAN></TD>";
                strBody += "<TD>" + strMember + "&nbsp;</TD></TR>";
            }

            strBody += "</table>";
            strRet = "<font color='blue'>【查詢條件】</font>" + (strSQLWhere == "" ? "全部" : strSQLWhere) + " <font color='blue'>【查詢筆數】</font>" + intRowsCount + "筆" + strBody;

        }
        else if (intRowsCount > 100)
        {
            strRet = "<div align='center'>** 查詢出有" + intRowsCount + "筆，已超過100筆，請增加查詢條件來縮小資料的範圍，以利作業順暢。 **</div>";
        }
        else
        {
            strRet = "<div align='center'>** 沒有符合條件的資料 **</div>";
        }

        return strRet;
    }

    private string GetGroupItemMember()
    {

        string strUid = Request.Form["DonorID"];
        string strItemUid = Request.Form["GroupItemUid"];
        string strMember = "";
        string strSql2 = @"
					select Donor_Id,Donor_Name,Sex,Title,Donor_Type,Cellular_Phone,Tel_Office,Tel_Office_Loc,Tel_Office_Ext,Email
					,(Case When D.IsAbroad = 'N' then B.Name + C.Name + D.[Address] Else D.[Address] End) as [Address] 
					,Invoice_Type,Remark
                    from GroupMapping GM1 
                    join Donor D on GM1.DonorId = D.Donor_Id and D.DeleteDate is null
			        left Join CODECITY As B on D.City = B.ZipCode 
				    left Join CODECITY As C on D.Area = C.ZipCode
	    			where GM1.DeleteDate is null and GM1.GroupItemUid = @ItemUid and GM1.DonorId = @Donor_Id
                    order by Donor_Id
                        ";


        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dict2.Add("ItemUid", strItemUid);
        dict2.Add("Donor_Id", strUid);
        DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);
        if (dt2.Rows.Count > 0)
        {

            DataRow dr2;
            dr2 = dt2.Rows[0];
            string strMemberTitle = "";
            string strTel = dr2["Tel_Office_Loc"].ToString().Trim() == "" ? "" : "(" + dr2["Tel_Office_Loc"].ToString().Trim() + ")";
            strTel += dr2["Tel_Office"].ToString().Trim();
            strTel += strTel + dr2["Tel_Office_Ext"].ToString().Trim() == "" ? "" : "#" + dr2["Tel_Office_Ext"].ToString().Trim();
            strMemberTitle += "性別：" + dr2["Sex"].ToString().Trim() + "\n";
            strMemberTitle += "稱謂：" + dr2["Title"].ToString().Trim() + "\n";
            strMemberTitle += "身分別：" + dr2["Donor_Type"].ToString().Trim() + "\n";
            strMemberTitle += "手機：" + dr2["Cellular_Phone"].ToString().Trim() + "\n";
            strMemberTitle += "電話(日)：" + strTel + "\n";
            strMemberTitle += "E-Mail：" + dr2["Email"].ToString().Trim() + "\n";
            strMemberTitle += "通訊地址：" + dr2["Address"].ToString().Trim() + "\n";
            strMemberTitle += "收據開立：" + dr2["Invoice_Type"].ToString().Trim() + "\n";
            strMemberTitle += "捐款人備註：\n" + dr2["Remark"].ToString().Trim() + "\n";
            strMember = "<SPAN onclick=\"window.event.cancelBubble=true;window.open('" +
                Util.RedirectByTime("/DonorMgr/DonorInfo_Edit.aspx", "Donor_Id=" + dr2["Donor_Id"].ToString().Trim()) +
                "','_self','')\" title='" + strMemberTitle + "'><font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "</font>(" + dr2["Donor_Name"].ToString().Trim() + ")</SPAN>";

        }

        return strMember;

    }

    private string GetGroupItemDetail()
    {

        string strRet = "";
        string strUid = Request.Form["DonorID"];
        string strItemUid = Request.Form["GroupItemUid"];
        string strMember = "";
        string strSql2 = @"
					select Donor_Id,Donor_Name,Sex,Title,Donor_Type,Cellular_Phone,Tel_Office,Tel_Office_Loc,Tel_Office_Ext,Email
					,(Case When D.IsAbroad = 'N' then B.Name + C.Name + D.[Address] Else D.[Address] End) as [Address] 
					,Invoice_Type,Remark,G.GroupClassName,I.GroupItemName,I.[Supplement]
                    from GroupMapping GM1 
                    join GroupItem I
                    on GM1.GroupItemUid=I.Uid
                    and I.DeleteDate is null
                    join GroupClass G
                    on I.GroupClassUid = g.[Uid]
                    and G.DeleteDate is null
                    join Donor D on GM1.DonorId = D.Donor_Id and D.DeleteDate is null
			        left Join CODECITY As B on D.City = B.ZipCode 
				    left Join CODECITY As C on D.Area = C.ZipCode
	    			where GM1.DeleteDate is null and GM1.GroupItemUid = @ItemUid
                    order by Donor_Id
                        ";

        Dictionary<string, object> dict2 = new Dictionary<string, object>();
        dict2.Add("ItemUid", strItemUid);
        DataTable dt2 = NpoDB.GetDataTableS(strSql2, dict2);
        if (dt2.Rows.Count > 0)
        {

            DataRow dr2;
            string GroupClassName = "";
            string GroupItemName = "";
            string Supplement = "";
            for (int j = 0; j < dt2.Rows.Count; j++)
            {
                dr2 = dt2.Rows[j];
                string strMemberTitle = "";
                string strTel = dr2["Tel_Office_Loc"].ToString().Trim() == "" ? "" : "(" + dr2["Tel_Office_Loc"].ToString().Trim() + ")";
                strTel += dr2["Tel_Office"].ToString().Trim();
                strTel += strTel + dr2["Tel_Office_Ext"].ToString().Trim() == "" ? "" : "#" + dr2["Tel_Office_Ext"].ToString().Trim();
                strMemberTitle += "性別：" + dr2["Sex"].ToString().Trim() + "\n";
                strMemberTitle += "稱謂：" + dr2["Title"].ToString().Trim() + "\n";
                strMemberTitle += "身分別：" + dr2["Donor_Type"].ToString().Trim() + "\n";
                strMemberTitle += "手機：" + dr2["Cellular_Phone"].ToString().Trim() + "\n";
                strMemberTitle += "電話(日)：" + strTel + "\n";
                strMemberTitle += "E-Mail：" + dr2["Email"].ToString().Trim() + "\n";
                strMemberTitle += "通訊地址：" + dr2["Address"].ToString().Trim() + "\n";
                strMemberTitle += "收據開立：" + dr2["Invoice_Type"].ToString().Trim() + "\n";
                strMemberTitle += "捐款人備註：\n" + dr2["Remark"].ToString().Trim() + "\n";
                if (strUid != dr2["Donor_Id"].ToString().Trim())
                {
                    strMember += (strMember == "" ? "<a href='" + Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=") +
                        dr2["Donor_Id"].ToString().Trim() + "' title='" + strMemberTitle + "'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")"
                        + "</a>" : ", " + "<a href='" + Util.RedirectByTime("DonorInfo_Edit.aspx", "Donor_Id=") +
                        dr2["Donor_Id"].ToString().Trim() + "' title='" + strMemberTitle + "'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")" + "</a>");
                }
                else
                {
                    strMember += (strMember == "" ? "<font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")"
                        + "</font>" : ", " + "<font color='blue'>" + dr2["Donor_Id"].ToString().Trim() + "(" + dr2["Donor_Name"].ToString().Trim() + ")" + "</font>");
                }
                GroupClassName = dr2["GroupClassName"].ToString().Trim();
                GroupItemName = dr2["GroupItemName"].ToString().Trim();
                Supplement = dr2["Supplement"].ToString().Trim();

            }

            strRet = @"<tr style='background-color: #FFFFCC'><td>" + GroupClassName + "</td><td><A href='" + Util.RedirectByTime("../DonorGroupMgr/GroupItemQry.aspx", "GroupItemUid=") + strItemUid + "'>" + GroupItemName + "</A></td><td>" + Supplement + "&nbsp;</td><td>" + strMember + "&nbsp;</td><td><img style='CURSOR: pointer' src='../images/delete2.gif' border=0 alt='刪除群組代表' onclick='deleteGroupItem(this," + strItemUid + ");' /></td></tr>";

        }

        return strRet;

    }

    private string DeleteGroupItem()
    {

        string strDonorID = Request.Form["DonorID"];
        string strGroupItemUid = Request.Form["GroupItemUid"];
        string strDeleteID = Request.Cookies["UserID"].Value;   //Request.Form["InsertID"];

        string strSql = " update [GroupMapping] set [DeleteID] = @DeleteID, [DeleteDate] = getdate() ";
        strSql += " where [DonorID] = @DonorID and [GroupItemUid] = @GroupItemUid ";
        strSql += " SELECT [GroupItemUid] FROM GroupMapping where DeleteDate is null and [DonorID] = @DonorID and [GroupItemUid] = @GroupItemUid ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("DonorID", strDonorID);
        dict.Add("GroupItemUid", strGroupItemUid);
        dict.Add("DeleteID", strDeleteID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        string strRet = "";
        if (dt.Rows.Count > 0)
        {
            //strRet = "Y";
        }
        else
        {
            strRet = "Y";
        }
        return strRet;

    }

    private string SaveGroupItem()
    {

        string strRet = "";
        string strDonorID = Request.Form["DonorID"];
        string strGroupItemUid = Request.Form["GroupItemUid"];
        string strInsertID = Request.Cookies["UserID"].Value;   //Request.Form["InsertID"];

        string strSql = " insert into [GroupMapping] ([GroupItemUid],[DonorID],[InsertID],[InsertDate])";
        strSql += " values (@GroupItemUid,@DonorID,@InsertID,getdate()) ";
        strSql += " SELECT SCOPE_IDENTITY() ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("DonorID", strDonorID);
        dict.Add("GroupItemUid", strGroupItemUid);
        dict.Add("InsertID", strInsertID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = "Y";
        }
        return strRet;
    }

    private string GetGroupClass()
    {
        string strSql = " SELECT  [Uid],[GroupClassName] FROM [GroupClass] where [DeleteDate] is null order by [Uid] ";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string retStr = "";
        retStr += "<option value=''>請選擇...</option>";
        foreach (DataRow dr in dt.Rows)
        {
            retStr += "<option value='" + dr["Uid"].ToString() + "'>" + dr["GroupClassName"].ToString() + "</option>";
        }
        return retStr;
    }

    private string GetGroupItem()
    {
        string GroupClassUid = Request.Form["GroupClassUid"];

        string strSql = " SELECT [Uid],[GroupItemName] FROM GroupItem where DeleteDate is null and GroupClassUid = @GroupClassUid ";
        strSql += " order by [GroupItemName] ";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("GroupClassUid", GroupClassUid);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string retStr = "<select id='GroupItem'>";
        retStr += "<option value=''>請選擇...</option>";
        foreach (DataRow dr in dt.Rows)
        {
            retStr += "<option value='" + dr["Uid"].ToString() + "'>" + dr["GroupItemName"].ToString() + "</option>";
        }
        retStr += "</select>";
        return retStr;
    }

    // 2014/6/25 增加新增捐款人時檢查已有相同通訊地址就會提示
    private string GetAddress()
    {
        string strRet = "";
        string strIsAbroad = Request.Form["IsAbroad"] == "true" ? "N" : "Y"; // Y/N
        string strCity = Request.Form["City"];
        string strStreet = Request.Form["Street"];
        string strSection = Request.Form["Section"];
        string strLane = Request.Form["Lane"];
        string strAlley = Request.Form["Alley"];
        string strHouseNo = Request.Form["HouseNo"];
        string strHouseNoSub = Request.Form["HouseNoSub"];
        string strFloor = Request.Form["Floor"];
        string strFloorSub = Request.Form["FloorSub"];
        string strRoom = Request.Form["Room"];
        //string strAttn = Request.Form["Attn"];
        string strOverseasCountry = Request.Form["OverseasCountry"];
        string strOverseasAddress = Request.Form["OverseasAddress"];

        string strSql = "select Donor_Id FROM DONOR where DeleteDate is null ";

        //國內
        if (strIsAbroad == "N")
        {
            strSql += " and Isnull([IsAbroad],'')='N' and Isnull([City],'')=@City and Isnull([Street],'')=@Street and Isnull([Section],'')=@Section ";
            strSql += " and Isnull([Lane],'')=@Lane and Isnull([Alley],'')=@Alley and Isnull([HouseNo],'')=@HouseNo and Isnull([HouseNoSub],'')=@HouseNoSub ";
            strSql += " and Isnull([Floor],'')=@Floor and Isnull([FloorSub],'')=@FloorSub and Isnull([Room],'')=@Room ;";
            //strSql += " and [Floor]=@Floor and [FloorSub]=@FloorSub and [Room]=@Room and [Attn]=@Attn ;";

        }
        //國外
        else if (strIsAbroad == "Y")
        {
            strSql += " and Isnull([IsAbroad],'')='Y' and Isnull([OverseasCountry],'')+Isnull([OverseasAddress],'') = @OverseasCountry+@OverseasAddress ;";

        }
        else
        {
            return "X";
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("City", strCity);
        dict.Add("Street", strStreet);
        dict.Add("Section", strSection);
        dict.Add("Lane", strLane);
        dict.Add("Alley", strAlley);
        dict.Add("HouseNo", strHouseNo);
        dict.Add("HouseNoSub", strHouseNoSub);
        dict.Add("Floor", strFloor);
        dict.Add("FloorSub", strFloorSub);
        dict.Add("Room", strRoom);
        //dict.Add("Attn", strAttn);
        dict.Add("OverseasCountry", strOverseasCountry);
        dict.Add("OverseasAddress", strOverseasAddress);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = "Y";
        }
        return strRet;

    }

    // 2014/6/25 增加新增捐款人時檢查已有相同收據地址就會提示
    private string GetInvoiceAddress()
    {

        string strRet = "";
        string strIsAbroad = Request.Form["IsAbroad"] == "true" ? "N" : "Y"; // Y/N
        string strCity = Request.Form["City"];
        string strStreet = Request.Form["Street"];
        string strSection = Request.Form["Section"];
        string strLane = Request.Form["Lane"];
        string strAlley = Request.Form["Alley"];
        string strHouseNo = Request.Form["HouseNo"];
        string strHouseNoSub = Request.Form["HouseNoSub"];
        string strFloor = Request.Form["Floor"];
        string strFloorSub = Request.Form["FloorSub"];
        string strRoom = Request.Form["Room"];
        //string strAttn = Request.Form["Attn"];
        string strOverseasCountry = Request.Form["OverseasCountry"];
        string strOverseasAddress = Request.Form["OverseasAddress"];

        string strSql = "select Donor_Id FROM DONOR where DeleteDate is null ";

        //國內
        if (strIsAbroad == "N")
        {
            strSql += " and Isnull([IsAbroad_Invoice],'')='N' and Isnull([Invoice_City],'')=@City and Isnull([Invoice_Street],'')=@Street ";
            strSql += " and Isnull([Invoice_Section],'')=@Section and Isnull([Invoice_Lane],'')=@Lane and Isnull([Invoice_Alley],'')=@Alley ";
            strSql += " and Isnull([Invoice_HouseNo],'')=@HouseNo and Isnull([Invoice_HouseNoSub],'')=@HouseNoSub and Isnull([Invoice_Floor],'')=@Floor ";
            strSql += " and Isnull([Invoice_FloorSub],'')=@FloorSub and Isnull([Invoice_Room],'')=@Room ;";
            //strSql += " and [Invoice_FloorSub]=@FloorSub and [Invoice_Room]=@Room and [Invoice_Attn]=@Attn ;";

        }
        //國外
        else if (strIsAbroad == "Y")
        {
            strSql += " and Isnull([IsAbroad_Invoice],'')='Y' ";
            strSql += " and Isnull([Invoice_OverseasCountry],'')+Isnull([Invoice_OverseasAddress],'') = @OverseasCountry+@OverseasAddress ;";

        }
        else
        {
            return "X";
        }

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("City", strCity);
        dict.Add("Street", strStreet);
        dict.Add("Section", strSection);
        dict.Add("Lane", strLane);
        dict.Add("Alley", strAlley);
        dict.Add("HouseNo", strHouseNo);
        dict.Add("HouseNoSub", strHouseNoSub);
        dict.Add("Floor", strFloor);
        dict.Add("FloorSub", strFloorSub);
        dict.Add("Room", strRoom);
        //dict.Add("Attn", strAttn);
        dict.Add("OverseasCountry", strOverseasCountry);
        dict.Add("OverseasAddress", strOverseasAddress);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = "Y";
        }
        return strRet;
    }

    //-------------------------------------------------------------------------
    private string GetNewsDetail()
    {
        //string NewsID = Request.Form["NewsID"];
        //string strSql = "select u.User_name, o.Org_name, news.*\n";
        //strSql += "from News\n";
        //strSql += "left join Admin_User u on News.PostUser_id=u.User_id\n";
        //strSql += "left join Organization o on u.Org_id=o.Org_id\n";
        //strSql += "where mbr_status=1 and mbr_num=@mbr_num";
        //Dictionary<string, object> dict = new Dictionary<string, object>();
        //dict.Add("mbr_num", NewsID);
        ////dict.Add("MBR_CHKREAD", Session["user_uid"].ToString());
        //DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        //if (dt.Rows.Count == 0)   6
        //{
        //    return "無此新聞";
        //}
        //DataRow dr = dt.Rows[0];
        //string strTemp = "";
        //HtmlTable table = new HtmlTable();
        //HtmlTableRow row;
        //HtmlTableCell cell;
        //CssStyleCollection css;

        //table.Border = 1;
        //table.CellPadding = 0;
        //table.CellSpacing = 0;
        //css = table.Style;
        //css.Add("font-size", "24px");
        //css.Add("font-family", "標楷體");
        //css.Add("width", "100%");

        //string strFontSize = "14px";

        //string strTitle = "系統公告";
        //row = new HtmlTableRow();
        //cell = new HtmlTableCell();
        //cell.InnerHtml = strTitle;
        //css = cell.Style;
        //css.Add("text-align", "center");
        //css.Add("font-size", "24");
        //css.Add("background-color", "#9CC2E7");
        //cell.ColSpan = 2;
        //row.Cells.Add(cell);
        //table.Rows.Add(row);

        //row = new HtmlTableRow();
        //cell = new HtmlTableCell();
        //cell.InnerHtml = "時間";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //css.Add("width", "20%");
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = GSSNPO.TransCDate(DBFunction.DB_DateTimeToShortString(dr["Mbr_Date"]));
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "標題";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = dr["Mbr_Title"].ToString();
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);

        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "發佈單位";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = "OrgName";
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "發佈者";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = "Username";
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "內容";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = "<textarea name='a' rows='20' cols='15' id='a' readonly='readonly' class='font9' style='width:100%;'>" + dr["Mbr_Content"].ToString() + "</textarea>";
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------
        //row = new HtmlTableRow();

        //cell = new HtmlTableCell();
        //cell.InnerHtml = "附件";
        //css = cell.Style;
        //css.Add("text-align", "right");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);

        //cell = new HtmlTableCell();
        //strTemp = dr["Mbr_File"].ToString();
        //string RealFilename = strTemp.Substring(strTemp.LastIndexOf('_') + 1);
        //cell.InnerHtml = strTemp == "" ? "&nbsp" : strTemp;
        //css = cell.Style;
        //css.Add("text-align", "left");
        //css.Add("font-size", strFontSize);
        //row.Cells.Add(cell);
        //table.Rows.Add(row);
        ////--------------------------------------------

        ////轉成 html 碼
        //StringWriter sw = new StringWriter();
        //HtmlTextWriter htw = new HtmlTextWriter(sw);
        //table.RenderControl(htw);

        //return htw.InnerWriter.ToString();
        return "";
    }
    //-------------------------------------------------------------------------
    private string GetHouseholdAddress()
    {
        //string CaseData_Uid = Request.Form["CaseData_Uid"];

        //string strSql = "";
        //if (type == 1)
        //{
        //    strSql += "select HouseholdCity as City, HouseholdVillage as Village, HouseholdZip as Zip, HouseholdAddress as Address\n";
        //}
        //else
        //{
        //    strSql += "select ContactCity as City, ContactVillage as Village, ContactZip as Zip, ContactAddress as Address\n";
        //}
        //strSql += "from ServiceAdvisory sa\n";
        //strSql += "inner join CaseData cd on sa.uid=cd.ServiceAdvisory_Uid\n";
        //strSql += "where cd.uid=@uid\n";

        //Dictionary<string, object> dict = new Dictionary<string, object>();
        //dict.Add("uid", CaseData_Uid);
        //DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        //if (dt.Rows.Count != 0)
        //{
        //    return dt.Rows[0]["City"].ToString().Trim() + "|" +
        //           dt.Rows[0]["Village"].ToString().Trim() + "|" +
        //           dt.Rows[0]["Zip"].ToString().Trim() + "|" +
        //           dt.Rows[0]["Address"].ToString().Trim();
        // }
        return "";
    }
    //-------------------------------------------------------------------------
    private string GetAreaList()
    {
        string CityID = Request.Form["CityID"];

        string strSql = "select ZipCode, Name from CODECITYNew\n";
        strSql += "where ParentCityID=@CityID\n";
        strSql += "order by Sort\n";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("CityID", CityID);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        string retStr = "";
        retStr += "<option value=''></option>";
        foreach (DataRow dr in dt.Rows)
        {
            retStr += "<option value='" + dr["ZipCode"].ToString() + "'>" + dr["Name"].ToString() + "</option>";
        }
        return retStr;
    }
    //-------------------------------------------------------------------------
    private string EmailIsExist()
    {
        //Thread.Sleep(20000);
        string strRet = "N"; 
        string Email = Request.Form["Email"];

        string strSql = "select Email from Donor\n";
        strSql += "where DeleteDate is null and Email=@Email and Donor_Pwd Is not NULL \n";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Email", Email);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = "Y";
        }
        return strRet;
    }
    //-------------------------------------------------------------------------
    public void DonorStat_Qry()
    {
        string Condition = Request.Form["Category"];
        Session["Condition"] = Condition;
        string strSql = "";
        if (Condition == "1")
        {
            strSql = @"select 
                        sum(case when Category='學生' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category='學生' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Category='個人' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category='個人' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Category='團體' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category='團體' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Category<>'學生' and Category<>'個人' and Category<>'團體'  then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category<>'學生' and Category<>'個人' and Category<>'團體' then 1 else 0 end) AS FLOAT)/(count (Category)) * 100 ,2) as [百分比]
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Category)) * 100 ,2) as [百分比]
                        from Donor
                        where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;

        }
        else if(Condition == "2")
        {
            strSql = @"select 
                        sum(case when ISNULL(Sex,'')='男' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='男' then 1 else 0 end) AS FLOAT)/(count (ISNULL(Sex,''))) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')='女' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='女' then 1 else 0 end) AS FLOAT)/(count (ISNULL(Sex,''))) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')='歿' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='歿' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')<>'男' and ISNULL(Sex,'')<>'女' and ISNULL(Sex,'')<>'歿'  then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')<>'男' and ISNULL(Sex,'')<>'女' and ISNULL(Sex,'')<>'歿' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        from Donor where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if(Condition == "3")
        {
            strSql = @"select 
                        sum(case when DATEDIFF(year ,Birthday,getdate())<21 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate())<21 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 21 and 25 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 21 and 25 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 26 and 30 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 26 and 30 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 31 and 35 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 31 and 35 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 36 and 40 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 36 and 40 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 41 and 45 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 41 and 45 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 46 and 50 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 46 and 50 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 51 and 55 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 51 and 55 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 56 and 60 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 56 and 60 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 61 and 65 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 61 and 65 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 66 and 70 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 66 and 70 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) >70 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) >70 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Birthday is null then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when Birthday is null then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if(Condition == "4")
        {
            strSql = @"select 
                        sum(case when Education='不識字' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='不識字' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='國小' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='國小' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='國中' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='國中' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='高中' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='高中' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='大學' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='大學' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='碩士' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='碩士' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='博士' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='博士' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='博士後研究' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='博士後研究' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='' or Education is null then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='' or Education is null then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
  
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if(Condition == "5")
        {
            strSql = @"select 
                        sum(case when Occupation='公教' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='公教' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='軍警' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='軍警' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='學生' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='學生' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='農' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='農' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='公' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='公' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='商' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='商' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='家管' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='家管' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='服務' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='服務' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='自由' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='自由' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='醫護' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='醫護' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='退休' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='退休' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='其他' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='其他' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation is null or Occupation='' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation is null or Occupation='' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if(Condition == "6")
        {
            strSql = @"select 
                        sum(case when Marriage='未婚' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='未婚' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='已婚' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='已婚' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='分居' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='分居' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='離婚' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='離婚' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='喪偶' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='喪偶' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='其他' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='其他' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage is null or Marriage = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage is null or Marriage = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if(Condition == "7")
        {
            strSql = @"select 
                        sum(case when Religion='基督教' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Religion='基督教' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Religion is null or Religion = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Religion is null or Religion = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if(Condition == "8")
        {
            strSql = @"select 
                        sum(case when A.mValue='基隆市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='基隆市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台北市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台北市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新北市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新北市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='桃園市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='桃園市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新竹市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新竹市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新竹縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新竹縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='苗栗縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='苗栗縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台中市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台中市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='彰化縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='彰化縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='南投縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='南投縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='雲林縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='雲林縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='嘉義市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='嘉義市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='嘉義縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='嘉義縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台南市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台南市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='高雄市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='高雄市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='屏東縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='屏東縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='宜蘭縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='宜蘭縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='花蓮縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='花蓮縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台東縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台東縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='澎湖縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='澎湖縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='金門縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='金門縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='連江縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='連江縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when City is null or City = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when City is null or City = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor Left Join CODECITY As A On DONOR.City=A.mCode where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "9")
        {
            strSql = @"select 
                        sum(case when A.mValue='基隆市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='基隆市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台北市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台北市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新北市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新北市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='桃園市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='桃園市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新竹市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新竹市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新竹縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新竹縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='苗栗縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='苗栗縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台中市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台中市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='彰化縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='彰化縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='南投縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='南投縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='雲林縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='雲林縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='嘉義市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='嘉義市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='嘉義縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='嘉義縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
 
                        ,sum(case when A.mValue='台南市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台南市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='高雄市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='高雄市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
 
                        ,sum(case when A.mValue='屏東縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='屏東縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='宜蘭縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='宜蘭縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='花蓮縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='花蓮縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台東縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台東縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='澎湖縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='澎湖縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='金門縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='金門縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='連江縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='連江縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Invoice_City is null or Invoice_City = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Invoice_City is null or Invoice_City = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor Left Join CODECITY As A On DONOR.Invoice_City=A.mCode where DeleteDate is null and ISNULL(IsMember,'') <> 'Y' ";
            Session["strSql"] = strSql;
        }
    }
    //-------------------------------------------------------------------------
    public void MemberStat_Qry()
    {
        string Condition = Request.Form["Category"];
        Session["Condition"] = Condition;
        string strSql = "";
        if (Condition == "1")
        {
            strSql = @"select 
                        sum(case when Category='學生' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category='學生' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Category='個人' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category='個人' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Category='團體' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category='團體' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Category<>'學生' and Category<>'個人' and Category<>'團體'  then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Category<>'學生' and Category<>'個人' and Category<>'團體' then 1 else 0 end) AS FLOAT)/(count (Category)) * 100 ,2) as [百分比]
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Category)) * 100 ,2) as [百分比]
                        from Donor
                        where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;

        }
        else if (Condition == "2")
        {
            strSql = @"select 
                        sum(case when ISNULL(Sex,'')='男' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='男' then 1 else 0 end) AS FLOAT)/(count (ISNULL(Sex,''))) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')='女' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='女' then 1 else 0 end) AS FLOAT)/(count (ISNULL(Sex,''))) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')='歿' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')='歿' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        ,sum(case when ISNULL(Sex,'')<>'男' and ISNULL(Sex,'')<>'女' and ISNULL(Sex,'')<>'歿'  then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when ISNULL(Sex,'')<>'男' and ISNULL(Sex,'')<>'女' and ISNULL(Sex,'')<>'歿' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]
                        from Donor where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "3")
        {
            strSql = @"select 
                        sum(case when DATEDIFF(year ,Birthday,getdate())<21 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate())<21 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 21 and 25 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 21 and 25 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 26 and 30 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 26 and 30 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 31 and 35 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 31 and 35 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 36 and 40 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 36 and 40 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 41 and 45 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 41 and 45 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 46 and 50 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 46 and 50 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 51 and 55 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 51 and 55 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 56 and 60 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 56 and 60 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 61 and 65 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 61 and 65 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) between 66 and 70 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) between 66 and 70 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when DATEDIFF(year ,Birthday,getdate()) >70 then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when DATEDIFF(year ,Birthday,getdate()) >70 then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Birthday is null then 1 else 0 end)  as [人數]
                        ,Round(CAST(sum(case when Birthday is null then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "4")
        {
            strSql = @"select 
                        sum(case when Education='不識字' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='不識字' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='國小' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='國小' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='國中' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='國中' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='高中' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='高中' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='大學' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='大學' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='碩士' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='碩士' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='博士' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='博士' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='博士後研究' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='博士後研究' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Education='' or Education is null then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Education='' or Education is null then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
  
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null  and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "5")
        {
            strSql = @"select 
                        sum(case when Occupation='公教' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='公教' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='軍警' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='軍警' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='學生' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='學生' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='農' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='農' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='公' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='公' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='商' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='商' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='家管' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='家管' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='服務' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='服務' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='自由' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='自由' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='醫護' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='醫護' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='退休' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='退休' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation='其他' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation='其他' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,sum(case when Occupation is null or Occupation='' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Occupation is null or Occupation='' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "6")
        {
            strSql = @"select 
                        sum(case when Marriage='未婚' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='未婚' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='已婚' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='已婚' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='分居' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='分居' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='離婚' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='離婚' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='喪偶' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='喪偶' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage='其他' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage='其他' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Marriage is null or Marriage = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Marriage is null or Marriage = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "7")
        {
            strSql = @"select 
                        sum(case when Religion='基督教' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Religion='基督教' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Religion is null or Religion = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Religion is null or Religion = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null  and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "8")
        {
            strSql = @"select 
                        sum(case when A.mValue='基隆市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='基隆市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台北市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台北市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新北市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新北市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='桃園市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='桃園市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新竹市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新竹市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='新竹縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='新竹縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='苗栗縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='苗栗縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台中市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台中市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='彰化縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='彰化縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='南投縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='南投縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='雲林縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='雲林縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='嘉義市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='嘉義市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='嘉義縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='嘉義縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台南市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台南市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='高雄市' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='高雄市' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='屏東縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='屏東縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='宜蘭縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='宜蘭縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='花蓮縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='花蓮縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='台東縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='台東縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='澎湖縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='澎湖縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='金門縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='金門縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when A.mValue='連江縣' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when A.mValue='連江縣' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when City is null or City = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when City is null or City = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor Left Join CODECITY As A On DONOR.City=A.mCode where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
        else if (Condition == "9")
        {
            strSql = @"select 
                        sum(case when Member_Status='新入會' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='新入會' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status='服務中' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='服務中' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status='停權' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='停權' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status='除名' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='除名' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status='退會' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='退會' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status='註銷' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='註銷' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status='死亡' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status='死亡' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,sum(case when Member_Status is null or Member_Status = '' then 1 else 0 end) as [人數]
                        ,Round(CAST(sum(case when Member_Status is null or Member_Status = '' then 1 else 0 end) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]

                        ,count (Donor_Id) as [人數]
                        ,Round(CAST(count (Donor_Id) AS FLOAT)/(count (Donor_Id)) * 100 ,2) as [百分比]
                        from Donor where DeleteDate is null and IsMember = 'Y' ";
            Session["strSql"] = strSql;
        }
    }
    //-------------------------------------------------------------------------
    public string LoadPledgeData()
    {
        string strSql = @"Select '0' as 'disabled',Pledge_Id,Pledge_Id as '授權編號',D.Donor_Name as 捐款人,Donate_Payment as 授權方式,
                          REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amt),1),'.00','')  as 扣款金額,
                          CONVERT(VarChar,Donate_FromDate,111) as 授權起日, CONVERT(VarChar,Donate_ToDate,111) as 授權迄日,Donate_Period as 轉帳週期,CONVERT(VarChar,Next_DonateDate,111) as 下次扣款日,
                          (Case When Valid_Date<>'' Then Left(Valid_Date,2)+'/'+Right(Valid_Date,2) Else '' End) as 有效月年, CONVERT(VarChar,Transfer_Date,111)  as 最後扣款日
                   From PLEDGE P Join DONOR D On P.Donor_Id=D.Donor_Id 
                   Where Status='授權中' And ((Post_SavingsNo<>'' And Post_AccountNo<>'') Or Account_No<>'' Or P_RCLNO<>'') ";
        string DonateDate = Request.Form["DonateDate"];
        if(DonateDate.Trim()!="")
        {
            strSql += " And Donate_FromDate <= '" + DonateDate + "' And Donate_ToDate>= '" + DonateDate + "' And ((Year(Next_DonateDate)<'" + Convert.ToDateTime(DonateDate).Year + "') Or (Year(Next_DonateDate)='" + Convert.ToDateTime(DonateDate).Year + "' And Month(Next_DonateDate)<='" + Convert.ToDateTime(DonateDate).Month + "')) ";
        }
        DataTable dt = NpoDB.GetDataTableS(strSql, null);

        if (dt.Rows.Count == 0)
        {
            return "";
        }
        //Grid initial
        NPOGridView npoGridView = new NPOGridView();
        npoGridView.Source = NPOGridViewDataSource.fromDataTable;
        npoGridView.dataTable = dt;
        npoGridView.Keys.Add("Pledge_Id");
        npoGridView.DisableColumn.Add("Pledge_Id");
        npoGridView.DisableColumn.Add("disabled");
        //-------------------------------------------------------------------------
        NPOGridViewColumn col = new NPOGridViewColumn("選擇");
        col.ColumnType = NPOColumnType.Checkbox;
        col.ControlKeyColumn.Add("Pledge_Id");
        col.ColumnName.Add("disabled");//決定是否check
        col.ControlText.Add("");
        col.ControlName.Add("PledgeId");
        col.ControlValue.Add("");
        col.ControlId.Add("checkbox");
        col.DisableValue = "0";
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權編號");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權編號");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("捐款人");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("捐款人");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權方式");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權方式");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("扣款金額");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("扣款金額");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權起日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權起日");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("授權迄日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("授權迄日");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("轉帳週期");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("轉帳週期");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("下次扣款日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("下次扣款日");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("有效月年");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("有效月年");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        col = new NPOGridViewColumn("最後扣款日");
        col.ColumnType = NPOColumnType.NormalText;
        col.ColumnName.Add("最後扣款日");
        npoGridView.Columns.Add(col);
        //-------------------------------------------------------------------------
        return npoGridView.Render();
    }
    //-------------------------------------------------------------------------
    private string PledgeBatchFileNameIsExist()
    {
        //Thread.Sleep(20000);
        string strRet = "N";
        string strFileName = Request.Params["file"].ToString();

        string strSql = @"select top 1 PledgeBatchFileName FROM Donate where Issue_Type<>'D' 
        and PledgeBatchFileName = @PledgeBatchFileName ;";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("PledgeBatchFileName", strFileName);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = "Y";
        }
        return strRet;
    }

    //-------------------------------------------------------------------------
    private string GetDonorIdFromEmail()
    {

        string strRet = "";
        string strAccount = Request.Form["txtAccount"];
        string strPwd = Request.Form["txtPwd"];

        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = @"select Donor_Id from DONOR where Email=@Email and Donor_Pwd=@Donor_Pwd and DeleteDate is null";
        dict.Add("Email", strAccount);
        dict.Add("Donor_Pwd", strPwd);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            strRet = dt.Rows[0]["Donor_Id"].ToString();
        }
        return strRet;
    }
    //-------------------------------------------------------------------------
    private string InvoiceTypeIsExist()
    {
        //Thread.Sleep(20000);
        string strRet = "N";
        string DonorId = Request.Form["DonorId"];

        string strSql = "select Year(ISNULL(IsInvoiceTypeExist,'')) as IsInvoiceTypeExist from Donor\n";
        strSql += "where DeleteDate is null and Donor_Id=@Donor_Id \n";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("Donor_Id", DonorId);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);

        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            if (dr["IsInvoiceTypeExist"].ToString().Trim() == DateTime.Now.Year.ToString())
            {
                strRet = "Y";
            }
        }
        return strRet;
    }
    //-------------------------------------------------------------------------

}   
