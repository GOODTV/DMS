using System;
using System.Data;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.IO;

public enum NPOGridViewDataSource
{
    fromDataTable,
    fromSQLCommand
}
public enum NPOColumnType
{
    Checkbox,
    Radiobox,
    Dropdownlist,
    Textbox,
    NormalText,
    CurrecyText,
    Button,
    Link
}
public class NPOGridViewColumn
{
    public string Caption;
    public string Length;
    public NPOColumnType ColumnType;
    public string AddLink;
    public string Link;
    public string LinkText;
    public string IconPath;
    public string IconTooltip;
    public string RowSpan;
    public string ColSpan;
    public bool ShowTitle;
    //for Control
    public bool DataFromDataTable;
    public List<string> ControlName;
    public List<string> ControlId;
    public List<string> ControlValue;
    public List<string> ControlText;
    public List<string> ControlDefaultValue;
    public List<string> ControlTextSize;
    public List<string> ColumnName;
    public List<string> ControlOptionValue;
    public List<string> ControlOptionText;
    public List<string> ControlKeyColumn;
    public bool CaptionWithControl;
    public string DisableValue; //�� value �����Ȯ�, control �� disable ���A
    public string SortColumnName; //�ƧǮɪ����W��
    public bool Readonly;
    public bool ShowConfirmDialog;
    public string ConfirmDialogMsg;

    public NPOGridViewColumn()
    {
        ColumnType = NPOColumnType.NormalText;
        DataFromDataTable = true;
        ControlName = new List<string>();
        ControlId = new List<string>();
        ControlValue = new List<string>();
        ControlText = new List<string>();
        ControlDefaultValue = new List<string>();
        ControlTextSize = new List<string>();
        ColumnName = new List<string>();
        ControlKeyColumn = new List<string>();
        ControlOptionValue = new List<string>();
        ControlOptionText = new List<string>();
        AddLink = "";
        Link = "";
        LinkText = "";
        IconPath = "";
        IconTooltip = "";
        CaptionWithControl = false;
        DisableValue = "";
        ShowConfirmDialog = false;
        ConfirmDialogMsg = "";
        SortColumnName = "";
        RowSpan = "";
        ColSpan = "";
        ShowTitle = true;
    }
    public NPOGridViewColumn(string caption)
        : this()
    {
        Caption = caption;
    }
} //end of NPOGridViewColumn

public class NPOGridView
{
    public bool ShowTitlwWhenEmpty;
    public bool UseUserDefineColumn;
    public int PageSize;
    public int CurrentPage;
    public string CurrentSortFieldLinkID;
    public string CurrentSortDirLinkID;
    public string AddLink;
    public string EditLink;
    public string AddLinkTarget;
    public string EditLinkTarget;
    public string sqlCommand;
    public NPOGridViewDataSource Source;
    public DataTable dataTable;
    public Dictionary<string, object> dict;
    public string PreviousPageLinkID;
    public string NextPageLinkID;
    public string GoPageLinkID;
    public string CurrentPageLinkID;
    public string GoPageControlName;
    public string AddLinkFontSize;
    public string AddLinkFontFamily;
    public string AddLinkText;
    public bool ShowPage;
    public string AddLinkImage;
    public string CssClass;
    public string TableID;
    public List<NPOGridViewColumn> ExtraColumns;
    public List<NPOGridViewColumn> Columns;
    public List<string> DisableColumn;
    public List<string> Keys;
    public List<string> QuerySyting;
    public List<string> QuerySytingData;
    public string EditClick;
    //20140509 �s�W by Ian_Kao
    //�ܧ����C���
    public string ColumnNameToChangeColor;
    public string ColumnDataToChangeColor;
    //----------------------------------------------------------------------
    public NPOGridView()
    {
        ShowTitlwWhenEmpty = false;
        PageSize = 30;
        CurrentPage = 1;
        CurrentSortFieldLinkID = "HFD_CurrentSortField";
        CurrentSortDirLinkID = "HFD_CurrentSortDir";
        AddLinkTarget = "_self";
        EditLinkTarget = "_self";
        AddLink = "";
        EditLink = "";
        Source = NPOGridViewDataSource.fromSQLCommand;
        PreviousPageLinkID = "btnPreviousPage";
        NextPageLinkID = "btnNextPage";
        GoPageLinkID = "btnGoPage";
        CurrentPageLinkID = "HFD_CurrentPage";
        GoPageControlName = "GoPage";
        AddLinkFontSize = "9pt";
        AddLinkFontFamily = "�s�ө���";
        AddLinkText = "�s�W���";
        AddLinkImage = "../images/DIR_tri.gif";
        CssClass = "table_h";
        TableID = "";
        ShowPage = true;
        Columns = new List<NPOGridViewColumn>();
        ExtraColumns = new List<NPOGridViewColumn>();
        DisableColumn = new List<string>();
        QuerySyting = new List<string>();
        QuerySytingData = new List<string>();
        Keys = new List<string>();
        UseUserDefineColumn = false;
        EditClick = "";
        //20140509 �s�W by Ian_Kao
        //�ܧ����C���
        ColumnNameToChangeColor = "";
        ColumnDataToChangeColor = "";
    }
    //----------------------------------------------------------------------
    public string Render()
    {
        string strTemp = "";
        HtmlTable tableHeader = new HtmlTable();
        HtmlTableRow row;
        HtmlTableCell cell;
        CssStyleCollection css;

        if (Source == NPOGridViewDataSource.fromSQLCommand)
        {
            GetData();
        }

        int RecordCount = dataTable.Rows.Count;
        //�����P���ƭp��
        int TotalPage = 0;
        if (RecordCount % PageSize == 0)
        {
            TotalPage = RecordCount / PageSize;
        }
        else
        {
            TotalPage = RecordCount / PageSize + 1;
        }

        if (CurrentPage > TotalPage)
        {
            CurrentPage = TotalPage;
        }
        if (CurrentPage <= 0)
        {
            CurrentPage = 1;
        }
        int endRecNo = CurrentPage * PageSize - 1;
        int startRecNo = (CurrentPage - 1) * PageSize;
        if (endRecNo > RecordCount - 1)
        {
            endRecNo = RecordCount - 1;
        }
        //�N�����n��ܪ���ƽƻs�@��
        DataTable dtTemp = dataTable.Clone();
        if (ShowPage == true)
        {
            for (int i = startRecNo; i <= endRecNo; i++)
            {
                dtTemp.ImportRow(dataTable.Rows[i]);
            }
        }
        else
        {
            dtTemp = dataTable;
        }
        //end of �����P���ƭp��
        //-------------------------------------------------------------------
        //�e���W�誺�����M�W�@���U�@���\��        
        string strPage = "<table class='pager' border='0' cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>\n";
        strPage += "<tr><td width='20%'></td>";
        strPage += "<td width='60%' align='center'>";
        //��ܭ�ñ�P�_
        if (ShowPage == true && RecordCount > 0)
        {
            strPage += "   <span>�� " + CurrentPage + "/" + TotalPage + " ��&nbsp;&nbsp;</span>";
            if (CurrentPage != 1)
            {
                string Link = " | <a onclick=\"document.getElementById('" + CurrentPageLinkID + "').value='" + CurrentPage + "';window.event.cancelBubble=true;document.getElementById('" + PreviousPageLinkID + "').click();\" >�W�@��</a>";
                strPage += Link;
            }
            if (CurrentPage != TotalPage && CurrentPage < TotalPage)
            {
                string Link = " | <a onclick=\"document.getElementById('" + CurrentPageLinkID + "').value='" + CurrentPage + "';window.event.cancelBubble=true;document.getElementById('" + NextPageLinkID + "').click();\"  >�U�@��</a>";
                strPage += Link;
            }
            strPage += " |&nbsp;<span> ���ܲ� <select name=" + GoPageControlName + " size='1' onchange=\"document.getElementById('" + CurrentPageLinkID + "').value=" + GoPageControlName + ".options[" + GoPageControlName + ".selectedIndex].value;window.event.cancelBubble=true;document.getElementById('" + GoPageLinkID + "').click();\" style='text-decoration:none'>";
            int iPage = 0;
            string strSelected = "";
            for (iPage = 1; iPage <= TotalPage; iPage++)
            {
                if (iPage == CurrentPage)
                {
                    strSelected = "selected";
                }
                else
                {
                    strSelected = "";
                }
                strPage += "<option value=\'" + iPage + "\' " + strSelected + ">" + iPage + "</option></br>";
            }
            strPage += "</select> ��</span>�@�@ " + RecordCount + " ��</br>";
        }
        else
        {
            strPage += "";
        }
        strPage += "<td width='20%'></td>\n";
        strPage += "</table>\n";
        //end of �e���W�誺�����M�W�@���U�@���\��        
        //-------------------------------------------------------------------
        //�����ܻP Link
        string showhand = "";
        string onClickString = "";
        string Cellshowhand = "";
        string CellonClickString = "";
        string strBody = "";
        string strTableID = "";

        //��ܩ��Y�C
        if (TableID != "")
        {
            strTableID = " id='" + TableID + "' ";
        }
        strBody += "<table " + strTableID + " class='" + CssClass + "'" + " width='100%'>";

        //�� Title �i�H�h�@��,�[�j��ı�ĪG
        if (ExtraColumns.Count != 0)
        {
            strBody += "<tr>";
            foreach (NPOGridViewColumn col in ExtraColumns)
            {
                string RowColSpan = " ";
                if (col.RowSpan != "")
                {
                    RowColSpan += " rowspan='" + col.RowSpan + "' ";
                }
                if (col.ColSpan != "")
                {
                    RowColSpan += " colspan='" + col.ColSpan + "' ";
                }
                strBody += "<th nowrap " + RowColSpan + "><span>" + col.Caption + "</span></th>\n";
            }
            strBody += "</tr>";
        }

        strBody += "<tr>";


        //�����ন�p�g,�קK��Ʈw���j�p�g���D
        for (int i = 0; i < DisableColumn.Count; i++)
        {
            DisableColumn[i] = DisableColumn[i].ToLower();
        }

        List<ControlData> list = new List<ControlData>();


        //��ܩ��Y
        if (Columns.Count == 0) //�S���w�q���,��  DataTable �� Columns
        {
            foreach (DataColumn dc in dtTemp.Columns)
            {   //DisableColumn �ҦC����줣���,�u�� key �ȥ�
                if (!DisableColumn.Contains(dc.ColumnName.ToLower()))
                {
                    strBody += "<th nowrap><span>" + dc.ColumnName + "</span></th>\n";
                }
            }
        }
        else
        {
            //�Φۦ�w�q�������, ��ܩ��Y
            foreach (NPOGridViewColumn col in Columns)
            {
                if (col.ShowTitle == false)
                {
                    continue;
                }
                string Content = "";
                if (col.CaptionWithControl == true)
                {
                    list.Clear();
                    switch (col.ColumnType)
                    {
                        case NPOColumnType.Checkbox:
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                list.Add(new ControlData("Checkbox", col.ControlName[j] + "All", col.ControlId[j] + "All", col.ControlValue[j], "", false, ""));
                            }
                            break;
                        case NPOColumnType.Radiobox:
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                list.Add(new ControlData("Radio", col.ControlName[j] + "All", col.ControlId[j] + "All", col.ControlValue[j], "", false, ""));
                            }
                            break;
                        case NPOColumnType.Dropdownlist:
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                list.Add(new ControlData("DropdownList", col.ControlName[j] + "All", col.ControlId[0] + "All", col.ControlOptionValue.ToArray(), col.ControlOptionText.ToArray(), false, ""));
                            }
                            break;
                    }
                    Content = HtmlUtil.RenderControl(list) + "&nbsp";
                }
                //�I���Y�i�H�Ƨ�
                string Link = "";
                if (col.SortColumnName != "")
                {
                    Link = "onclick='$(\"#" + CurrentSortFieldLinkID + "\").val(\"" + col.SortColumnName + "\"); $(\"#" + CurrentPageLinkID + "\").val(\"1\");;if ($(\"#" + CurrentSortDirLinkID + "\").val() == \"asc\"){ $(\"#" + CurrentSortDirLinkID + "\").val(\"desc\");} else $(\"#" + CurrentSortDirLinkID + "\").val(\"asc\");; $(\"#" + GoPageLinkID + "\").click();' style='text-decoration:none;cursor:hand' ";
                }

                string RowColSpan = " ";
                if (col.RowSpan != "")
                {
                    RowColSpan += " rowspan='" + col.RowSpan + "' ";
                }
                if (col.ColSpan != "")
                {
                    RowColSpan += " colspan='" + col.ColSpan + "' ";
                }
                strBody += "<th nowrap " + RowColSpan + Link + "><span>" + Content + col.Caption + "</span></th>\n";
                //Todo:�[ asc �� desc �ϥ�
                //strBody += "<th nowrap " + Link + "><span>" + Content + col.Caption + "</span><img src='../images/asc.gif' ></th>\n";
            }
        }
        strBody += "</tr>";

        Page page = HttpContext.Current.CurrentHandler as Page;
        //��ƦC
        int IdCount = 1;
        foreach (DataRow dr in dtTemp.Rows)
        {
            if (EditLink != "")
            {
                string EditURL = "";
                string KeyValue = "";
                for (int i = 0; i < Keys.Count; i++)
                {
                    KeyValue += dr[Keys[i]];
                }
                EditURL = EditLink + page.Server.UrlEncode(KeyValue);
                if (QuerySyting.Count != 0)
                {
                    foreach (string str in QuerySyting)
                    {
                        EditURL = EditURL + "&" + str + "=" + dr[str].ToString();
                    }
                }
                onClickString = " onclick =\"window.open(\'" + EditURL + "\',\'" + EditLinkTarget + "\',\'\')\" ";
                showhand = @"style='cursor:hand'";
            }
            if (EditClick != "")
            {
                string Edit_JavaScript = "";
                string KeyValue = "";
                for (int i = 0; i < Keys.Count; i++)
                {
                    KeyValue += dr[Keys[i]]+"|";
                }
                Edit_JavaScript = EditClick + "('" + KeyValue + "')";
                onClickString = " OnClick =" + "\"" + Edit_JavaScript + ";" + "\"";
                showhand = @"style='cursor:hand'";
            }
            if (Columns.Count == 0)
            {
                strBody += "<tr " + showhand + "  " + onClickString + ">";
                foreach (DataColumn dc in dtTemp.Columns)
                {
                    if (!DisableColumn.Contains(dc.ColumnName.ToLower()))
                    {
                        string Content = dr[dc.ColumnName].ToString();
                        if (Content == "")
                        {
                            Content = "&nbsp;";
                        }
                        strBody += "<td><span>" + Content + "</span></td>\n";
                    }
                }
            }
            else
            { //���w�q������
                strBody += "<tr " + showhand + "  " + onClickString + ">\n";
                foreach (NPOGridViewColumn col in Columns)
                {
                    string Content = "";
                    string ColumnValue = "";
                    string DefValue = "";
                    string KeyValue = "";
                    string ConfirmStr = "";
                    for (int j = 0; j < col.ControlKeyColumn.Count; j++)
                    {
                        if (col.ControlKeyColumn[j] != "")
                        {
                            KeyValue += dr[col.ControlKeyColumn[j]].ToString();
                        }
                    }
                    switch (col.ColumnType)
                    {
                        case NPOColumnType.NormalText:
                            if (col.DataFromDataTable == true)
                            {
                                //Content = dr[col.ColumnName[0]].ToString();
                                Content = HtmlUtil.ToHtml(dr[col.ColumnName[0]].ToString());
                            }
                            else
                            {
                                Content = col.ControlValue[0];
                            }
                            Content = Content == "" ? "&nbsp;" : Content;
                            break;
                        case NPOColumnType.CurrecyText:
                            if (col.DataFromDataTable == true)
                            {
                                //Content = dr[col.ColumnName[0]].ToString();
                                Content = "<font color='red'>NT$ " + HtmlUtil.ToHtml(string.Format("{0:N}", Convert.ToUInt32(dr[col.ColumnName[0]].ToString()))) + "</font>";
                            }
                            else
                            {
                                Content = col.ControlValue[0];
                            }
                            Content = Content == "" ? "&nbsp;" : Content;
                            break;
                        case NPOColumnType.Checkbox:
                            list.Clear();
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                DefValue = "";
                                if (col.ColumnName[j] != "")
                                {
                                    DefValue = dr[col.ColumnName[j]].ToString();
                                }
                                list.Add(new ControlData("Checkbox", col.ControlName[j] + "_" + KeyValue, col.ControlId[j], col.ControlValue[j] + KeyValue, col.ControlText[j], false, DefValue));
                                if (DefValue != "" && DefValue == col.DisableValue)
                                {
                                    list[list.Count - 1].Disabled = true;
                                }
                                //20140509 �s�W by Ian_Kao �H�Υd���Ĥ�~�L���ᤣ�Ŀ�
                                if (ColumnNameToChangeColor != "")
                                {
                                    string strTmp = ColumnDataToChangeColor;
                                    DateTime start = DateTime.Now;
                                    DateTime end = DateTime.Now;
                                    start = DateTime.Parse(dr[ColumnNameToChangeColor].ToString());
                                    end = DateTime.Parse(strTmp);
                                    int month = CaseUtil.DateDiff(start, end);
                                    if (month < 0)
                                    {
                                        list[list.Count - 1].Disabled = true;
                                    }
                                }
                            }
                            Content = HtmlUtil.RenderControl(list);
                            break;
                        case NPOColumnType.Radiobox:
                            list.Clear();
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                DefValue = "";
                                if (col.ColumnName[j] != "")
                                {
                                    DefValue = dr[col.ColumnName[j]].ToString();
                                }
                                list.Add(new ControlData("Radio", col.ControlName[j] + "_" + KeyValue, col.ControlId[j] + "_" + KeyValue, col.ControlValue[j], col.ControlText[j], false, DefValue));
                                if (DefValue != "" && DefValue == col.DisableValue)
                                {
                                    list[list.Count - 1].Disabled = true;
                                }
                            }
                            Content = HtmlUtil.RenderControl(list);
                            break;
                        case NPOColumnType.Dropdownlist:
                            list.Clear();
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                DefValue = "";
                                if (col.ColumnName[j] != "")
                                {
                                    DefValue = dr[col.ColumnName[j]].ToString();
                                }
                                list.Add(new ControlData("DropdownList", col.ControlName[j] + "_" + KeyValue, col.ControlId[j] + "_" + KeyValue, col.ControlOptionValue.ToArray(), col.ControlOptionText.ToArray(), false, DefValue));
                            }
                            Content = HtmlUtil.RenderControl(list);
                            break;
                        case NPOColumnType.Textbox:
                            list.Clear();
                            for (int j = 0; j < col.ControlName.Count; j++)
                            {
                                if (col.DataFromDataTable == true)
                                {
                                    //����ƨӷ�,�S�]�w����,�N�H DataTable ����쬰�ӷ�
                                    if (col.ColumnName.Count == 0)
                                    {
                                        ColumnValue = "";
                                    }
                                    else
                                    {
                                        //�w�]������W�٬��ťծ�,
                                        if (col.ColumnName.Count == 0)
                                        {
                                            ColumnValue = "";
                                        }
                                        else
                                        {
                                            if (col.ColumnName[j] == "")
                                            {
                                                ColumnValue = "";
                                            }
                                            else
                                            {
                                                //�H���W�٧���
                                                ColumnValue = dr[col.ColumnName[j]].ToString();
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    ColumnValue = col.ControlValue[j];
                                }
                                list.Add(new ControlData("Text", col.ControlName[j] + "_" + KeyValue, col.ControlId[j] + "_" + KeyValue, ColumnValue, "", false, ""));
                                //�� Textbox ��쪺 DisableValue ���]�w�Ȫ���,���ޭȬ��󳣻{�w�� disabled
                                if (col.Readonly == true)
                                {
                                    list[list.Count - 1].Readonly = true;
                                }
                            }
                            Content = HtmlUtil.RenderControl(list);
                            break;
                        case NPOColumnType.Link:
                            list.Clear();
                            if (col.ShowConfirmDialog == true)
                            {
                                ConfirmStr = "if (!window.confirm('" + col.ConfirmDialogMsg + "')) return;";
                            }
                            //CellonClickString = " onclick =\"" + ConfirmStr + "window.open(\'" + col.Link + page.Server.UrlEncode(KeyValue) + "\',\'" + EditLinkTarget + "\',\'\')\" ";
                            CellonClickString = " onclick =\"" + ConfirmStr + "window.open(\'" + col.Link + page.Server.UrlEncode(KeyValue) + "\',\'" + EditLinkTarget + "\',\'\')\" ";
                            Cellshowhand = @"style='cursor:hand;color:#8b0000;text-decoration:underline;font-weight:100;'";
                            if (col.LinkText != "")
                            {
                                Content += "<span " + Cellshowhand + "  " + CellonClickString + ">" + col.LinkText + "</span>";
                            }
                            else
                            {
                                //�ťիh�ϥ� icon �ϥ�
                                string icon = "<img src='" + col.IconPath + "' border=0 width='20' height='20' title='" + col.IconTooltip + "'>";
                                Content += "<span " + Cellshowhand + "  " + CellonClickString + ">" + icon + "</span>";
                            }
                            break;
                        case NPOColumnType.Button:
                            list.Clear();
                            if (col.ShowConfirmDialog == true)
                            {
                                ConfirmStr = "if (!window.confirm('" + col.ConfirmDialogMsg + "')) return;";
                            }
                            CellonClickString = " onclick =\"window.event.cancelBubble=true;" + ConfirmStr + "window.open(\'" + col.Link + page.Server.UrlEncode(KeyValue) + "\',\'" + EditLinkTarget + "\',\'\');\" ";
                            Cellshowhand = @"style='cursor:hand;'";
                            string Disabled = "";
                            //Button �O�_�n Disabled
                            string value = "";
                            if (col.ColumnName.Count != 0)
                            {
                                value = dr[col.ColumnName[0]].ToString();
                            }
                            if (value != "" && value == col.DisableValue)
                            {
                                Disabled = "disabled='disabled'";
                            }
                            Content += "<input type='button'" + Disabled + " value='" + col.ControlText[0] + "'" + Cellshowhand + "  " + CellonClickString + "></input>";
                            break;
                    }
                    //-------------------
                    string strChangeColor = "";
                    if (ColumnNameToChangeColor != "")
                    {
                        string strTmp = ColumnDataToChangeColor;
                        DateTime start = DateTime.Now;
                        DateTime end = DateTime.Now;
                        start = DateTime.Parse(dr[ColumnNameToChangeColor].ToString());
                        end = DateTime.Parse(strTmp);
                        int month = CaseUtil.DateDiff(start,end);
                        if (month < 0 )
                        {
                            strChangeColor = @"style='background-color:#FF6666;'";
                        }
                        else if (month == 0)
                        {
                            strChangeColor = @"style='background-color:#FFCCFF;'";
                        }
                        else if (month == 1)
                        {
                            strChangeColor = @"style='background-color:#FFFF99;'";
                        }
                        else if (month == 2)
                        {
                            strChangeColor = @"style='background-color:#66FF99;'";
                        }
                    }
                    //-------------------
                    strBody += "<td " + strChangeColor + "><span>" + Content + "</span></td>\n";
                }
            }
            strBody += "</tr>\n";
            IdCount++;
        }
        strBody += "</table>\n";
        //--------------------------------------------
        string Result = strPage + strBody;
        return Result;
    }
    //----------------------------------------------------------------------
    private void GetData()
    {
        dataTable = NpoDB.GetDataTableS(sqlCommand, dict);
    }
    //----------------------------------------------------------------------
    private string GetHtmlText(HtmlTable table)
    {
        StringWriter sw = new StringWriter();
        HtmlTextWriter htw = new HtmlTextWriter(sw);
        table.RenderControl(htw);

        return htw.InnerWriter.ToString();
    }
    //----------------------------------------------------------------------
} //end of NPOGridView
