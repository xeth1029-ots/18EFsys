<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_010.aspx.vb" Inherits="WDAIIP.SYS_03_010" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>學員班級資料維護</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function Show_ListSort(tmpObj, tmpText, tmpValue, tmpAct) {
            var tmpID = tmpObj.selectedIndex;
            var tmpItem = tmpObj.value;

            if (tmpAct == "up") {
                if (tmpID > 0) {
                    tmpObj.options[tmpID].text = tmpObj.options[tmpID - 1].text;
                    tmpObj.options[tmpID].value = tmpObj.options[tmpID - 1].value;
                    tmpObj.options[tmpID - 1].text = tmpText;
                    tmpObj.options[tmpID - 1].value = tmpValue;
                }
            } else {
                if (tmpID < tmpObj.length - 1) {
                    tmpObj.options[tmpID].text = tmpObj.options[tmpID + 1].text;
                    tmpObj.options[tmpID].value = tmpObj.options[tmpID + 1].value;
                    tmpObj.options[tmpID + 1].text = tmpText;
                    tmpObj.options[tmpID + 1].value = tmpValue;
                }
            }
            tmpObj.value = tmpItem;
        }

        function Check_Data() {
            var errMsg = "";
            if (document.getElementById("txt_GroupName").value == "") {
                errMsg += "請輸入群組名稱。\n";
            }
            if (errMsg == "") {
                return true;
            } else {
                alert(errMsg);
                return false;
            }
        }

        function Show_SelectAll(tmpName1, tmpName2, tmpCnt) {
            if (document.getElementById(tmpName1)) {
                for (i = 0; i < tmpCnt; i++) {
                    if (document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2)))) {
                        if (document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2))).disabled == false) {
                            document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2))).checked = document.getElementById(tmpName1).checked;
                        }
                    }
                }
            }
        }

        function Show_NotCheck(tmpName1, tmpName2, tmpName3, tmpName4, tmpName5, tmpName6) {
            document.getElementById(tmpName1).checked = false;
            document.getElementById(tmpName2).checked = false;
            document.getElementById(tmpName3).checked = false;
            document.getElementById(tmpName4).checked = false;
            document.getElementById(tmpName5).checked = false;
            document.getElementById(tmpName6).checked = false;
        }

        function Show_SubList(itmName, tdName1, tdName2, subs, tdItems) {
            if (document.getElementById(itmName + "1").style.display == "inline" || document.getElementById(itmName + "1").style.display == "") {
                for (i = 1; i < subs; i++) {
                    document.getElementById(itmName + i).style.display = "none";
                }
                document.getElementById(itmName + "td0").rowSpan = 1;
                document.getElementById(itmName + "td1").colspan = tdItems - 1;
                for (i = 2; i < tdItems; i++) {
                    document.getElementById(itmName + "td" + i).style.display = "none";
                }
                document.getElementById(tdName1).style.display = "none";
                document.getElementById(tdName2).style.display = "inline";
            } else {
                for (i = 1; i < subs; i++) {
                    document.getElementById(itmName + i).style.display = "inline";
                }
                document.getElementById(itmName + "td0").rowSpan = subs;
                document.getElementById(itmName + "td1").colspan = 1;
                for (i = 2; i < tdItems; i++) {
                    document.getElementById(itmName + "td" + i).style.display = "inline";
                }
                document.getElementById(tdName1).style.display = "inline";
                document.getElementById(tdName2).style.display = "none";
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;學員班級資料維護</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Frametable" cellspacing="1" cellpadding="1" width="100%" border="0">

            <tr id="tr_Info" runat="server">
                <td>
                    <table class="table_nw" id="tb_Query" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號：
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txt_IDNO" runat="server" MaxLength="100" Width="50%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Button ID="btn_Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
							<%--<asp:LinkButton ID="LinkButton1" runat="server" Text="[功能2]" CssClass="linkbutton" Enabled="False" Visible="False"></asp:LinkButton>--%>
                            </td>
                        </tr>

                    </table>
                    <table class="font" id="tb_List" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <div align="left">
                                    E_Member&nbsp; (e網會員資料)：
								<asp:Label ID="lab_Msg10" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid10" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo10" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="mem_name" HeaderText="會員姓名" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="mem_idno" HeaderText="身分證號" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="會員生日" HeaderStyle-Width="9%">
                                            <ItemTemplate>
                                                <asp:Label ID="lmem_birth" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="mem_foreign" HeaderText="身分別" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="eloginT" HeaderText="最後登錄時間" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="mem_memo" HeaderText="備註" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="11%" />
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btn_Edit10" runat="server" Text="修改" CommandName="btnUpdata" CssClass="linkbutton"></asp:LinkButton><br />
                                                <asp:LinkButton ID="btn_Delete10" runat="server" Text="刪除" CommandName="btnDelete" CssClass="linkbutton"></asp:LinkButton><br />
                                                <asp:LinkButton ID="btn_Stop10" runat="server" Text="停用" CommandName="btnStop" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="btn_UnStop10" runat="server" Text="解除停用" CommandName="btnUnStop" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_StudentInfo&nbsp; (學員基本資料)：
								<asp:Label ID="lab_Msg1" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證號">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Name1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="生日">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Birthday1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="SID">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SID1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程資料">
                                            <HeaderStyle Width="36%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Class1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="SUBID">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SUBID1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%" />
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btn_Edit1" runat="server" Text="修改" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="btn_Dele1" runat="server" Text="刪除" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_EnterTemp&nbsp; (現場(TIMS)報名資料)：
								<asp:Label ID="lab_Msg2" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證號">
                                            <HeaderStyle Width="9%" />
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle Width="9%" />
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Name2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="生日">
                                            <HeaderStyle Width="9%" />
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Birthday2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="SETID">
                                            <HeaderStyle Width="9%" />
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SETID2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="eSETID">
                                            <HeaderStyle Width="9%" />
                                            <ItemTemplate>
                                                <asp:Label ID="lab_eSETID2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程資料">
                                            <HeaderStyle Width="36%" />
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Class2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%" />
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btn_Edit2" runat="server" Text="修改" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_EnterTemp2&nbsp; (e網報名資料)：
								<asp:Label ID="lab_Msg3" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證號">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Name3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="生日">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Birthday3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="SETID">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SETID3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="eSETID">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_eSETID3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程資料">
                                            <HeaderStyle Width="36%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Class3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btn_Edit3" runat="server" Text="修改" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    STUD_SELRESULTBLI&nbsp; (在職報名投保狀況資料)：
								<asp:Label ID="lab_Msg13" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid13" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo13" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證號">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO13" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Name13" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="SB3ID">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SB3ID" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程資料">
                                            <HeaderStyle Width="39%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Class13" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete13" runat="server" Text="清除" CommandName="btnDelete" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_trainingResults&nbsp; (結訓成績檔)：
								<asp:Label ID="lab_Msg4" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo4" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="SOCID" DataField="SOCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員姓名" DataField="name" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="身分證號" DataField="idno" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="classcname" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="courid" DataField="courid" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程名稱" DataField="CourseName" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="分數" DataField="Results" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="10%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete" runat="server" Text="刪除" CommandName="btnDelete" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_Conduct&nbsp; (操行明細檔)：
								<asp:Label ID="lab_Msg8" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid8" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo8" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="SOCID" DataField="SOCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員姓名" DataField="name" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="classcname" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="導師加減分" DataField="TechPoint" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="輔導課加減分" DataField="RemedPoint" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="出勤扣分" DataField="MinusLeave" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="獎懲扣分" DataField="MinusSanction" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="10%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete8" runat="server" Text="刪除" CommandName="btnDelete" Enabled="False" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_tranClassRecord&nbsp; (轉班資料)：
								<asp:Label ID="lab_Msg9" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid9" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo9" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="SOCID" DataField="SOCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員姓名" DataField="name" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="原班別代碼" DataField="OrigClassID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="原班別資料" DataField="classcname1" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="轉班級代碼" DataField="NewClassID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="轉班別資料" DataField="classcname2" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="轉班日期" DataField="ApplyDate" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="轉班原因" DataField="Reason" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="10%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete9" runat="server" Text="刪除" CommandName="btnDelete" Enabled="False" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Adp_GOVtrNData&nbsp; (推介記錄檔)：
								<asp:Label ID="lab_Msg5" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid5" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo5" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="推介姓名" DataField="name" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="推介身分證號" DataField="idno" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="推介班級trN_CLASS" DataField="trN_CLASS" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="推介是否轉入TIMS" DataField="transToTIMS" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="推介TIMS異動日" DataField="TIMSModifyDate" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="推介SOCID" DataField="SOCID1" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SOCID" DataField="SOCID2" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="classcname" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnUpdate" runat="server" Text="送3合1" CommandName="btnUpdate" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <%--STUD_BLACKLIST-學員處分--%>
                                <div align="left">
                                    STUD_BLACKLIST&nbsp; (學員處分)：
								<asp:Label ID="lab_Msg14" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="DataGrid14" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo14" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="處分起日" DataField="SBSDATE" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="年限" DataField="SBYEARS" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="事由" DataField="SBCOMMENT" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="處分緣由" DataField="SBTERMS" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <%-- <asp:TemplateColumn HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate> </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>

                                <%--STUD_BLIGATEDATA28-學員投保狀況檢核表-匯出e網民眾投保狀況檢核表--%>
                                <div align="left">
                                    STUD_BLIGATEDATA28&nbsp; (學員投保狀況檢核表)：
								<asp:Label ID="lab_Msg15" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="DataGrid15" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo15" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="CLASSCNAME" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="保險異動日" DataField="MDATE" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="保險證號" DataField="ACTNO" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="單位名稱" DataField="COMNAME" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="異動別" DataField="CHANGEMODE" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="投保薪資" DataField="SALARY" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="工作部門" DataField="DEPARTMENT" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="就保註記" DataField="BIEF" HeaderStyle-Width="10%"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>

                                <div align="left">
                                    Class_StudentsOfClass&nbsp; (班級學員檔)：
								<asp:Label ID="lab_Msg6" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid6" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo6" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="學員姓名" DataField="name" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員身分證號" DataField="idno" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SOCID" DataField="SOCID" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SID" DataField="SID" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="ClassName" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="報到日期" DataField="EnterDate" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員狀態" DataField="StudStatus" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="6%">
                                            <HeaderTemplate>功能</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center" Wrap="false"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDG6SCH1A" runat="server" Text="查詢重複" CommandName="SCH1A" CssClass="linkbutton"></asp:LinkButton><br />
                                                <asp:LinkButton ID="btnDG6UPD1A" runat="server" Text="(產投)<br/>暫時離訓" CommandName="UPD1A" CssClass="linkbutton"></asp:LinkButton><br />
                                                <asp:LinkButton ID="btnDG6UPD1B" runat="server" Text="還原<br/>(暫時離訓)" CommandName="UPD1B" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_ResultStudData&nbsp; (結訓學員資料卡檔)：
								<asp:Label ID="lab_Msg7" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid7" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo7" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="結訓資料卡號" DataField="dlid" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="結訓資料卡子號" DataField="SubNo" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="離訓日期" DataField="Rejecttdate1" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="退訓日期" DataField="Rejecttdate2" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SOCID" DataField="SOCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SID" DataField="SID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="classcname" HeaderStyle-Width="20%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete7" runat="server" Text="清除" CommandName="btnDelete" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_Turnout&nbsp; (學員出缺勤資料檔)：
								<asp:Label ID="lab_Msg11" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid11" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo11" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SOCID" DataField="SOCID" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="ClassName" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="開結訓日" DataField="sftdate" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="請假日期" DataField="leaveDate" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="假別代號" DataField="LeaveID" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="假別" DataField="c1Name" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="請假時數" DataField="Hours" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="請假節次" DataField="c12" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否計算出缺勤" DataField="TurnoutIgnore" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="訓練時數" DataField="THours" HeaderStyle-Width="8%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete11" runat="server" Text="清除" CommandName="btnDelete" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                                <div align="left">
                                    Stud_Turnout2&nbsp; (產投學員出缺勤資料檔)：
								<asp:Label ID="lab_Msg12" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </div>
                                <asp:DataGrid ID="Datagrid12" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo12" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="STOID" DataField="STOID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="OCID" DataField="OCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="學員SOCID" DataField="SOCID" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="課程資料" DataField="ClassName" HeaderStyle-Width="20%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="開結訓日" DataField="sftdate" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="請假日期" DataField="leaveDate" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="請假時數" DataField="Hours" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="訓練時數" DataField="THours" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete12" runat="server" Text="清除" CommandName="btnDelete" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="tr_Edit1" runat="server">
                <td>
                    <table class="table_nw" width="100%" runat="server" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditIDNO1" runat="server" Columns="12" MaxLength="15" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">姓名：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditName1" runat="server" Columns="8" MaxLength="20" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">生日：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditBirthday1" runat="server" Columns="10" MaxLength="10" Width="90%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">SID：
                            </td>
                            <td class="whitecol">
                                <asp:Label ID="lab_EditSID1" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol">SUBID：
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="lab_EditSUBID1" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訊息:</td>
                            <td colspan="5" class="whitecol">
                                <asp:Label ID="lab_msg_stud" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程資料：
                            </td>
                            <td colspan="5" class="whitecol">
                                <asp:RadioButtonList ID="rdo_EditClass1" runat="server" AutoPostBack="true" CssClass="font">
                                </asp:RadioButtonList>
                                <table class="table_sub" id="tb_EditClass1" width="100%" runat="server" cellspacing="1" cellpadding="1">
                                    <tr>
                                        <td class="bluecol_sub">課程名稱：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="lab_EditClass1" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_sub">SID：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_EditClassSID1" runat="server" Columns="20" MaxLength="24" Width="70%"></asp:TextBox><asp:DropDownList ID="list_EditSID1" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                        <td class="bluecol_sub">SETID：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_EditClassSETID1" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox><asp:DropDownList ID="list_EditSETID1" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" style="border-right-width: 0px; border-top-width: 0px; border-bottom-width: 0px; border-left-width: 0px" align="center" colspan="4">
                                            <asp:Button ID="btn_EditSaveClass1" runat="server" Text="儲存課程變更" CssClass="asp_button_M"></asp:Button>&nbsp;
										    <asp:Button ID="btn_EditdelClass1" runat="server" Text="刪除課程資料" CssClass="asp_button_M"></asp:Button>&nbsp;
										    <asp:Button ID="btn_Stud" runat="server" Text="學員資料" CssClass="asp_button_M"></asp:Button>&nbsp;
                                            <asp:Button ID="btn_EditSaveClass1B" runat="server" Text="(批次)儲存課程變更" CssClass="asp_button_M"></asp:Button>&nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="6" class="whitecol">
                                <asp:Button ID="btn_EditSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
							<asp:Button ID="btn_EditCancel1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="tr_Edit2" runat="server">
                <td>
                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditIDNO2" runat="server" Columns="12" MaxLength="15" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">姓名：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditName2" runat="server" Columns="8" MaxLength="20" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">生日：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditBirthday2" runat="server" Columns="10" MaxLength="10" Width="90%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">SETID：
                            </td>
                            <td class="whitecol">
                                <asp:Label ID="lab_EditSETID2" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol">eSETID：
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txt_EditeSETID2" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訊息:</td>
                            <td colspan="5" class="whitecol">
                                <asp:Label ID="lab_SelResult_msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程資料：
                            </td>
                            <td class="whitecol" colspan="5">
                                <asp:RadioButtonList ID="rdo_EditClass2" runat="server" AutoPostBack="true" CssClass="font">
                                </asp:RadioButtonList>
                                <table class="table_sub" id="tb_EditClass2" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="whitecol" align="center" colspan="4">
                                            <asp:Button ID="BTNSCH_T1_A" runat="server" Text="(產投)查詢重複" CssClass="asp_button_M"></asp:Button>
                                            <%--<asp:Button ID="BTNUPD_T1_5A" runat="server" Text="(產投)暫時報名失敗" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="BTNUPD_T1_5B" runat="server" Text="還原(暫時報名失敗)" CssClass="asp_button_M"></asp:Button>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_sub" width="15%">課程名稱：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="lab_EditClass2" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_sub" width="15%">SETID：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_EditClassSETID2" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox><asp:DropDownList ID="list_EditSETID2" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                        <td class="bluecol_sub" width="15%">eSETID：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_EditClasseSETID2" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox><asp:DropDownList ID="list_EditeSETID2" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center" colspan="4">
                                            <asp:Button ID="btn_EditSaveClass2" runat="server" Text="儲存課程變更" CssClass="asp_button_M"></asp:Button>&nbsp;
										    <asp:Button ID="btn_EditdelClass2" runat="server" Text="刪除課程資料" CssClass="asp_button_M"></asp:Button>&nbsp;
										    <asp:Button ID="btn_EditUpdateCls2" runat="server" Text="取消課程報到資料" CssClass="asp_button_M"></asp:Button>&nbsp;
                                            <asp:Button ID="btn_EditSaveClass2B" runat="server" Text="(批次)儲存課程變更" CssClass="asp_button_M"></asp:Button>&nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center" colspan="6">
                                <asp:Button ID="btn_EditSave2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
							<asp:Button ID="btn_EditCancel2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="tr_Edit3" runat="server">
                <td>
                    <table class="table_nw" id="table1" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditIDNO3" runat="server" Columns="12" MaxLength="15" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">姓名：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditName3" runat="server" Columns="8" MaxLength="20" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">生日：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditBirthday3" runat="server" Columns="10" MaxLength="10" Width="90%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">eSETID：
                            </td>
                            <td class="whitecol">
                                <asp:Label ID="lab_EditeSETID3" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol">SETID：
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txt_EditSETID3" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訊息:</td>
                            <td colspan="5" class="whitecol">
                                <asp:Label ID="Lab_ENTERTYPE2_mag" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程資料：
                            </td>
                            <td class="whitecol" colspan="5">
                                <asp:RadioButtonList ID="rdo_EditClass3" runat="server" AutoPostBack="true" CssClass="font">
                                </asp:RadioButtonList>
                                <table class="table_sub" id="tb_EditClass3" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol_sub" width="15%">課程名稱：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="lab_EditClass3" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_sub" width="15%">eSETID：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_EditClasseSETID3" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox><asp:DropDownList ID="list_EditeSETID3" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                        <td class="bluecol_sub" width="15%">SETID：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_EditClassSETID3" runat="server" Columns="10" MaxLength="32" Width="70%"></asp:TextBox><asp:DropDownList ID="list_EditSETID3" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center" colspan="4">
                                            <asp:Button ID="btn_EditSaveClass3" runat="server" Text="儲存課程變更" CssClass="asp_button_M"></asp:Button>&nbsp;
										    <asp:Button ID="btn_EditdelClass3" runat="server" Text="刪除課程資料" CssClass="asp_button_M"></asp:Button>&nbsp;
                                            <asp:Button ID="btn_EditSaveClass3B" runat="server" Text="(批次)儲存課程變更" CssClass="asp_button_M"></asp:Button>&nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center" colspan="6">
                                <asp:Button ID="btn_EditSave3" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
							<asp:Button ID="btn_EditCancel3" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="tr_Edit10" runat="server">
                <td>
                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditIdno10" runat="server" Columns="12" MaxLength="15" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">姓名：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_EditName10" runat="server" Columns="8" MaxLength="20" Width="90%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 16%">生日：
                            </td>
                            <td class="whitecol" style="width: 16%">
                                <asp:TextBox ID="txt_BirthDay10" runat="server" Columns="10" MaxLength="10" Width="90%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">mem_sn：
                            </td>
                            <td class="whitecol" colspan="5">
                                <asp:Label ID="lab_mem_sn" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center" colspan="6">
                                <asp:Button ID="btn_EditSave10" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
							<asp:Button ID="btn_EditCanedl10" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
