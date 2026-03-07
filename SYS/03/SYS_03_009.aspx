<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_009.aspx.vb" Inherits="WDAIIP.SYS_03_009" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>學員資料刪除</title>
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
                document.getElementById(itmName + "td1").colSpan = tdItems - 1;
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
                document.getElementById(itmName + "td1").colSpan = 1;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;學員資料刪除</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Frametable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="tb_Option" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;<font color="#990000">學員資料刪除</font> </td>
					</tr>
				</table>--%>
                    <table class="table_nw" id="tb_Query" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號： </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txt_IDNO" runat="server"></asp:TextBox>
                            </td>
                            <td class="bluecol">姓名： </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txt_Name" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">功能： </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rdo_Option" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="font">
                                    <asp:ListItem Value="StudentInfo" Selected="true">StudentInfo</asp:ListItem>
                                    <asp:ListItem Value="EnterTemp">EnterTemp</asp:ListItem>
                                    <asp:ListItem Value="EnterTemp2">EnterTemp2</asp:ListItem>
                                    <asp:ListItem Value="3">StudentInfo(DEL LOG)</asp:ListItem>
                                    <asp:ListItem Value="4">EnterTemp(DEL LOG)</asp:ListItem>
                                    <asp:ListItem Value="5">EnterTemp2(DEL LOG)</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Button ID="btn_Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp; </td>
                        </tr>
                    </table>
                    <table class="font" id="tb_List" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證號 / 資料編號">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO" runat="server"></asp:Label><br>
                                                <asp:Label ID="lab_SID" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Name" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="生日">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Birthday" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="25%"></HeaderStyle>
                                            <HeaderTemplate>
                                                <asp:Label ID="lab_Title1" runat="server"></asp:Label>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Class" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="25%"></HeaderStyle>
                                            <HeaderTemplate>
                                                <asp:Label ID="lab_Title2" runat="server"></asp:Label>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Lapm" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btn_Del" runat="server" Text="刪除" CommandName="Del" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid><asp:Label ID="lab_Msg" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="tb_Edit" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    </table>
                </td>
            </tr>
        </table>
        <input type="hidden" id="hSETID" runat="server" name="hSETID" />
        <input type="hidden" id="heSETID" runat="server" name="heSETID" />
        <input type="hidden" id="hidSID" runat="server" name="hidSID" />
        <input type="hidden" id="hidSOCID" runat="server" name="hidSOCID" />
        <input type="hidden" id="hsType" runat="server" name="hsType" />
    </form>
</body>
</html>
