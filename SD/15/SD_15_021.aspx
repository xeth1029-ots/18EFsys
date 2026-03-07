<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_021.aspx.vb" Inherits="WDAIIP.SD_15_021" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>設備採購資訊查詢</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript" type="text/javascript">
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            //alert(getCheckBoxListValue(obj));
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                //若有全選
                if (getCheckBoxListValue(obj).charAt(0) == '1') {
                    myallcheck.checked = false; //全選改為false
                    document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
                }

            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;<FONT color="#990000">設備採購資訊查詢</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <div id="divSearch1" runat="server">
                        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol" width="17%">年度
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="yearlist1" runat="server">
                                    </asp:DropDownList>
                                    ～
								<asp:DropDownList ID="yearlist2" runat="server">
                                </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">分署
                                </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="Distid" runat="server" RepeatDirection="Horizontal" RepeatColumns="3">
                                    </asp:CheckBoxList>
                                    <input id="DistHidden" type="hidden" value="0" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">設備名稱
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="tDevName" runat="server" MaxLength="20" Width="210px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">委託採購廠商
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="tCPMOName" runat="server" MaxLength="20" Width="210px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4">
                                    <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                                    &nbsp;<asp:Button ID="btnAdd1" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div id="Div1" runat="server">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True">
                            <AlternatingItemStyle BackColor="#f5f5f5" />
                            <HeaderStyle CssClass="head_navy" />
                            <ItemStyle BackColor="White"></ItemStyle>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                    <div id="divSearch2" runat="server">
                        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_S"></asp:Button>
                                </td>
                            </tr>
                            <%--<tr>
							<td align="center">
								<asp:Button ID="BtnExp" runat="server" Text="匯出" CssClass="asp_button_S"></asp:Button>
							</td>
						</tr>--%>
                        </table>
                    </div>
                    <div id="divAdd1" runat="server">
                        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol_need" width="17%">年度
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlYears1" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">分署
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlDistID1" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">設備名稱
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TextBox1" runat="server" MaxLength="20" Width="210px">設備名稱</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">委託採購廠商
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TextBox2" runat="server" MaxLength="20" Width="210px">委託採購廠商</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">單價
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TextBox3" runat="server" MaxLength="20" Width="110px">950</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">數量
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TextBox4" runat="server" MaxLength="20" Width="110px">4</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">總價
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TextBox5" runat="server" MaxLength="20" Width="110px">3100</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">採購日期
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="STDate2" runat="server" Columns="10">2016/11/29</asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">備註
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TextBox8" runat="server" MaxLength="20" Width="210px">備註</asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4">&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4">
                                    <%--<asp:Button ID="btnSave1" runat="server" Text="新增下一筆" CssClass="asp_button_M"></asp:Button>--%>
								&nbsp;<asp:Button ID="btnSave2" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
