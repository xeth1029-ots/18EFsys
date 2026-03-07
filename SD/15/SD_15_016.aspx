<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_016.aspx.vb" Inherits="WDAIIP.SD_15_016" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>遞補人數統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        /* 全選 */
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;遞補人數統計表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">

            <tr>
                <td>
                    <table class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist" CssClass="font"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Distid" runat="server" RepeatColumns="3" RepeatDirection="Horizontal" Width="80%">
                                </asp:CheckBoxList>
                                <input id="DistHidden" value="0" type="hidden" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">辦訓地縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Tcitycode" runat="server" RepeatColumns="7" RepeatDirection="Horizontal" Width="80%">
                                </asp:CheckBoxList>
                                <input id="TcityHidden" value="0" type="hidden" name="TcityHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">立案地縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Ocitycode" runat="server" RepeatColumns="7" RepeatDirection="Horizontal" Width="80%">
                                </asp:CheckBoxList>
                                <input id="OcityHidden" value="0" type="hidden" name="OcityHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期
                            </td>
                            <td class="whitecol">
                                <span id="span01" runat="server">
                                    <asp:TextBox ID="SDate1" runat="server" Width="15%"></asp:TextBox>&nbsp;
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SDate1.ClientId %>','','','CY/MM/DD');" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                    &nbsp;~&nbsp;
                                <asp:TextBox ID="SDate2" runat="server" Width="15%"></asp:TextBox>&nbsp;
							    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SDate2.ClientId %>','','','CY/MM/DD');" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓日期
                            </td>
                            <td class="whitecol">
                                <span id="span02" runat="server">
                                    <asp:TextBox ID="EDate1" runat="server" Width="15%"></asp:TextBox>&nbsp;
                                <img style="cursor: pointer" onclick="Javascript:show_calendar('<%= EDate1.ClientId %>','','','CY/MM/DD');" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                    &nbsp;~&nbsp;
                                <asp:TextBox ID="EDate2" runat="server" Width="15%"></asp:TextBox>&nbsp;
                                <img style="cursor: pointer" onclick="Javascript:show_calendar('<%= EDate2.ClientId %>','','','CY/MM/DD');" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                                <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="btnExport" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                                <p align="center">
                                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="ResultTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <div id="Div1" runat="server">
                                    <%--<asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#f5f5f5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <ItemStyle BackColor="White"></ItemStyle>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>--%>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#f5f5f5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <%--<asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Wrap="False"></ItemStyle>
                                            </asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORGNAME" HeaderText="單位名稱"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDATE" HeaderText="開訓日期"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDATE" HeaderText="結訓日期"></asp:BoundColumn>

                                            <asp:BoundColumn DataField="TNUM" HeaderText="核定人數"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="NOENTNUM" HeaderText="遞補人數"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ENTERNUM" HeaderText="報名人數"></asp:BoundColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>

                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
