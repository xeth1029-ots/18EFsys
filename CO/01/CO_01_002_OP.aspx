<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_002_OP.aspx.vb" Inherits="WDAIIP.CO_01_002_OP" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>總場次資料</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;計畫參與度-總場次</asp:Label>
                </td>
            </tr>
        </table>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" style="width: 20%">轄區
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:CheckBoxList ID="sDistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        </asp:CheckBoxList>
                        <input id="sDistHidden" type="hidden" value="0" name="sDistHidden" runat="server">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need" style="width: 20%">年度
                    </td>
                    <td class="whitecol" style="width: 30%">
                        <asp:DropDownList ID="sYEARlist" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td class="bluecol_need" style="width: 20%">上／下半年度<%--申請階段--%></td>
                    <td class="whitecol" style="width: 30%">
                        <asp:DropDownList ID="sHALFYEAR" runat="server">
                            <asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="1">上年度</asp:ListItem>
                            <asp:ListItem Value="2">下年度</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">辦理活動場次日期
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="schPTYDATE1" Width="100" onfocus="this.blur()" runat="server" MaxLength="11"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schPTYDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        ～
						<asp:TextBox ID="schPTYDATE2" Width="100" onfocus="this.blur()" runat="server" MaxLength="11"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schPTYDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">活動場次名稱
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="sPTYNAME" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                </tr>

                <tr>
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        <asp:Button ID="BtnBack2" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <%--<asp:Button ID="btnExp1" runat="server" Text="匯出場次代碼" CssClass="asp_Export_M"></asp:Button>--%>
                    </td>
                </tr>
            </table>
            <table id="table_sch_show" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                        PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Years_ROC" HeaderText="年度" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="halfYearN" HeaderText="上/下<br>半年度" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="PLANSUB_N" HeaderText="計畫別" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="PTYDATE_RC" HeaderText="辦理活動場次日期" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="PTYNAME" HeaderText="活動場次名稱" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="編輯" CommandName="btnEdit" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>

                                        </Columns>
                                        <PagerStyle Visible="false"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>

        <div id="divEdt1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">

                <tr>
                    <td class="bluecol_need" width="15%">年度
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabPartyYears" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">轄區
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabDISTNAME" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">上/下半年度
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabhalfYear" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">計畫別
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="labPLANSUB" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">辦理活動場次日期
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="tPTYDATE" onfocus="this.blur()" runat="server" MaxLength="11" Width="20%"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= tPTYDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">活動場次名稱
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="tPTYNAME" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>--%>
                            <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </div>

        <asp:HiddenField ID="Hid_YEARS" runat="server" />
        <asp:HiddenField ID="Hid_DISTID" runat="server" />
        <asp:HiddenField ID="Hid_TPLANID" runat="server" />
        <asp:HiddenField ID="Hid_HALFYEAR" runat="server" />
        <asp:HiddenField ID="Hid_PTYID" runat="server" />

    </form>
</body>
</html>
