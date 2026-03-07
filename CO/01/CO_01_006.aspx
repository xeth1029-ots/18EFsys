<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_006.aspx.vb" Inherits="WDAIIP.CO_01_006" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>TTQS單位確認開關機制</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <%--<script type="text/javascript" language="javascript"></script>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;TTQS單位確認開關機制</asp:Label>
                </td>
            </tr>
        </table>
        <div id="div_sch1" runat="server">
            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr>
                    <td width="20%" class="bluecol">設定日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="schCCDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span1" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schCCDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span2" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schCCDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schCCDATE2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span3" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schCCDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span4" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schCCDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">控制起始日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="schCSDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span5" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schCSDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span6" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schCSDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schCSDATE2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span7" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schCSDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span8" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schCSDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">控制結束日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="schCFDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span9" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schCFDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span10" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schCFDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schCFDATE2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span11" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schCFDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span12" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schCFDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">查詢狀態：</td>
                    <td class="whitecol">
                        <asp:CheckBox ID="CHK_ISDELETE" runat="server" Text="包含已刪除" />
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol" align="center">
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_S" AuthType="QRY"></asp:Button>
                        <asp:Button ID="bt_add" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_S" AuthType="ADD"></asp:Button>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol">
                        <asp:Label ID="lab_msg1" runat="server" ForeColor="Red"></asp:Label></td>
                </tr>
            </table>
            <table id="tb_Sch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                <tr>
                    <td align="center">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="CCDATE" HeaderText="設定日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="CSDATE" HeaderText="控制起始日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="15%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="CFDATE" HeaderText="控制結束日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="15%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="EXPLAIN" HeaderText="控制說明" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="20%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="YEARS_ROC" HeaderText="截止年度" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="MONTHS_N" HeaderText="截止月份" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="14%">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="btnEDIT1" runat="server" CommandName="edit" CssClass="asp_button_M">修改</asp:LinkButton>
                                        <asp:LinkButton ID="btnDEL1" runat="server" CommandName="del" CssClass="asp_button_M">刪除</asp:LinkButton>

                                        <%--<asp:Button ID="btnEDIT1" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M" AuthType="UPD" />
                                        <asp:Button ID="btnDEL1" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M" AuthType="DEL" />--%>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </td>
                </tr>
            </table>
        </div>

        <div id="div_edit" runat="server">
            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr>
                    <td width="20%" class="bluecol_need">設定日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txCCDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span17" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txCCDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span18" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txCCDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">控制起始：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txCSDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span13" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txCSDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span14" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txCSDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                        <asp:DropDownList ID="ddlCSDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlCSDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">控制結束：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txCFDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span15" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txCFDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span16" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txCFDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                        <asp:DropDownList ID="ddlCFDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlCFDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">截止年度：</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlYEARS" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">截止月份：</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblMONTHS" runat="server" RepeatDirection="Horizontal"></asp:RadioButtonList></td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">審查計分區間：</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlYEARS1" runat="server"></asp:DropDownList>年<asp:DropDownList ID="ddlHALFYEAR1" runat="server"></asp:DropDownList>
                        ～<asp:DropDownList ID="ddlYEARS2" runat="server"></asp:DropDownList>年<asp:DropDownList ID="ddlHALFYEAR2" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">控制說明：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txEXPLAIN" Width="80%" runat="server" placeholder="請輸入 控制說明 訊息" MaxLength="300" Rows="7" TextMode="MultiLine"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="2" valign="middle" align="center">
                        <asp:Label ID="LabISDELETE" runat="server" Text="(資料已刪除)" Visible="False" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_S" AuthType="SAVE"></asp:Button>
                        <asp:Button ID="bt_cancle" runat="server" Text="取消" CausesValidation="False" CssClass="asp_button_S" AuthType="CANCLE"></asp:Button>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol" colspan="2">
                        <asp:Label ID="lab_msg2" runat="server" ForeColor="Red"></asp:Label></td>
                </tr>
            </table>

        </div>

        <asp:HiddenField ID="Hid_OTLID" runat="server" />

    </form>
</body>
</html>
