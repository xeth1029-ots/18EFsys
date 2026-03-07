<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RWB_01_005.aspx.vb" Inherits="WDAIIP.RWB_01_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>隱私權及安全政策</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <%--<script type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <%--<script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181019
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);
    </script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;報名網維護&gt;&gt;隱私權及安全政策</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol">上稿日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="schC_CDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schC_CDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= schC_CDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schC_CDATE2" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span3" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schC_CDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span4" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= schC_CDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">上架日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="schC_SDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span5" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schC_SDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span6" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= schC_SDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schC_SDATE2" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span7" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schC_SDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span8" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= schC_SDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">內容關鍵字：</td>
                <td class="whitecol">
                    <asp:TextBox ID="schCONTENT1" Width="80%" runat="server"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">停用日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="schC_EDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span9" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schC_EDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span10" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= schC_EDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schC_EDATE2" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span11" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schC_EDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span12" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= schC_EDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
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
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
        <table id="tb_Sch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="ROWNUM" HeaderText="序號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="CCDATE" HeaderText="上稿日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="CSDATE" HeaderText="上架日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="CEDATE" HeaderText="停用日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="CCONTENT1" HeaderText="內容" ItemStyle-Width="50%">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="14%">
                                <ItemTemplate>
                                    <asp:Button ID="btnEDIT1" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M" AuthType="UPD" />
                                    <asp:Button ID="btnDEL1" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M" AuthType="DEL" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <input type="hidden" runat="server" id="hid_V" />
    </form>
</body>
</html>
