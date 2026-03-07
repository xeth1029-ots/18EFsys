<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_003_add.aspx.vb"
    Inherits="WDAIIP.OB_01_003_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>OB_01_003_add</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="JavaScript">
			
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">會議日期及地點查詢建檔</font></FONT>
                            </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" id="TableLay2" cellspacing="1" cellpadding="1">
                    <tbody>
                        <tr>
                            <td width="100" class="bluecol_need">
                                年度
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="ddl_years" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                標案名稱
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="ddlTenderCName" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                會議主題
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="MTSubject" runat="server" Width="300px" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                會議日期
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="MTDate" Width="80" MaxLength="10" runat="server"></asp:TextBox><img
                                    style="cursor: pointer" onclick="javascript:show_calendar('<%= MTDate.ClientId %>','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                會議地點
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="MTPlace" runat="server" Width="150px" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                會議議程內容
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="MTContent" runat="server" Width="300px" TextMode="MultiLine" Rows="3"></asp:TextBox>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <p align="center">
                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button><font
                        face="新細明體">&nbsp;
                        <asp:Button ID="btnReturn" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button></font></p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
