<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_002_Wrong.aspx.vb" Inherits="WDAIIP.SD_04_002_Wrong" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>錯誤資料</title>
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
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" AllowPaging="True">
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆錯誤">
                                <ItemStyle HorizontalAlign="Center" Width="10%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SchoolDate" HeaderText="排課日期">
                                <ItemStyle HorizontalAlign="Center" Width="16%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OCID" HeaderText="開班編號">
                                <ItemStyle HorizontalAlign="Center" Width="24%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因">
                                <ItemStyle HorizontalAlign="Center" Width="50%" />
                            </asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
