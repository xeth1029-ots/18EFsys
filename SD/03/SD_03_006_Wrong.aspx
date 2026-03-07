<%@ Register TagPrefix="uc1" TagName="PageControler" Src="../../PageControler.ascx" %>

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_006_Wrong.aspx.vb" Inherits="TIMS.SD_03_006_Wrong" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>錯誤資料</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" AllowPaging="True">
                        <AlternatingItemStyle BackColor="#f5f5f5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆錯誤"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因"></asp:BoundColumn>
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
    </font>
    </form>
</body>
</html>
