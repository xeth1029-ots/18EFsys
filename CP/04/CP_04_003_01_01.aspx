<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_003_01_01.aspx.vb" Inherits="WDAIIP.CP_04_003_01_01" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員參訓歷史</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        body { overflow-y: auto; }
    </style>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="Stud_DG" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Visible="False" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區中心"></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="DistName" HeaderText="轄區分署"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="TrinUnit" HeaderText="培訓單位"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ClassName" HeaderText="班別"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Sex" HeaderText="性別"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Ident" HeaderText="身分別"></asp:BoundColumn>
                            <%--
                            <asp:BoundColumn Visible="False" DataField="source" HeaderText="來源"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="Stdid" HeaderText="序號"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="DistID" HeaderText="DistID"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="TPlanID" HeaderText="TPlanID"></asp:BoundColumn>
                            --%>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <%--<tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>--%>
            <tr>
                <td>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></div>
                </td>
            </tr>
             <tr>
                <td align="center">
                    <input id="btnClose2" type="button" value="關閉視窗" name="Button1" runat="server" class="button_b_M"></td>
            </tr>
        </table>
    </form>
</body>
</html>
