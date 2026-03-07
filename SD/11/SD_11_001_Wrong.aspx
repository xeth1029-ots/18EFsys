<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_001_Wrong.aspx.vb" Inherits="WDAIIP.SD_11_001_Wrong" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>訓練期末學員滿意度</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
					首頁&gt;&gt;學員動態管理&gt;&gt;就業輔導問卷&gt;&gt;訓練期末學員滿意度
                    </asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%"
                        CssClass="font" AllowPaging="True">
                        <AlternatingItemStyle BackColor="#f5f5f5" />
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆錯誤"></asp:BoundColumn>
                            <asp:BoundColumn DataField="FillFormDate" HeaderText="讀卡日期"></asp:BoundColumn>
                            <asp:BoundColumn DataField="StudID" HeaderText="學號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OCID" HeaderText="開班編號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Result" HeaderText="狀態"></asp:BoundColumn>
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

    </form>
</body>
</html>
