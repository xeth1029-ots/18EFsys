<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_008_c.aspx.vb" Inherits="WDAIIP.SD_04_008_c" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>資料處理中</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <meta http-equiv="refresh" content="20">
    <link href="../../style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" border="0" class="font" width="170">
            <tr>
                <td align="center"><img alt="" src="../../images/loading.gif" width="94" height="17"></td>
            </tr>
            <tr>
                <td align="center">資料處理中，目前進度<asp:Label ID="Percent" runat="server"></asp:Label>%<br>
                    <input type="button" value="重新整理" onclick="location.reload();"><input type="button" value="關閉視窗" onclick="window.close();"></td>
            </tr>
        </table>
    </form>
</body>
</html>
