<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_02_017_R.aspx.vb" Inherits="WDAIIP.CP_02_017_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>分署訓練人數按訓練職類分</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">
        function search() {
            var msg = '';
            if (isEmpty(document.form1.start_date) && isEmpty(document.form1.end_date)) {
                msg += '請選擇日期範圍!\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
        <tr>
            <td>
                <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;<font color="#990000">分署訓練人數按訓練職類分</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td bgcolor="#cc6666">
                            <font color="#ffffff">&nbsp;&nbsp;&nbsp; 結訓日期</font><font color="#ffff80">*</font>
                        </td>
                        <td bgcolor="#ffecec">
                            <asp:TextBox ID="start_date" runat="server" Width="100px"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
                            ～<asp:TextBox ID="end_date" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer"
                                onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
                        </td>
                    </tr>
                </table>
                <p align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
