<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_006_R.aspx.vb" Inherits="WDAIIP.SD_09_006_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開課一覽表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            wopen('../02/SD_02_ch.aspx?RID=' + RID, 'Class', 540, 520, 1);
        }

        function print() {
            if (document.form1.start_date.value == '' || document.form1.end_date.value == '') {
                if (document.form1.start_date2.value == '' || document.form1.end_date2.value == '') {
                    window.alert('日期範圍請擇一挑選');
                    return false;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;開課一覽表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">訓練機構</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                    <input type="button" value="..." id="Button2" name="Button2" runat="server" class="button_b_Mini" />
                    <span id="HistoryList2" style="display: none; position: absolute">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓日期區間</td>
                <td colspan="3" class="whitecol">
                    <span id="span01" runat="server">
                        <asp:TextBox ID="start_date" runat="server" Columns="15" MaxLength="10"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        ~
                        <asp:TextBox ID="end_date" runat="server" Columns="15" MaxLength="10"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />(開訓日期與報名日期請擇一挑選)
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓日期區間
                </td>
                <td colspan="3" class="whitecol">
                    <span id="span02" runat="server">
                        <asp:TextBox ID="start_date2" runat="server" Columns="15" MaxLength="10"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
                        <asp:TextBox ID="end_date2" runat="server" Columns="15" MaxLength="10"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">查詢結果</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButton ID="RadioButton1" runat="server" Text="包含下層單位"></asp:RadioButton></td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="whitecol" align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="btnExp1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    <%--<div style="width: 100%" align="center" class="whitecol"></div>--%>
                </td>
            </tr>
        </table>
        <div align="center">
            <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
        </div>
        <br />
        <%--  <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1" width="100%"></table>--%>
    </form>
</body>
</html>
