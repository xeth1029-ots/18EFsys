<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_034_R.aspx.vb" Inherits="WDAIIP.SD_05_034_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>屆退官兵出缺勤統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        /*
		function GETvalue() {
		document.getElementById('Button3').click();
		}
		function choose_class() {
		//onclick="javascript:openClass('../02/SD_02_ch.aspx?special=5&RID='+document.form1.RIDValue.value);"
		//special=5 提供開訓、結訓日期欄位special=5&DateS=start_date&DateF=end_date
		var vRIDValue = document.form1.RIDValue.value;
		var sUrl = '../02/SD_02_ch.aspx?special=5&DateS=start_date&DateF=end_date&RID=' + vRIDValue;
		openClass(sUrl);
		}
		*/

        function print() {
            var msg = '';
            var RblPrintType1 = document.form1.RblPrintType1;
            var STDate1 = document.form1.STDate1;
            var STDate2 = document.form1.STDate2;
            if (STDate1.value == '' && STDate2.value == '') msg += '請輸入時間區間\n';
            if (!isChecked(RblPrintType1)) msg += '請選擇列印方式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;屆退官兵出缺勤統計表</asp:Label>--%>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;志願役官兵出缺勤統計表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol_need" width="20%">訓練計畫 </td>
                            <td class="whitecol"><asp:CheckBoxList ID="cblTPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="hbtnOrg" type="button" value="..." runat="server">
                                <input id="RIDValue" type="hidden" runat="server">
                                <span id="HistoryList2" style="display: none; position: absolute"><asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">時間區間 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Width="18%" onfocus="this.blur()"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"> ～
                                <asp:TextBox ID="STDate2" runat="server" Width="18%" onfocus="this.blur()"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">列印格式 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RblPrintType1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">出缺勤明細表</asp:ListItem>
                                    <asp:ListItem Value="2">請假、缺曠課累計時數統計表</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center"><asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>