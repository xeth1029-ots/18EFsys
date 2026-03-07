<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_008.aspx.vb" Inherits="WDAIIP.CM_03_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>離退訓人數統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">

        function ClearData() {
            var PlanID = document.getElementById('PlanID');
            var center = document.getElementById('center');
            var RIDValue = document.getElementById('RIDValue');
            var OCID = document.getElementById('OCID');
            var msg = document.getElementById('msg');

            PlanID.value = '';
            center.value = '';
            RIDValue.value = '';
            for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
                document.form1.OCID.options[i] = null;
            }
            OCID.style.display = 'none';
            msg.innerHTML = '請先選擇機構';
        }

        //檢查列印條件為
        function CheckPrint() {
            var msg = '';
            var STDate1 = document.getElementById('STDate1');
            var STDate2 = document.getElementById('STDate2');
            var FTDate1 = document.getElementById('FTDate1');
            var FTDate2 = document.getElementById('FTDate2');

            if (STDate1.value != '') {
                if (!checkDate(STDate1.value)) msg += '[開訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (STDate2.value != '') {
                if (!checkDate(STDate2.value)) msg += '[開訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (FTDate1.value != '') {
                if (!checkDate(FTDate1.value)) msg += '[結訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (FTDate2.value != '') {
                if (!checkDate(FTDate2.value)) msg += '[結訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (STDate2.value != '' && STDate1.value != '' && STDate2.value < STDate1.value) {
                msg += '[開訓區間的迄日]必需大於[開訓區間的起日]\n';
            }
            if (FTDate2.value != '' && FTDate1.value != '' && FTDate2.value < FTDate1.value) {
                msg += '[結訓區間的迄日]必需大於[結訓區間的起日]\n';
            }

            var Identity1 = getCheckBoxListValue('Identity');
            var DistID1 = getCheckBoxListValue('DistID');
            var TPlanID1 = getCheckBoxListValue('TPlanID');

            if (parseInt(DistID1) == 0) { msg += '請選擇轄區\n'; }
            if (parseInt(TPlanID1) == 0) { msg += '請選擇計畫\n'; }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;綜合動態報表</asp:Label>
                    <%--   首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;離退訓人數統計表--%>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">動態報表 </td>
                <td class="whitecol" width="80%">
                    <uc1:WUC2 runat="server" ID="WUC2" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">年度 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Syear" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓區間 </td>
                <td class="whitecol">
                    <asp:TextBox ID="STDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    ~<asp:TextBox ID="STDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓區間 </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="FTDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    ~<asp:TextBox ID="FTDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">轄區 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                    </asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server" size="1">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練計畫 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
                    </asp:CheckBoxList>
                    <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server" size="1">
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練機構 </td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini"><input id="RIDValue" type="hidden" name="RIDValue" runat="server"><input id="PlanID" type="hidden" name="PlanID" runat="server">
                    <asp:Button ID="Button3" runat="server" Text="查詢班級" CssClass="asp_Export_M"></asp:Button><br>
                    (勾選班級後會省略[年度]、[開訓區間]、[結訓區間]的條件) </td>
            </tr>
            <tr>
                <td class="bluecol">班別 </td>
                <td class="whitecol">
                    <asp:ListBox ID="OCID" runat="server" Width="300px" SelectionMode="Multiple" Rows="6"></asp:ListBox>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>(按Ctrl可以複選班級) </td>
            </tr>
            <tr id="IdentityTR" runat="server">
                <td class="bluecol">身分別 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="Identity" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="Identity_List" type="hidden" value="0" name="Identity_List" runat="server" size="1">
                </td>
            </tr>
            <tr>
                <td class="bluecol">預算別 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="BudID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">查詢方式 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="searcha_type1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">統計資料</asp:ListItem>
                        <asp:ListItem Value="2">明細資料</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;
					<asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>

                </td>
            </tr>
        </table>

        <p align="center">
            <asp:Label ID="ExportMsg" runat="server" ForeColor="Red"></asp:Label>
        </p>
        <div id="Div1" runat="server">
            <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%">
                <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
            </asp:DataGrid>
        </div>


    </form>
</body>
</html>
