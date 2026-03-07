<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_008_R.aspx.vb" Inherits="WDAIIP.TR_05_008_R" %>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練計畫特定對象人數統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">

        function GetOrg() {
            var msg = '';
            var DistID = getRadioValue(document.getElementsByName('DistID'));
            var TPlanID = getRadioValue(document.getElementsByName('TPlanID'));
            if (DistID == '') msg += '請先選擇轄區\n';
            if (TPlanID == '') msg += '請先選擇訓練計畫\n';

            if (msg != '') {
                alert(msg);
            }
            else {
                wopen('../../Common/MainOrg.aspx?DistID=' + DistID + '&TPlanID=' + TPlanID + '&BtnName=Button3', '查詢機構', 400, 400, 1);
            }
        }
        function ClearData() {
            document.getElementById('PlanID').value = '';
            document.getElementById('center').value = '';
            document.getElementById('RIDValue').value = '';
            for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
                document.form1.OCID.options[i] = null;
            }
            document.getElementById('OCID').style.display = 'none';
            document.getElementById('msg').innerHTML = '請先選擇機構';

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


        function search() {

            var msg = '';
            if (isEmpty(document.form1.start_date) && isEmpty(document.form1.end_date)) msg += '請選擇日期範圍!\n';
            /*if(!isChecked(document.getElementsByName('TPlanID'))) msg+='請選擇訓練計畫\n';*/

            if (msg != '') {
                alert(msg);
                return false;
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
                    <%--首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">訓練計畫特定對象人數統計表</FONT>--%>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">報表總類 </td>
                <td class="whitecol" width="80%">
                    <uc1:WUC2 runat="server" ID="WUC2" />
                </td>
            </tr>

            <tr>
                <td class="bluecol">開訓期間</td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="start_date" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('start_date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~<asp:TextBox ID="end_date" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('end_date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓區間</td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~<asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
            </tr>
            <tr>
                <td class="bluecol">轄區</td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" CssClass="font" RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練計畫</td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="77%"></asp:TextBox>
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" />
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                    <input id="PlanID" type="hidden" name="PlanID" runat="server" />
                    <asp:Button ID="Button3" runat="server" Text="查詢班級"></asp:Button>(勾選班級後會省略結訓日期的條件)
                </td>
            </tr>
            <tr>
                <td class="bluecol">班別</td>
                <td class="whitecol">
                    <asp:ListBox ID="OCID" runat="server" Rows="6" SelectionMode="Multiple" Width="300px"></asp:ListBox>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>(按Ctrl可以複選班級)
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
                <td colspan="2" align="center" class="whitecol">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="Export1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
