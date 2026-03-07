<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_005.aspx.vb" Inherits="WDAIIP.SD_08_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_08_005</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript">
        //全選判斷
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

        //查詢範圍判斷
        function Enabled_OCID(orgname, Rid, Planid) {
            if (document.form1.chkData.checked == true) {
                document.getElementById("center").value = orgname;
                document.getElementById("RIDValue").value = Rid;
                document.getElementById("PlanID").value = Planid;
                document.getElementById("Button2").disabled = true;
                document.getElementById("Class_TR").style.display = 'none';
                document.getElementById("Org_TR").style.display = 'none';

            } else {
                document.getElementById("Button2").disabled = false;
                document.getElementById("Org_TR").style.display = '';//  'inline';
                document.getElementById("btnSchClass").click();
            }
        }

        function chkPrt() {
            var msg = '';
            var ddlYear = document.form1.ddlYear;
            var center = document.form1.center;
            var txtSCSDate = document.form1.txtSCSDate;
            var txtECSDate = document.form1.txtECSDate;
            var txtSTEDate = document.form1.txtSTEDate;
            var txtETEDate = document.form1.txtETEDate;

            if (ddlYear.value == '') msg += "請選擇年度!\n";
            if (center.value == '') msg += '請選擇訓練機構\n';

            if (txtSCSDate.value != '') {
                if (!checkDate(txtSCSDate.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
                else if (txtSCSDate.value.substring(0, 4) != ddlYear.value) msg += '開訓日期的起始日年度與所選年度要相同\n';
            }

            if (txtECSDate.value != '') {
                if (!checkDate(txtECSDate.value)) msg += '開訓日期的迄日不是正確的日期格式\n';
                else if (txtECSDate.value.substring(0, 4) != ddlYear.value) msg += '開訓日期的迄日年度與所選年度要相同\n';
            }

            if (checkDate(txtSCSDate.value) && checkDate(txtECSDate.value)) {
                if (txtSCSDate.value > txtECSDate.value) msg += '開訓迄日不可大於開訓起日!\n';
            }

            if (txtSTEDate.value != '') {
                if (!checkDate(txtSTEDate.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
                else if (txtSTEDate.value.substring(0, 4) != ddlYear.value) msg += '結訓日期的起始日年度與所選年度要相同\n';
            }

            if (txtETEDate.value != '') {
                if (!checkDate(txtETEDate.value)) msg += '結訓日期的迄日不是正確的日期格式\n';
                else if (txtETEDate.value.substring(0, 4) != ddlYear.value) msg += '結訓日期的迄日年度與所選年度要相同\n';
            }

            if (checkDate(txtSTEDate.value) && checkDate(txtETEDate.value)) {
                if (txtSTEDate.value > txtETEDate.value) msg += '結訓迄日不可大於開訓起日!\n';
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
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td align="center">
                    <table class="font" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;學員動態管理&gt;&gt;職業訓練生活津貼&gt;&gt;<font color="#990000">離退訓繳回狀況表</font></td>
                        </tr>
                    </table>
                    <table cellspacing="1" cellpadding="1" width="100%" border="0" class="table_nw">
                        <tr id="Year_TR" runat="server">
                            <td class="bluecol_need" width="100">年度</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlYear" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="DistID_TR" runat="server">
                            <td class="bluecol_need">轄區</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cklDistID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <input id="hidDistID" type="hidden" value="0" runat="server" /></td>
                        </tr>
                        <tr id="PlanID_TR" runat="server">
                            <td class="bluecol_need">訓練計畫</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cklTPlanID" CssClass="font" runat="server" RepeatDirection="Horizontal" RepeatColumns="3"
                                    CellPadding="0" CellSpacing="0">
                                </asp:CheckBoxList>
                                <input id="hidTPlanID" type="hidden" value="0" runat="server" size="1"></td>
                        </tr>
                        <tr id="Check_TR" runat="server">
                            <td class="bluecol">查詢範圍</td>
                            <td class="whitecol">
                                <input id="chkData" type="checkbox" runat="server" />全訓練機構</td>
                        </tr>
                        <tr id="Org_TR" runat="server">
                            <td class="bluecol" width="100">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="PlanID" type="hidden" name="PlanID" runat="server" />
                                <asp:Button ID="btnSchClass" runat="server" Text="查詢班級" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                        <tr id="Class_TR" runat="server">
                            <td class="bluecol">班別</td>
                            <td class="whitecol">
                                <asp:ListBox ID="lsbOCID" runat="server" Width="100%" SelectionMode="Multiple" Rows="6"></asp:ListBox>
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label><br />
                                (按Ctrl可以複選班級)
									<asp:TextBox ID="txtTMID" runat="server" Visible="False"></asp:TextBox>
                                <asp:TextBox ID="txtOCID" runat="server" Visible="False"></asp:TextBox>
                                <input id="hidTMID" type="hidden" runat="server" />
                                <input id="hidOCID" type="hidden" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtSCSDate" runat="server" Columns="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('txtSCSDate','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                ~
									<asp:TextBox ID="txtECSDate" runat="server" Columns="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('txtECSDate','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtSTEDate" runat="server" Columns="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('txtSTEDate','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                ~
									<asp:TextBox ID="txtETEDate" runat="server" Columns="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('txtETEDate','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="btnPrt" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
