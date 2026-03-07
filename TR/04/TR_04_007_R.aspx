<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_007_R.aspx.vb" Inherits="WDAIIP.TR_04_007_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_04_007_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script>
        function GetMode() {
            document.form1.center.value = '';
            document.form1.RIDValue.value = '';
            document.form1.OCIDValue.value = '';
            for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
                document.form1.OCID.options[i] = null;
            }
            document.form1.OCID.options[0] = new Option('請選擇機構');

            if (document.form1.DistID.selectedIndex != 0 && document.form1.TPlanID.selectedIndex != 0) {
                document.form1.Button3.disabled = false;
            }
            else {
                document.form1.Button3.disabled = true;
            }
        }
        function search() {
            //if (document.form1.Syear.value == '') {
            //	alert("請選擇列印年度!");
            //	return false;
            //}
            if (document.form1.DistID.value == '') {
                //alert("請選擇轄區中心!");
                alert("請選擇轄區分署!");
                return false;
            }
            if (document.form1.TPlanID.value == '') {
                alert("請選擇訓練計畫!");
                return false;
            }

            var msg = '';

            if (document.form1.STDate1.value != '') {
                if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
            }
            if (document.form1.STDate2.value != '') {
                if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的迄日不是正確的日期格式\n';
            }
            if (document.form1.FTDate1.value != '') {
                if (!IsDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
            }
            if (document.form1.FTDate2.value != '') {
                if (!IsDate(document.form1.FTDate2.value)) msg += '結訓日期的迄日不是正確的日期格式\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }

        }

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }

        function showHide(type) {
            if (type == 1) {
                document.all.nxlayer_01.style.visibility = 'visible'
                document.all.Syear.style.visibility = 'hidden'
                document.all.DistID.style.visibility = 'hidden'
                document.all.OCID.style.visibility = 'hidden'
                document.all.TPlanID.style.visibility = 'hidden'
            } else {
                document.all.nxlayer_01.style.visibility = 'hidden'
                document.all.Syear.style.visibility = 'visible'
                document.all.DistID.style.visibility = 'visible'
                document.all.OCID.style.visibility = 'visible'
                document.all.TPlanID.style.visibility = 'visible'
            }
        }
						
    </script>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <%--
                --%>
                <table class="font" id="nxlayer_01" style="border-right: #9eb5cd 1px solid; border-top: #9eb5cd 1px solid; filter: progid: DXImageTransform.Microsoft.Shadow(Color=#919899, Strength=4, Direction=135); visibility: hidden; border-left: #9eb5cd 1px solid; border-bottom: #9eb5cd 1px solid; position: absolute" cellspacing="0" cellpadding="0" width="100%" border="0">
                    <tbody>
                        <tr>
                            <td width="90%" bgcolor="#ffffff">
                                首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">職業訓練情形累計週報表</font> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a onclick="showHide(0)" href="#"><font color="black">關閉</font></a>
                            </td>
                        </tr>
                        <tr>
                            <td class="dashline" style="height: 1px" height="1">
                                <u></u>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding-right: 8px; padding-left: 8px; padding-bottom: 6px; padding-top: 8px" width="100%" bgcolor="#f1faff" colspan="2">
                                說明：<br>
                                若輸入開訓期間:2006/01/01~2006/12/31 則會查出2006年開訓的班級 (含2006年開訓但結訓日跨2007年的班級)
                                <br>
                                若輸入結訓期間:2006/01/01~2006/12/31 則會查出2006年結訓的班級 (含2005年開訓,但是結訓日期是在2006年的班級)
                                <br>
                                若輸入開訓期間:2006/01/01~ 不填
                                <br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 結訓期間: &nbsp;&nbsp;&nbsp; 不填 &nbsp; &nbsp; ~ 2006/12/31 則會查出在2006年開訓及結訓的班級 (不含跨年度)
                                <br>
                                以上情況的前提是:[年度] 需選擇 [====請選擇====]
                            </td>
                        </tr>
                    </tbody>
                </table>
                <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td title="點選可以查看說明" style="cursor: pointer" onclick="showHide(1)">
                            首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">職業訓練情形累計週報表</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol">
                            年度
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:DropDownList ID="Syear" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <%--<td class="bluecol_need">轄區中心</td>--%>
                        <td class="bluecol_need">轄區分署</td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:DropDownList ID="DistID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            訓練計畫
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:DropDownList ID="TPlanID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            訓練機構
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:TextBox ID="center" runat="server"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server"><input id="Button3" onclick="javascript:wopen('../../Common/MainOrg.aspx?DistID='+document.form1.DistID.value+'&amp;TPlanID='+document.form1.TPlanID.value,'訓練機構',400,400,1)" type="button" value="..." name="Button3" runat="server" class="button_b_Mini">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            班別
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:DropDownList ID="OCID" runat="server">
                            </asp:DropDownList>
                            <input id="OCIDValue" type="hidden" name="OCIDValue" runat="server">
                            <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>(當有指定班別時，系統將會忽略開訓期間)
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            班級範圍
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:RadioButtonList ID="CPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                <asp:ListItem Value="1" Selected="True">全部</asp:ListItem>
                                <asp:ListItem Value="2">已開訓</asp:ListItem>
                                <asp:ListItem Value="3">已結訓</asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            開訓期間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
                                <asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            結訓期間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                        </td>
                    </tr>
                </table>
                <table width="100%">
                    <tr>
                        <td align="center" colspan="4">
                            <asp:Panel ID="Panel1" runat="server" HorizontalAlign="Left" CssClass="font">
                                說明：<br>
                                班級範圍選擇[已開訓]，則是將班級的開訓日在今日(含)之前的給顯示出來<br>
                                班級範圍選擇[已結訓]，則是將班級的結訓日在今日(含)之前的給顯示出來</asp:Panel>
                            <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="4">
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
