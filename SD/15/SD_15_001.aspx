<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_001.aspx.vb" Inherits="WDAIIP.SD_15_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練人數統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        var cst_inline1 = "";

        //檢查列印依劇若是訓練單位則顯示訓練單位選項
        function ChangePrintMode() {
            var PrintMode = document.getElementsByName('PrintMode');
            var DistID_TR = document.getElementById('DistID_TR');
            var Org_TR = document.getElementById('Org_TR');

            DistID_TR.style.display = 'none';
            Org_TR.style.display = 'none';
            if (getRadioValue(PrintMode) == '5') {
                DistID_TR.style.display = cst_inline1;
                Org_TR.style.display = cst_inline1;
            }
        }

        function ChangeMode() {
            //var SearchMode = document.getElementsByName('SearchMode');
            var vSearchMode = $('#<%= SearchMode.ClientID %> input:checked').val();
            $('#mode1_1').hide();
            $('#mode1_2').hide();
            $('#mode2_1').show();
            $('#mode2_2').show();
            //var mode1_1 = document.getElementById('mode1_1');
            //var mode1_2 = document.getElementById('mode1_2');
            //var mode2_1 = document.getElementById('mode2_1');
            //var mode2_2 = document.getElementById('mode2_2');
            //mode1_1.style.display = 'none';
            //mode1_2.style.display = 'none';
            //mode2_1.style.display = cst_inline1;
            //mode2_2.style.display = cst_inline1;
            //alert('alert--:' + getRadioValue(SearchMode));
            //getRadioValue(SearchMode) == '1'
            if (vSearchMode == '1') {
                $('#mode1_1').show();
                $('#mode1_2').show();
                $('#mode2_1').hide();
                $('#mode2_2').hide();
                //mode1_1.style.display = cst_inline1;
                //mode1_2.style.display = cst_inline1;
                //mode2_1.style.display = 'none';
                //mode2_2.style.display = 'none';
            }
        }

        function OpenOrg(vTPlanID) {
            var DistID = document.getElementById('DistID');
            if (DistID.selectedIndex == 0) {
                alert('請先選擇轄區');
                return false;
            }
            wopen('../../common/MainOrg.aspx?DistID=' + DistID.value + '&TPlanID=' + vTPlanID, '', 400, 400, 'yes');
        }

        function CheckPrint() {
            var PrintMode = document.getElementsByName('PrintMode');
            var SearchMode = document.getElementsByName('SearchMode');
            var OrgKind2 = document.getElementsByName('OrgKind2');
            var vRIDValue = document.getElementById('RIDValue').value;
            var RIDValue = '';
            var OrgKind = '';
            var msg = '';

            if (getRadioValue(SearchMode) == '1') {
                if (vRIDValue == 'A') { RIDValue = ''; }
                if (vRIDValue != 'A') { RIDValue = vRIDValue; }

                var grvOrgKind2 = getRadioValue(OrgKind2);
                if (grvOrgKind2 == 'A') { OrgKind = ''; }
                if (grvOrgKind2 != 'A') { OrgKind = grvOrgKind2; }
                var STDate1 = document.getElementById('STDate1').value;
                var STDate2 = document.getElementById('STDate2').value;
                var FTDate1 = document.getElementById('FTDate1').value;
                var FTDate2 = document.getElementById('FTDate2').value;

                if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
                if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';
                if (!checkDate(FTDate1) && FTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
                if (!checkDate(FTDate2) && FTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';

                if (getRadioValue(PrintMode) == '5') {
                    if (vRIDValue == '' && document.getElementById('DistID').selectedIndex != 0)
                    { msg += '請選擇機構\n'; }
                }

                if (msg != '') {
                    alert(msg);
                    return false;
                }
            }

            if (getRadioValue(SearchMode) != '1') {
                if (getRadioValue(PrintMode) == '5') {
                    if (vRIDValue == '' && document.getElementById('DistID').selectedIndex != 0)
                    { msg += '請選擇機構\n'; }
                }
                if (document.getElementById('Years').selectedIndex == 0) {
                    msg += '請選擇年度\n';
                }
                if (msg != '') {
                    alert(msg);
                    return false;
                }
                if (msg == '') {
                    if (vRIDValue == 'A') { RIDValue = ''; }
                    if (vRIDValue != 'A') { RIDValue = vRIDValue; }

                    var grvOrgKind2 = getRadioValue(OrgKind2);
                    if (grvOrgKind2 != 'A') { OrgKind = grvOrgKind2; }
                    if (grvOrgKind2 == 'A') { OrgKind = ''; }

                    switch (getRadioValue(PrintMode)) {
                        case '2':
                            alert('找不到對應的報表');
                            return false;
                            break;
                    }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;綜合動態報表</asp:Label>
                    <%--首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<FONT color="#990000">訓練人數統計表</FONT>--%>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <table id="Table2" class="table_sch" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">動態報表 </td>
                            <td class="whitecol" width="80%">
                                <uc1:WUC1 runat="server" ID="WUC1" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">查詢模式
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="SearchMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1" Selected="True">依照訓練期間查詢</asp:ListItem>
                                    <asp:ListItem Value="2">依照年度月份查詢</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="mode1_1">
                            <td class="bluecol">開訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">～
								<asp:TextBox ID="STDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="mode1_2">
                            <td class="bluecol">結訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">～
								<asp:TextBox ID="FTDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="mode2_1" style="display: none">
                            <td class="bluecol">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Years" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="mode2_2" style="display: none">
                            <td class="bluecol">月份
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Months" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol">計畫範圍
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="OrgKind2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol">包班種類
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PackageType" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印依據
                            </td>
                            <td class="whitecol">
                                <%--<asp:ListItem Value="2">訓練類別</asp:ListItem>--%>
                                <%--<asp:ListItem Value="11 ">新南向政策</asp:ListItem>--%>
                                <asp:RadioButtonList ID="PrintMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" AutoPostBack="True" RepeatColumns="7">
                                    <asp:ListItem Value="1" Selected="True">轄區</asp:ListItem>
                                    <asp:ListItem Value="3">職能別</asp:ListItem>
                                    <asp:ListItem Value="4">訓練單位類別</asp:ListItem>
                                    <asp:ListItem Value="5">訓練單位</asp:ListItem>

                                    <asp:ListItem Value="6">課程分類</asp:ListItem>
                                    <asp:ListItem Value="7">生產力4.0</asp:ListItem>
                                    <asp:ListItem Value="8">新興產業</asp:ListItem>
                                    <asp:ListItem Value="9">重點服務業</asp:ListItem>
                                    <asp:ListItem Value="10">新興智慧型產業</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="DistID_TR" runat="server">
                            <td class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DistID" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="Org_TR" runat="server">
                            <td class="bluecol">機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
                                <input id="Button1" value="..." type="button" name="Button1" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="PlanID" type="hidden" name="PlanID" runat="server">
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
