<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_004.aspx.vb" Inherits="WDAIIP.SD_15_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_15_004</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function OpenOrg(vTPlanID) {
            if (document.getElementById('DistID').selectedIndex == 0) {
                alert('請先選擇轄區');
                return false;
            }
            else {
                wopen('../../common/MainOrg.aspx?DistID=' + document.getElementById('DistID').value + '&TPlanID=' + vTPlanID, '', 400, 400, 'yes');
            }
        }
        function CheckPrint() {
            var FTDate1 = document.getElementById('FTDate1').value;
            var FTDate2 = document.getElementById('FTDate2').value;
            var DistID = document.getElementById('DistID').value;
            var PlanID = document.getElementById('PlanID').value;
            var RID = document.getElementById('RIDValue').value;
            var SearchPlan = getRadioValue(document.form1.SearchPlan);

            var SearchPlan1 = '';
            if (SearchPlan == 'W' || SearchPlan == 'G') {
                SearchPlan1 = SearchPlan;
            }

            //var SearchPlan=document.getElementById('SearchPlan').value;

            var msg = '';
            if (!checkDate(FTDate1) && FTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(FTDate2) && FTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;訓練人時成本分析表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td>
                    <%--<table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<FONT color="#990000">訓練人時成本分析表</FONT>
                            </asp:Label>
                        </td>
                    </tr>
                </table>--%>
                    <table id="Table3" class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">結訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
                            <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DistID" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="Button1" type="button" value="..." name="Button1" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" runat="server">
                                <input type="hidden" id="PlanID" name="PlanID" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印依據
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PrintMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">轄區</asp:ListItem>
                                    <asp:ListItem Value="2">機構</asp:ListItem>
                                    <asp:ListItem Value="3">訓練職類</asp:ListItem>
                                    <asp:ListItem Value="4">課程類別</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol">計畫範圍
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol">包班種類</td>
                            <td class="whitecol" colspan="4">
                                <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal"
                                    RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
