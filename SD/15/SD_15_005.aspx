<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_005.aspx.vb" Inherits="WDAIIP.SD_15_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_15_005</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
            var STDate1 = document.getElementById('STDate1').value;
            var STDate2 = document.getElementById('STDate2').value;
            var DistID = document.getElementById('DistID').value;
            var PlanID = document.getElementById('PlanID').value;
            var RID = document.getElementById('RIDValue').value;

            var msg = '';
            if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }

            switch (getRadioValue(document.form1.PrintMode)) {
                case '1':
                    openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_005_1&path=TIMS&STDate1=' + STDate1 + '&STDate2=' + STDate2 + '&DistID=' + DistID + '&PlanID=' + PlanID + '&RID=' + RID);
                    break;
                case '2':
                    openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_005_2&path=TIMS&STDate1=' + STDate1 + '&STDate2=' + STDate2 + '&DistID=' + DistID + '&PlanID=' + PlanID + '&RID=' + RID);
                    break;
                case '3':
                    openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_005_3&path=TIMS&STDate1=' + STDate1 + '&STDate2=' + STDate2 + '&DistID=' + DistID + '&PlanID=' + PlanID + '&RID=' + RID);
                    break;
                case '4':
                    openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_005_4&path=TIMS&STDate1=' + STDate1 + '&STDate2=' + STDate2 + '&DistID=' + DistID + '&PlanID=' + PlanID + '&RID=' + RID);
                    break;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;課程開班人數分佈圖表
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="Table3" class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">開訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
                            <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
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
                                <input id="RIDValue" type="hidden" runat="server" name="RIDValue">
                                <input type="hidden" id="PlanID" name="PlanID" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印依據
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PrintMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">轄區</asp:ListItem>
                                    <asp:ListItem Value="2">縣市</asp:ListItem>
                                    <asp:ListItem Value="3">訓練職類</asp:ListItem>
                                    <asp:ListItem Value="4">課程類別</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <input type="button" value="列印" onclick="CheckPrint();" class="asp_Export_M">
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
