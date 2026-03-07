<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_024.aspx.vb" Inherits="WDAIIP.SD_15_024" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>綜合動態報表-交叉分析統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            //var smsg1 = "num:" + num + ", myallcheck:" + myallcheck; alert(smsg1);return false;
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%-- <table id="FrameTable2" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr id="tr_rblDYNAMIC1" runat="server">
                <td class="bluecol" width="20%">動態報表</td>
                <td class="whitecol" width="80%">
                    <asp:RadioButtonList ID="rblDYNAMIC1" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1"></asp:RadioButtonList>
                </td>
            </tr>
        </table>--%>
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;綜合動態報表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable2" border="0" cellspacing="1" cellpadding="1" width="100%">
        </table>
        <table id="tb_CM_03_011" cellspacing="1" cellpadding="1" width="100%" border="0">
            <%--style="display: none">--%>
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">動態報表 </td>
                            <td class="whitecol" width="80%">
                                <uc1:WUC2 runat="server" ID="WUC2" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">計畫年度</td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="ddlYear" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">轄區</td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="DistID" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">開訓期間</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
								<asp:TextBox ID="STDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">結訓期間</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
								<asp:TextBox ID="FTDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">計畫範圍</td>
                            <td class="whitecol" colspan="4" width="80%">
                                <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3" CellSpacing="0" CellPadding="0"></asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="46%"></asp:TextBox>
                                <input id="Button3" type="button" value="..." name="Button1" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="PlanID" type="hidden" name="PlanID" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">統計範圍</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="StudStatus" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="1" CellPadding="1">
                                    <asp:ListItem Value="11">報名人數</asp:ListItem>
                                    <asp:ListItem Value="12">開訓人數</asp:ListItem>
                                    <asp:ListItem Value="13">結訓人數</asp:ListItem>
                                    <asp:ListItem Value="31">甄試人數</asp:ListItem>
                                    <asp:ListItem Value="32">離訓人數</asp:ListItem>
                                </asp:RadioButtonList>
                                <%--<asp:ListItem Value="14">就業人數</asp:ListItem>
                                    <asp:ListItem Value="15">在職者(托育及照服員計畫)</asp:ListItem>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">X軸</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="XRoll" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1" Width="100%">
                                </asp:RadioButtonList>
                                <%--<asp:ListItem Value="1">性別</asp:ListItem>
                                    <asp:ListItem Value="2">年齡</asp:ListItem>
                                    <asp:ListItem Value="3">教育程度</asp:ListItem>
                                    <asp:ListItem Value="4">身分別</asp:ListItem>
                                    <asp:ListItem Value="5">受訓學員(通訊)地理分佈</asp:ListItem>
                                    <asp:ListItem Value="7">參訓單位類別</asp:ListItem>
                                    <asp:ListItem Value="8">開班縣市</asp:ListItem>
                                    <asp:ListItem Value="9">訓練時數</asp:ListItem>
                                    <asp:ListItem Value="21">訓練職類(大類)</asp:ListItem>
                                    <asp:ListItem Value="22">就職狀況</asp:ListItem>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">Y軸</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="YRoll" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1" Width="100%">
                                </asp:RadioButtonList>
                                <%--<asp:ListItem Value="1">性別</asp:ListItem>
                                    <asp:ListItem Value="2">年齡</asp:ListItem>
                                    <asp:ListItem Value="3">教育程度</asp:ListItem>
                                    <asp:ListItem Value="4">身分別</asp:ListItem>
                                    <asp:ListItem Value="5">受訓學員(通訊)地理分佈</asp:ListItem>
                                    <asp:ListItem Value="7">參加單位類別</asp:ListItem>
                                    <asp:ListItem Value="8">開班縣市</asp:ListItem>
                                    <asp:ListItem Value="9">訓練時數</asp:ListItem>
                                    <asp:ListItem Value="21">訓練職類(大類)</asp:ListItem>
                                    <asp:ListItem Value="22">就職狀況</asp:ListItem>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">匯出檔案格式</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%-- <tr>
                            <td class="bluecol_need" width="20%">特定對象</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="rbl_ddlkjdfs" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1" Width="100%">
                                </asp:RadioButtonList>
                            </td>
                        </tr>--%>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="btn301_search" runat="server" Text="查詢" CssClass="asp_Export_M"></asp:Button>
                    </p>
                    <table id="DataGroupTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <div id="Div1" runat="server">
                                    <asp:Table ID="DataTable1" runat="server" CssClass="table_sch" BackColor="WhiteSmoke" CellSpacing="0" CellPadding="2" Width="100%"></asp:Table>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Btn301_print1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                &nbsp;<asp:Button ID="btn301_Export" runat="server" Text="匯出Excel" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
