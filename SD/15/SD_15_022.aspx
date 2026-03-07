<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_022.aspx.vb" Inherits="WDAIIP.SD_15_022" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>提案彙總表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function CHK_RBL_CROSSDIST_SCH() {
            setTimeout(function () {
                //D:不區分/C:跨區提案單位/J:轄區提案單位
                var radioValue = $("input[type='radio'][name='RBL_CrossDist_SCH']:checked").val();
                if (!radioValue) { return; }
                (radioValue == "C") ? $("#center").hide() : $("#center").show();
                (radioValue == "C") ? $("#Button2").hide() : $("#Button2").show();
                (radioValue == "C") ? $("#lab_center_msg2").show() : $("#lab_center_msg2").hide();
                //(radioValue && radioValue == "C") ? $("#HistoryList2").hide() : $("#HistoryList2").show();
                //if (radioValue) { alert("Your are a - " + radioValue); }
            }, 500);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;提案彙總表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />&nbsp;
							    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini" />
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                                <asp:Label ID="lab_center_msg2" runat="server" Text="(選擇 跨區提案單位，排除【訓練機構】條件)" Style="color: #808080; display: none"></asp:Label>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rblOrgKind2" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ～
							<asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol_need">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;研提資料</td>
                            <td colspan="3" class="whitecol"><%--檢送資料-未檢送-含未檢送研提資料--%>
                                <%--<asp:CheckBox ID="CB_DataNotSent_SCH" runat="server" Text="含未檢送研提資料" />--%>
                                <asp:RadioButtonList ID="RBL_DataNotSent_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y" Selected="True">有檢送</asp:ListItem>
                                    <asp:ListItem Value="O">含未檢送研提資料</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_CrossDist_TP28" runat="server">
                            <td class="bluecol">跨區/轄區提案</td>
                            <td class="whitecol" colspan="3">
                                <%--<asp:RadioButton ID="RBL_CrossDist_SCH_D" runat="server" Text="不區分" GroupName="RBL_CrossDist_SCH" />
                                <asp:RadioButton ID="RBL_CrossDist_SCH_C" runat="server" Text="跨區提案單位" Checked="True" GroupName="RBL_CrossDist_SCH" />
                                <asp:RadioButton ID="RBL_CrossDist_SCH_J" runat="server" Text="轄區提案單位" GroupName="RBL_CrossDist_SCH" />--%>
                                <%-- //D:不區分/C:跨區提案單位/J:轄區提案單位--%>
                                <asp:RadioButtonList ID="RBL_CrossDist_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal" GroupName="RBL_CrossDist_SCH">
                                    <asp:ListItem Value="D" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="C">跨區提案單位</asp:ListItem>
                                    <asp:ListItem Value="J">轄區提案單位</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程分類</td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="cblDepot12" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="HidcblDepot12" type="hidden" value="0" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="BtnExp1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
