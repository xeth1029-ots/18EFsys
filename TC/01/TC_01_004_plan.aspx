<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_plan.aspx.vb" Inherits="WDAIIP.TC_01_004_plan" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料轉入</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function but_edit(planid, ComIDNO, SeqNO, pageid, RID) {
            location.href = 'TC_01_004_InsertPlan.aspx?planid=' + planid + '&ComIDNO=' + ComIDNO + '&SeqNO=' + SeqNO + '&ProcessType=Update&ID=' + pageid + '&RID=' + RID;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server"> 首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定&gt;&gt;開班資料轉入</asp:Label>
                </td>
            </tr>
        </table>
        <%--<input id="check_add" style="width: 56px; height: 22px" type="hidden" size="4" name="check_add" runat="server">--%>
        <table class="font" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="20%" class="bluecol">訓練單位</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                                <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol"><asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server" class="button_b_Mini">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol"><asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">期別</td>
                            <td colspan="3" class="whitecol"><asp:TextBox ID="TB_cycltype" runat="server" Columns="5" MaxLength="5" Width="15%"></asp:TextBox></td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="left" class="whitecol">
                                <div align="center" class="whitecol">
                                    <asp:Button ID="save" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <!--<INPUT id="Button1" type="button" value="回上一頁" name="Button1" runat="server">-->
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br>
        <%----%>
        <asp:Panel ID="Panel" runat="server" Visible="False"><table id="search_tbl" class="font" border="0" cellspacing="1" cellpadding="8" width="100%" runat="server"></table></asp:Panel>
    </form>
</body>
</html>