<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_009.aspx.vb" Inherits="WDAIIP.CM_03_009" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職業訓練生活津貼主要特定對象統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">

        //檢查列印條件為
        var itm = -1;
        function CheckPrint() {
            var msg = "";
            if (document.getElementById("txt_STDateS").value != "" && document.getElementById("txt_STDateE").value != "") {
                if (document.getElementById("txt_STDateS").value > document.getElementById("txt_STDateE").value) {
                    msg += "[開訓區間的迄日]必需大於[開訓區間的起日]\n";
                }
            }
            if (document.getElementById("txt_FTDateS").value != "" && document.getElementById("txt_FTDateE").value != "") {
                if (document.getElementById("txt_FTDateS").value > document.getElementById("txt_FTDateE").value) {
                    msg += "[結訓區間的迄日]必需大於[結訓區間的起日]\n";
                }
            }
            if (document.getElementById("txt_STDateS").value != "") {
                if (document.getElementById("txt_FTDateE").value != "") {
                    if (document.getElementById("txt_FTDateE").value < document.getElementById("txt_STDateS").value) {
                        msg += "[結訓區間的迄日]必需大於[開訓區間的起日]\n";
                    }
                }
            }
            if (document.getElementById("txt_STDateS").value == "" && document.getElementById("txt_STDateE").value == "" && document.getElementById("txt_FTDateS").value == "" && document.getElementById("txt_FTDateE").value == "" && document.getElementById("list_Year").value == "") {
                msg += '[年度]、[開訓區間]、[結訓區間],請擇一輸入查詢\n';
            }
            if (!CheckChecked(document.getElementById("chk_District"), "checkbox")) { msg += "請選擇轄區\n"; }
            if (!CheckChecked(document.getElementById("chk_TPlanID"), "checkbox")) { msg += "請選擇計畫\n"; }
            itm = -1;
            if (!CheckChecked(document.getElementById("rdo_Mode"), "radio")) { msg += "請選擇統計項目\n"; }
            if (itm > 0) {
                chk = false;
                obj = document.getElementById("chk_Identity");
                if (obj) {
                    for (i = 1; i < obj.childNodes.length; i++) {
                        if (obj.childNodes.item(i).type == "checkbox") {
                            if (obj.childNodes.item(i).checked) {
                                chk = true;
                            }
                        }
                    }
                }
                if (!chk) { msg += "請選擇身分別\n"; }
            }
            if (msg != "") {
                alert(msg);
                return false;
            } else {
                return true;
            }
        }

        //檢查是否選取
        function CheckChecked(obj, itmType) {
            var rst = false;

            if (obj) {
                for (i = 0; i < obj.childNodes.length; i++) {
                    if (obj.childNodes.item(i).type == itmType) {
                        if (obj.childNodes.item(i).checked) {
                            itm = i;
                            rst = true;
                        }
                    }
                }
            }
            return rst;
        }

        //檢查統計項目若是身分別就把身分別隱藏					
        function ChangeMode(obj) {
            var cnt = 0;
            var IdentityTR = document.getElementById('IdentityTR');
            if (obj) {
                for (i = 0; i < obj.childNodes.length; i++) {
                    if (obj.childNodes.item(i).type == "radio") {
                        if (obj.childNodes.item(i).checked) {
                            if (cnt == 0) {
                                IdentityTR.style.display = "none";
                            } else {
                                IdentityTR.style.display = "inline";
                            }
                        }
                        cnt++;
                    }
                }
            }
        }

        //選擇全部
        function SelectAll(obj) {
            if (obj) {
                for (i = 1; i < obj.childNodes.length; i++) {
                    if (obj.childNodes.item(i).type == "checkbox") {
                        obj.childNodes.item(i).checked = obj.childNodes.item(0).checked;
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁>>訓練與需求管理>>統計分析>>職業訓練生活津貼主要特定對象統計表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2" runat="server" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="100">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="list_Year" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓區間
                            </td>
                            <td class="whitecol" style="height: 28px">
                                <asp:TextBox ID="txt_STDateS" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txt_STDateS','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="txt_STDateE" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txt_STDateE','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓區間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txt_FTDateS" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txt_FTDateS','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="txt_FTDateE" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txt_FTDateE','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">轄區
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chk_District" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="4">
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">計畫
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chk_TPlanID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="2">
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="height: 20px">統計項目
                            </td>
                            <td class="whitecol" style="height: 20px">
                                <asp:RadioButtonList ID="rdo_Mode" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="0">身分別</asp:ListItem>
                                    <asp:ListItem Value="1">年齡</asp:ListItem>
                                    <asp:ListItem Value="2">訓練職類</asp:ListItem>
                                    <asp:ListItem Value="3">教育程度</asp:ListItem>
                                    <asp:ListItem Value="4">性別</asp:ListItem>
                                    <asp:ListItem Value="5">通俗職類</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="IdentityTR" runat="server">
                            <td class="bluecol_need">身分別
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chk_Identity" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="6">
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
