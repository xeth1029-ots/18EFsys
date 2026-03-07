<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_004_email_add.aspx.vb" Inherits="WDAIIP.SD_01_004_email_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>e 網審核郵件設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function checkData() {
            var msg = "";
            var tmpcontent = document.getElementById("eComment").value.replace(/\r\n/g, "");
            //alert('length=['+document.getelementbyid("eComment").value.length+']\n length=['+tmpcontent.length+']');
            if (document.getElementById("eComment").value.length > 500) {
                msg += '「說明事項」欄位 ' + document.getElementById("eComment").value.length + ' 個字元超過欄位 500 個字元最大限制!\n';
            }
            if (msg != "") {
                alert(msg);
                return false;
            } else {
                return true;
            }
        }
    </script>
    <%--<style type="text/css">
        .auto-style1 { height: 23px; }
    </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;e 網審核郵件設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table_nw" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tbody>
                <tr>
                    <td align="center">
                        <table class="table_sch" id="table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td id="td_dist" colspan="4" class="whitecol">轄區：<asp:Label ID="DistID" runat="server" CssClass="font"></asp:Label></td>
                            </tr>
                            <tr id="TR_PlanYear" runat="server">
                                <td id="Td_PlanYear" colspan="4" runat="server" class="whitecol">年度：<asp:Label ID="PlanYear" runat="server" CssClass="font"></asp:Label></td>
                            </tr>
                            <tr id="TR_CtrlOrg" runat="server">
                                <td id="TD_CtrlOrg" colspan="4" runat="server" class="whitecol">管控單位：<asp:Label ID="CtrlOrg" runat="server" CssClass="font"></asp:Label></td>
                            </tr>
                            <tr>
                                <td id="TD_org" colspan="4" runat="server" class="whitecol">訓練機構：<asp:Label ID="OrgName" runat="server" CssClass="font"></asp:Label>
                                    <input id="OrgID" style="width: 10%" type="hidden" size="4" name="OrgID" runat="server">&nbsp;
                                    <input id="RIDValue" style="width: 10%" type="hidden" size="4" runat="server" name="RIDValue">&nbsp;
                                    <input id="OCID" style="width: 10%" type="hidden" size="4" runat="server" name="OCID">
                                </td>
                            </tr>
                            <tr id="TR_Class" runat="server">
                                <td id="TD_Class" class="whitecol" colspan="4" runat="server">班級名稱：<asp:Label ID="ClassName" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol" align="center" width="20%">e 網報名成功說明事項：</td>
                                <td class="whitecol" colspan="3"><textarea id="eComment" runat="server" style="width: 60%" rows="8" cols="10"></textarea></td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4" class="whitecol">
                                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnBack" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table>
    </form>
</body>
</html>