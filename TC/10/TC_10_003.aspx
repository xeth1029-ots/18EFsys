<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_10_003.aspx.vb" Inherits="WDAIIP.TC_10_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>出席統計總表</title>
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
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var objList = document.getElementById(hidobj);
            if (!objList) { return; }
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (objList.value != getCheckBoxListValue(obj).charAt(0)) {
                objList.value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        objList.value = getCheckBoxListValue(obj).charAt(i);
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
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查委員管理&gt;&gt;出席統計總表</asp:Label>
                </td>
            </tr>
        </table>

        <asp:Panel ID="panelSch" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <tr>
                    <td class="bluecol" width="20%">轄區分署</td>
                    <td class="whitecol">
                        <asp:CheckBoxList ID="cblDISTID_SCH" runat="server" RepeatDirection="Horizontal" RepeatColumns="3"></asp:CheckBoxList>
                        <input id="cblDISTID_SCH_List" type="hidden" value="0" runat="server" />
                        <%--<asp:DropDownList ID="ddlDISTID_SCH" runat="server"></asp:DropDownList>--%></td>
                </tr>
                <tr>
                    <td class="bluecol_need">年度區間</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlMYEARS1_SCH" runat="server"></asp:DropDownList>～<asp:DropDownList ID="ddlMYEARS2_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol">計畫別</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblORGPLANKIND_sch" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                </tr>

                <tr>
                    <td class="bluecol">審查會議類別</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblCATEGORY_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="1">轄區</asp:ListItem>
                            <asp:ListItem Value="2">跨區</asp:ListItem>
                        </asp:RadioButtonList></td>
                </tr>

                <tr>
                    <td class="bluecol">受理階段</td>
                    <td class="whitecol">
                        <asp:CheckBoxList ID="cblACCEPTSTAGE_sch" runat="server" RepeatDirection="Horizontal" RepeatColumns="5">
                            <%--<asp:ListItem Value="ALL">全部</asp:ListItem>--%>
                            <asp:ListItem Value="A1">上半年</asp:ListItem>
                            <asp:ListItem Value="A2">上半年申復</asp:ListItem>
                            <asp:ListItem Value="B1">政策性</asp:ListItem>
                            <asp:ListItem Value="B2">政策性申復</asp:ListItem>
                            <asp:ListItem Value="C1">下半年</asp:ListItem>
                            <asp:ListItem Value="C2">下半年申復</asp:ListItem>
                        </asp:CheckBoxList>
                        <input id="cblACCEPTSTAGE_sch_List" type="hidden" value="0" runat="server" />
                    </td>
                </tr>
                <tr id="trRBListExpType" runat="server">
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <asp:Button ID="BtnPrint1" runat="server" Text="列印"  CssClass="asp_Export_M"/>
                        <asp:Button ID="BtnExport1" runat="server" Text="匯出" CssClass="asp_Export_M" />
                    </td>
                </tr>
            </table>
             <div align="center">
                <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
        </asp:Panel>

        <%--<asp:HiddenField ID="Hid_MTSEQ" runat="server" />
        <asp:HiddenField ID="hid_EXAMINER_TABLE_GUID1" runat="server" />--%>
    </form>
    <%--序號、遴聘類別、姓名、現職服務機構、職稱、推薦分署--%>
</body>
</html>
