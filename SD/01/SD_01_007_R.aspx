<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_007_R.aspx.vb" Inherits="WDAIIP.SD_01_007_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>參訓學員投保狀況檢核表</title>
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
    <script type="text/javascript">
        function GETvalue() { document.getElementById('Button12').click(); }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function CheckPrint() {
            var msg = '';
            if (document.form1.OCID1.value == '') msg += '請選擇班級!\n';
            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;參訓學員投保狀況檢核表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Frametable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="table_sch" id="table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <input id="DistValue" type="hidden" name="DistValue" runat="server" />
                                <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">班級名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" onclick="choose_class();" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">通俗職類</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr id="trPrtSortType1" runat="server">
                            <td class="bluecol">列印排序方式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="print_type" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">依【身分證字號】排序列印 </asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">依【學號】排序列印</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trStExChoice1" runat="server">
                            <td class="bluecol_need">勾稽條件</td>
                            <td class="whitecol" colspan="3">
                                <%--<asp:ListItem Value="1" Selected="True">不區分</asp:ListItem>--%>
                                <asp:RadioButtonList ID="rblStExChoice1" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="2">開訓日</asp:ListItem>
                                    <asp:ListItem Value="4">開訓日前4日</asp:ListItem>
                                    <asp:ListItem Value="3">甄試日</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--<tr>
                            <td class="bluecol">作業顯示模式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList Style="z-index: 0" ID="rblWorkMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1" Selected="True">模糊顯示</asp:ListItem>
                                    <asp:ListItem Value="2">正常顯示</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                    <%--<asp:ListItem Value="XLSX">XLSX</asp:ListItem>--%>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <div style="width: 100%" align="center" class="whitecol">
            <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
            <asp:Button ID="btnExport" runat="server" Text="匯出e網民眾投保狀況檢核表" CssClass="asp_Export_M" CommandName="exp1"></asp:Button>
            <asp:Button ID="btnExport2" runat="server" Text="匯出投保狀況檢核表" CssClass="asp_Export_M" CommandName="exp2"></asp:Button>
        </div>
        <asp:HiddenField ID="HidPrinttype" runat="server" />
    </form>
</body>
</html>
