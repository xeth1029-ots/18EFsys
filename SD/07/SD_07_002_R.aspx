<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_07_002_R.aspx.vb" Inherits="WDAIIP.SD_07_002_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印技能檢定名冊</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button13').click();
        }
        function SetOneOCID() {
            document.getElementById('Button8').click();
        }
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID + '&BtnName=btnSchExamTime');
        }

        function ReportPrint() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') {
                msg += '請選擇班級職類\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

    </script>
    <%-- <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 34px; }
        .auto-style2 { height: 34px; }
    </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;技能檢定管理&gt;&gt;列印技能檢定名冊</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;技能檢定管理&gt;&gt;列印技能檢定名冊</td>
					</tr>
				</table>--%>
                    <div align="center">
                        <table id="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0" align="center" class="table_nw">
                            <tr>
                                <td class="bluecol" style="width: 20%">訓練機構</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                    <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini">
                                    <asp:Button ID="btnSetOneOCID" Style="display: none" runat="server" Text="btnSetOneOCID"></asp:Button>
                                    <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級/職類</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input type="button" value="..." onclick="choose_class()" class="asp_button_Mini">
                                    <input id="TMIDValue1" type="hidden" name="Hidden2" runat="server" size="1">
                                    <input id="OCIDValue1" type="hidden" name="Hidden1" runat="server" size="1">
                                    <span id="HistoryList" style="position: absolute; left: 270px; display: none">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%">
                                        </asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢定職類/名稱</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlKindTime" Style="z-index: 0" runat="server">
                                    </asp:DropDownList>
                                    <asp:Button ID="btnSchExamTime" Style="display: none" runat="server" Text="btnSchExamTime"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
