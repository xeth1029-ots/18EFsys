<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_008_R.aspx.vb" Inherits="WDAIIP.SD_09_008_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>操行成績明細表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function choose_class() { openClass('../02/SD_02_ch.aspx'); }

        function ReportPrint() {
            if (document.form1.syears.selectedIndex == 0) {
                alert('請選擇年度');
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="0" cellpadding="0" width="740" border="0">
            <tr>
                <td class="font">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;<font color="#990000">列印</font><font color="#990000">操行成績明細表</font>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol_need">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="syears" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="76">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                                <input id="OCIDValue1" style="width: 30px; height: 22px" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" style="width: 33px; height: 22px" type="hidden" name="Hidden1" runat="server" />
                                <asp:Button ID="Button2" Style="display: none" runat="server" Text="Button2" CssClass="asp_button_S"></asp:Button>
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="font">
                    <font color="red">請先完成 學員動態管理&gt;&gt;教務管理&gt;&gt; 結訓成績登錄 -&gt; 選擇班級 -&gt; 成績總類：操行&nbsp;<br>
                        計算儲存後，再列印此報表。</font>
                </td>
            </tr>
        </table>
        <br />
        <div style="width: 600" align="center">
            <asp:Button ID="print_submit" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
        </div>
    </form>
</body>
</html>
