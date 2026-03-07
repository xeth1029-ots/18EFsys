<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_023.aspx.vb" Inherits="WDAIIP.SD_05_023" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_05_023</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script src="../../js/common.js"></script>
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function CheckData() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類班別\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }	
			
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">就業輔導費預算表</font></FONT>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <asp:Panel ID="TableSearch" runat="server">
                    <table class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="100" class="bluecol">
                                訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" AutoPostBack="True" Width="410px"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="BtnOrg" type="button" value="..." name="BtnOrg" runat="server"><br>
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol_need">
                                職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox><input onclick="choose_class()" type="button" value="...">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <font color="#ffffff"></font>
                                <br>
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <p align="center">
                                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button></p>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </td>
        </tr>
    </table>
    <table id="TableShowData" runat="server" width="740">
        <tr>
            <td align="right" class="whitecol">
                第<asp:Label ID="labTimes" runat="server" ForeColor="Red"></asp:Label>次撥款就業輔導預算表
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" id="TableShowData2" style="border-collapse: collapse" cellspacing="1" cellpadding="1" border="0">
                    <tr>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;班級中文名稱
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="ClassCName" runat="server"></asp:Label><input id="OCIDValue2" type="hidden" name="OCIDValue2" runat="server">
                        </td>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;期別
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="CyclType" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;計畫人數
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="TNum" runat="server"></asp:Label>
                        </td>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;訓練時數
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="THours" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;簽約訓練經費
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="ContractCost" runat="server" Width="90px" MaxLength="10"></asp:TextBox>
                        </td>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;個人就業<br>
                            &nbsp;&nbsp; 輔導費單價
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="RemedCost" runat="server" Width="90px" MaxLength="10"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;就業人數<br>
                            &nbsp;&nbsp; (檢附就業證明)
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="JobENum" runat="server" Width="60px" MaxLength="5"></asp:TextBox>
                        </td>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;就業人數<br>
                            &nbsp;&nbsp; (個案切結證明)
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="JobCNum" runat="server" Width="60px" MaxLength="5"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">
                            &nbsp;&nbsp;&nbsp;就業輔導<br>
                            &nbsp;&nbsp; 證明人數
                        </td>
                        <td class="whitecol" colspan="3">
                            <asp:TextBox ID="JobANum" runat="server" Width="60px" MaxLength="5"></asp:TextBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="center" class="whitecol">
                <asp:Button ID="btn_Save" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="btn_edit" runat="server" Text="編輯" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="btn_cancel" runat="server" Text="取消" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="btn_print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;<asp:Button ID="btn_back" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button><input id="times2" type="hidden" name="times2" runat="server">
                <%--<uc1:pagecontroler id="PageControler1" runat="server"></uc1:pagecontroler>--%>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
