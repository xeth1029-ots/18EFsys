<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_012_add.aspx.vb" Inherits="TIMS.TC_01_012_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構評鑑設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript">
        function bt_addrow_check() {
            if (document.form1.yearlist_add.value == '' || document.form1.RIDValue.value == '' || document.form1.ClassChar.value == '') {
                //|| document.form1.HistoryList2.value==''
                alert('請輸入年度、訓練機構、訓練性質評鑑代碼');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <p>
        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">訓練機構評鑑設定</font><font color="#990000">-新增(修改)</font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <table class="font" id="ShowDist" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table class="table_sch" id="Table1" runat="server">
                                    <tr>
                                        <td id="td1" runat="server" class="bluecol_need" width="100">
                                            年度
                                        </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:DropDownList ID="yearlist_add" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td id="td6" runat="server" class="bluecol_need">
                                            訓練機構
                                        </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox><input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                                            <input id="RIDValue" type="hidden" size="1" name="RIDValue" runat="server">
                                            <input id="OrgID" type="hidden" size="1" name="OrgID" runat="server"><br>
                                            <span id="HistoryList2" style="position: absolute; display: none">
                                                <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                                </asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td id="td5" runat="server" class="bluecol_need">
                                            訓練性質評鑑
                                        </td>
                                        <td style="height: 18px" colspan="3" class="whitecol">
                                            &nbsp;
                                            <asp:DropDownList ID="ClassChar" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center">
                                    <asp:Button ID="bt_addrow" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button><asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button></div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </p>
    </form>
</body>
</html>
