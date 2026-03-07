<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_003_chk.aspx.vb"
    Inherits="WDAIIP.SYS_01_003_chk" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號審核確認</title>
    <meta content="False" name="vs_snapToGrid">
    <meta content="False" name="vs_showGrid">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">			
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        &nbsp;&nbsp;&nbsp;
    <table class="font" width="100%">
        <tr>
            <td class="font">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號審核確認
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                    <tr>
                        <td class="bluecol" style="width: 20%">訓練機構</td>
                        <td class="whitecol" style="width: 80%">
                            <asp:TextBox ID="TBplan" runat="server" Width="349px" onfocus="this.blur()"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">帳號</td>
                        <td class="whitecol">
                            <asp:TextBox ID="nameid" runat="server" onfocus="this.blur()" MaxLength="15"></asp:TextBox>
                            <%--<input id="userpass" type="hidden" name="userpass" runat="server">--%>
                            <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server" />
                            <input id="OrgID" type="hidden" name="OrgID" runat="server" />
                            <%--
					            <INPUT onclick="GetID();" type="button" value="檢查帳號">
                            --%>
                            <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                            <input id="LIDValue" type="hidden" name="LIDValue" runat="server" />
                        </td>
                    </tr>
                    <%--<tr><td class="TD_TD1" style="WIDTH: 103px; HEIGHT: 26px" align="center"><div align="left">&nbsp;&nbsp;&nbsp;&nbsp;密碼</div></td><td style="HEIGHT: 26px"><asp:textbox id="userpass" Runat="server"></asp:textbox></td></tr><tr><td class="TD_TD1" style="WIDTH: 103px" align="center"><div align="left">&nbsp;&nbsp;&nbsp;&nbsp;確認密碼</div></td><td><asp:textbox id="userpass2" Runat="server"></asp:textbox></td></tr>--%>
                    <tr>
                        <td class="bluecol_need">角色
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="Role" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">姓名
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="name" runat="server" onfocus="this.blur()"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">身分證號碼
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="IDNO" runat="server" onfocus="this.blur()"></asp:TextBox>
                            <%--
					            <asp:button id="Button1" runat="server" Text="檢查身分證" CausesValidation="False"></asp:button>
                            --%>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">電話
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="telphone" runat="server" onfocus="this.blur()"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">E_MAIL
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="email" runat="server" onfocus="this.blur()"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">申請種類
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="ApplyType" runat="server" onfocus="this.blur()" BorderStyle="None"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">審核同意
                        </td>
                        <td class="whitecol">
                            <asp:CheckBox ID="Result" runat="server" Checked="True"></asp:CheckBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">備註
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="Note" runat="server" Width="347px" MaxLength="100"></asp:TextBox>
                        </td>
                    </tr>
                </table>
                <table width="100%">
                    <tr>
                        <td align="center" class="whitecol">
                            <asp:ValidationSummary ID="totalmsg" runat="server" ShowMessageBox="True" ShowSummary="False"
                                DisplayMode="List"></asp:ValidationSummary>
                            &nbsp;
                            <asp:Button ID="btu_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button><font
                                face="新細明體">&nbsp;</font>
                            <input id="back" type="button" value="回上一頁" name="back" runat="server" class="button_b_S">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

        &nbsp;&nbsp;
    </form>
</body>
</html>
