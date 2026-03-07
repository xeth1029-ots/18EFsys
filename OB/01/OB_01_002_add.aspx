<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_002_add.aspx.vb" Inherits="WDAIIP.OB_01_002_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>OB_01_002_add</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>    
    <script language="javascript" src="../../js/date-picker.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/common.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/openwin/openwin.js" type="text/javascript"></script>
    <script language="JavaScript" type="text/javascript">

        function set_Orgname1(obj) {
            //debugger;
            var center = document.getElementById('center');
            var Org = document.getElementById('Org');
            var RIDValue = document.getElementById('RIDValue');
            var orgid_value = document.getElementById('orgid_value');
            center.value = '';
            RIDValue.value = '';
            orgid_value.value = '';

            $("#center").unbind("focus");
            center.style.display = 'inline';
            Org.style.display = 'inline';
            if (obj == 'rb1') {
                //只能選擇 center.readOnly = true;
                center.style.display = 'inline';
                Org.style.display = 'inline';
            }

            if (obj == 'rb2') {
                //要可輸入 center.readOnly = false;
                $("#center").bind("focus", function () { $("#center").blur(); });
                center.style.display = 'inline';
                Org.style.display = 'none';
            }
        }
        function rb1_checked() {
            //debugger;
            var rb1 = document.getElementById('rb1');
            rb1.checked = true;

        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">工作小組成員資料建檔</font></FONT>
                            </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" id="TableLay2" cellspacing="1" cellpadding="1">
                    <tbody>
                        <tr>
                            <td class="bluecol" width="100">
                                服務單位型態
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButton ID="rb1" runat="server" GroupName="Type1" Text="TIMS訓練機構"></asp:RadioButton>
                                <asp:RadioButton ID="rb2" runat="server" GroupName="Type1" Text="其他服務單位"></asp:RadioButton>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                服務單位名稱
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" class="asp_button_Mini">
                                <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server"><input id="orgid_value" type="hidden" name="orgid_value" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                服務部門
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="DeptName" runat="server" MaxLength="20" Width="300px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                成員姓名
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="memName" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                具備採購法證照
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rblQualified" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="N">無</asp:ListItem>
                                    <asp:ListItem Value="B">基礎</asp:ListItem>
                                    <asp:ListItem Value="A">進階</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <p align="center">
                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button><font face="新細明體">&nbsp;</font>
                    <asp:Button ID="btnReturn" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button></p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
