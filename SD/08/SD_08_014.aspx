<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_014.aspx.vb" Inherits="WDAIIP.SD_08_014" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>津貼申請</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
        <tr>
            <td width="15%" bgcolor="#2aafc0">
                <font color="#ffffff">所屬單位</font>
            </td>
            <td colspan="3" bgcolor="#ebf8ff">
                <asp:TextBox ID="EOrgName" runat="server" onfocus="this.blur()" Columns="60"></asp:TextBox>
                <input id="EOrgId" type="hidden" runat="server" size="1">
            </td>
        </tr>
        <tr>
            <td width="15%" bgcolor="#2aafc0">
                <font color="#ffffff">審核狀態</font>
            </td>
            <td colspan="3" bgcolor="#ebf8ff">
                <table class="font" cellspacing="0" bordercolordark="#ffffff" cellpadding="0" width="100%" align="center" bordercolorlight="#666666" border="0">
                    <tr>
                        <td nowrap align="center" width="7%">
                            初審：
                        </td>
                        <td width="20%">
                            &nbsp;<asp:Label ID="lblAppliedStatusF" runat="server"></asp:Label>
                        </td>
                        <td nowrap align="center" width="10%">
                            送勞保局：
                        </td>
                        <td width="20%">
                            &nbsp;<asp:Label ID="lblisDownload" runat="server"></asp:Label>
                        </td>
                        <td nowrap align="center" width="7%">
                            勾稽：
                        </td>
                        <td>
                            &nbsp;<asp:Label ID="lblAppliedStatusFin" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">學員姓名</font>
            </td>
            <td width="35%" bgcolor="#ebf8ff">
                <asp:TextBox ID="EName" runat="server" onfocus="this.blur()" MaxLength="30"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0" width="15%">
                <font color="#ffffff">身分證號</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EIdno" runat="server" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">生日</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EBirthday" runat="server" onfocus="this.blur()" Columns="12" MaxLength="10"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">性別</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:RadioButtonList ID="ESex" runat="server" CssClass="font" Height="10px" RepeatColumns="2">
                    <asp:ListItem Value="M" Selected="True">男</asp:ListItem>
                    <asp:ListItem Value="F">女</asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">通訊地址</font>
            </td>
            <td colspan="3" bgcolor="#ebf8ff">
                <asp:TextBox ID="City1" runat="server" onfocus="this.blur()" Width="130px"></asp:TextBox><asp:TextBox ID="Address" runat="server" onfocus="this.blur()" Width="352px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">戶籍地址</font>
            </td>
            <td colspan="3" bgcolor="#ebf8ff">
                <asp:TextBox ID="City2" runat="server" onfocus="this.blur()" Width="130px"></asp:TextBox><asp:TextBox ID="HouseholdAddress" runat="server" onfocus="this.blur()" Width="352px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">聯絡電話</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EPhone" runat="server" onfocus="this.blur()" Columns="20" MaxLength="25"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">申請類別</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:DropDownList ID="EIdentityID" runat="server" Enabled="False">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">申請障別</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:DropDownList ID="EHandicat" runat="server" Enabled="False">
                </asp:DropDownList>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">身心障礙等級</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:DropDownList ID="EHandicatlevel" runat="server" Enabled="False">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">實際受訓起日</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="ETSDate" runat="server" MaxLength="10" Columns="12" onfocus="this.blur()"></asp:TextBox>
                <td bgcolor="#2aafc0">
                    <font color="#ffffff">課程起迄日</font>
                </td>
                <td>
                    <asp:TextBox ID="CSDate" runat="server" MaxLength="10" Columns="12" onfocus="this.blur()"></asp:TextBox>～
                    <asp:TextBox ID="ETEDate" runat="server" MaxLength="10" Columns="12" onfocus="this.blur()"></asp:TextBox>
                </td>
            </td>
        </tr>
        </TR>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">申請日期</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EApplydate" runat="server" onfocus="this.blur()" Columns="12" MaxLength="10"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">參與職類</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:DropDownList ID="ETrainType" runat="server" Enabled="False">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">參訓班別</font>
            </td>
            <td colspan="3" bgcolor="#ebf8ff">
                <asp:TextBox ID="EClassName" runat="server" onfocus="this.blur()" MaxLength="50" Width="416px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">補助月數</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="ETMonth" runat="server" onfocus="this.blur()" Columns="12" MaxLength="5"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">補助金額</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="ETMoney" runat="server" onfocus="this.blur()" Columns="12" MaxLength="6"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">申請月數</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EAMonth" runat="server" onfocus="this.blur()" Columns="12" MaxLength="5"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">申請金額</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EAMoney" runat="server" onfocus="this.blur()" Columns="12" MaxLength="6"></asp:TextBox>
                <input id="payValue" type="hidden" name="payValue" runat="server" size="1">
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">核發實領月數</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EPayMonth" runat="server" onfocus="this.blur()" Columns="12" MaxLength="5"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">核發實領金額</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="EPayMoney" runat="server" onfocus="this.blur()" Columns="12" MaxLength="6"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">離退訓日期</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="ELdate" runat="server" onfocus="this.blur()" Columns="12" MaxLength="10"></asp:TextBox><input id="ELflag1" type="radio" value="ELflag1" name="ELflag" runat="server" disabled>離訓<input id="ELflag2" type="radio" value="ELflag2" name="ELflag" runat="server" disabled checked>退訓
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">離退原因</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="RTReason" runat="server" onfocus="this.blur()" Columns="50"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">繳回月數</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="RtnMonth" runat="server" onfocus="this.blur()" Columns="12" MaxLength="5"></asp:TextBox>
            </td>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">繳回金額</font>
            </td>
            <td bgcolor="#ebf8ff">
                <asp:TextBox ID="ERtnMoney" runat="server" onfocus="this.blur()" Columns="12" MaxLength="6"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td bgcolor="#2aafc0">
                <font color="#ffffff">離退審核結果</font>
            </td>
            <td colspan="3" bgcolor="#ebf8ff">
                <asp:TextBox ID="LVerify" runat="server" onfocus="this.blur()" Columns="12" MaxLength="10"></asp:TextBox>
            </td>
        </tr>
    </table>
    <br style="line-height: 5px">
    <table class="font" cellspacing="0" bordercolordark="#ffffff" cellpadding="0" width="100%" align="center" bordercolorlight="#666666" border="0">
        <tr align="center">
            <td>
                <asp:Button ID="btnExit" runat="server" CssClass="btn_tab" CausesValidation="False" Text="關閉視窗"></asp:Button>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
