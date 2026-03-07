<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_019.aspx.vb" Inherits="WDAIIP.TC_01_019" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>最近一次TTQS評核結果確認</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;最近一次TTQS評核結果確認</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table3" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td colspan="2" align="center" class="table_title" width="100%">單位最近一次TTQS評核結果等級 </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">年度</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="YEARS_ROC" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">月份</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="MONTHS" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">分署訓練計畫</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="PLANNAME" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">轉入／資料更新時間</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="IMPORTDATE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">機構名稱</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="ORGNAME" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">統一編號</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="COMIDNO" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">機構別</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="ORGKIND_N" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">評核版別</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="SENDVER_N" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">申請目的</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="GOAL" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">評核結果</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="RESULT_N" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">展延</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="EXTLICENS" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">評核日期</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="SENDDATE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">發文日期</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="ISSUEDATE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">有效期限</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="VALIDDATE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">TTQS訓練機構名稱</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="MEMO2" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol_need">單位確認</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="rblCONFIRM" runat="server"></asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">原因說明</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="txREASON1" runat="server" TextMode="MultiLine" Rows="5" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="2" align="center">
                                <asp:Label ID="lab_msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="BTN_EDIT1" runat="server" Text="修改" CssClass="asp_button_S"></asp:Button>&nbsp;
                                <asp:Button ID="BTN_SAVE1" runat="server" Text="確認送出" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Label ID="labmmo1" runat="server" ForeColor="Red">※確認送出後即鎖定不可再修改!</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>

        </table>
        <asp:HiddenField ID="Hid_ORGID" runat="server" />
        <asp:HiddenField ID="Hid_OTTID" runat="server" />
        <asp:HiddenField ID="Hid_YEARS" runat="server" />
        <asp:HiddenField ID="Hid_MONTHS" runat="server" />
    </form>
</body>
</html>
