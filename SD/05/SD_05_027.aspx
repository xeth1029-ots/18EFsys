<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_027.aspx.vb" Inherits="WDAIIP.SD_05_027" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>操行分數設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        //檢核儲存內容格式
        function chkSave() {
            var Mytable = document.getElementById("DataGrid2");
            var msg = '';

            for (var i = 0; i < Mytable.rows.length; i++) {
                var labName = Mytable.rows(i).cells(0).children(1);
                var iScore = Mytable.rows(i).cells(1).children(0);

                if (isBlank(iScore)) {
                    msg += "請輸入" + labName.innerText + " 扣分分數\r\n";
                } else if (!isFloat2(iScore.value)) {
                    msg += labName.innerText + " 扣分分數格式錯誤\r\n";
                } else if (iScore.value < 0 || iScore.value > 100) {
                    msg += labName.innerText + " 扣分分數請輸入介於 0 ~ 100的數字 \r\n";
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" border="0">
                        <tr>
                            <td class="font">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">操行分數設定</font>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="tbSch" runat="server">
                        <table class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="bluecol" width="100">訓練職類
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="160px"></asp:TextBox>
                                    <input id="btu_sel" onclick="openTrain(document.getElementById('trainvalue').value);" type="button" value="..." name="btu_sel" runat="server" class="button_b_Mini">&nbsp;
                                <input id="trainvalue" type="hidden" name="trainvalue" runat="server">
                                    <input id="jobvalue" type="hidden" name="jobvalue" runat="server">
                                </td>
                                <td class="bluecol" width="100">&nbsp;通俗職類
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtcjob_name" runat="server" onfocus="this.blur()" Columns="20"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練機構
                                </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                    <input id="org" type="button" value="..." name="org" runat="server" class="button_b_Mini"><br>
                                    <span id="HistoryList2" style="position: absolute; display: none">
                                        <asp:Table ID="historyrid" runat="server" Width="310px">
                                        </asp:Table>
                                    </span>
                                    <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
                                    <input type="hidden" id="hidOrgID" runat="server">
                                    <asp:Button ID="btnOrgSet" runat="server" Text="單位操行分數設定" CssClass="asp_button_L"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtSchClass" runat="server"></asp:TextBox>
                                </td>
                                <td class="bluecol">期別
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtCyclType" runat="server" Columns="5"></asp:TextBox>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center">
                                    <br style="line-height: 5px">
                                    <asp:Button ID="btnSch" Text="查詢" runat="server" CssClass="asp_button_S"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <br style="line-height: 5px">
                                    <asp:Label ID="labMsg" Style="color: red" runat="server" Visible="False">查無資料!!</asp:Label>
                                    <asp:DataGrid ID="DataGrid1" runat="server" PagerStyle-Visible="False" AutoGenerateColumns="False" AllowPaging="true" AllowSorting="true" CssClass="font" Width="100%">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號" ItemStyle-Width="5%" ItemStyle-HorizontalAlign="Center"></asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="年度" ItemStyle-Width="10%" ItemStyle-HorizontalAlign="Center">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDYear" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="訓練機構" ItemStyle-Width="33%">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDOrgName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="班級名稱" ItemStyle-Width="30%">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDClassName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="分數設定" ItemStyle-Width="10%" ItemStyle-HorizontalAlign="Center">
                                                <ItemTemplate>
                                                    <asp:Label ID="labStatus" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                                <ItemTemplate>
                                                    <asp:Button ID="btnEdit" Text="設定" CommandName="edt" runat="server"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <input id="hidOCID" type="hidden" runat="server">
                    <asp:Panel ID="tbEdit" runat="server">
                        <table class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="bluecol" width="100">年度
                                </td>
                                <td class="whitecol">
                                    <asp:Label ID="labYear" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練機構
                                </td>
                                <td class="whitecol">
                                    <asp:Label ID="labOrg" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr id="trClass" runat="server">
                                <td class="bluecol">班級名稱
                                </td>
                                <td class="whitecol">
                                    <asp:Label ID="labClass" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">假別
                                </td>
                                <td class="whitecol">
                                    <asp:DataGrid ID="Datagrid2" runat="server" PagerStyle-Visible="False" AutoGenerateColumns="False" ShowHeader="False" CssClass="font" Width="100%">
                                        <HeaderStyle HorizontalAlign="Center" BackColor="#FFFFFF" ForeColor="#FFFFFF"></HeaderStyle>
                                        <ItemStyle BackColor="#FFFFFF"></ItemStyle>
                                        <AlternatingItemStyle BackColor="white"></AlternatingItemStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="" ItemStyle-Width="15%" ItemStyle-HorizontalAlign="right">
                                                <ItemTemplate>
                                                    <input type="hidden" id="hidLeaveID" runat="server">
                                                    <asp:Label ID="labName" runat="server"></asp:Label>扣
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtItem" Style="width: 60px; text-align: right" MaxLength="6" runat="server"></asp:TextBox>分
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" class="whitecol" colspan="2">
                                    <font style="color: red">*假別扣分設定最多輸入至小數點第二位</font>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="btnBack" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
