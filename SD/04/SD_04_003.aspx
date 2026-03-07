<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_003.aspx.vb" Inherits="WDAIIP.SD_04_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>排課列表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function ChangeMode() {
            var OCID1 = document.getElementById('OCID1');
            var OCID2 = document.getElementById('OCID2');
            OCID1.style.display = 'none';
            OCID2.style.display = 'none';
            if (getRadioValue(document.form1.ShowMode) == 1) {
                OCID1.style.display = ''; //'inline';
            }
            else {
                OCID2.style.display = ''; //'inline';
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;排課列表</asp:Label>
                </td>
            </tr>
        </table>
        <div>
            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <table class="table_nw" id="Table4" width="100%" cellpadding="1" cellspacing="1">
                                        <tr>
                                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                                            <td colspan="3" class="whitecol">
                                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()" AutoPostBack="True"></asp:TextBox>
                                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="button_b_Mini">
                                                <asp:Button ID="Button2" runat="server" Text="更新班級代碼" CssClass="asp_button_M"></asp:Button>
                                                <input id="SingleValue" type="hidden" name="SingleValue" runat="server">
                                                <span id="HistoryList2" style="display: none">
                                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                                </span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" style="width: 20%">排課種類</td>
                                            <td class="whitecol" style="width: 30%">
                                                <asp:RadioButtonList ID="ShowMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" AutoPostBack="True">
                                                    <asp:ListItem Value="1" Selected="True">正式</asp:ListItem>
                                                    <asp:ListItem Value="2">預排</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                            <td class="bluecol" style="width: 20%">班別名稱 </td>
                                            <td class="whitecol" style="width: 30%">
                                                <asp:DropDownList ID="OCID1" AutoPostBack="True" runat="server"></asp:DropDownList>
                                                <asp:DropDownList ID="OCID2" runat="server" AutoPostBack="True"></asp:DropDownList>
                                                <asp:Button ID="Button1" runat="server" Text="回全期排課" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" colspan="4" align="center">
                                                <asp:Button ID="btnSchAct3" runat="server" Text="重新查詢" />
                                                &nbsp;</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Btn_TR1" runat="server">
                                <td align="center" class="whitecol">
                                    <asp:Button ID="Button3B" runat="server" Text="審核確認" CommandName="ResultY" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button4B" runat="server" Text="取消審核" CommandName="ResultN" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="List_Class" Width="100%" CssClass="font" runat="server" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <ItemStyle HorizontalAlign="Center" />
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderStyle-Width="8%">
                                                <HeaderTemplate>日期</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="LSchoolDate" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="8%">
                                                <HeaderTemplate>星期</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="LWeekday" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第一節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les1" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea1" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea13" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom1" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第二節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les2" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea2" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea14" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom2" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第三節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les3" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea3" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea15" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom3" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第四節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les4" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea4" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea16" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom4" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第五節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les5" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea5" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea17" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom5" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第六節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les6" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea6" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea18" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom6" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第七節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les7" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea7" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea19" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom7" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第八節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les8" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea8" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea20" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom8" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第九節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les9" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea9" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea21" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom9" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第十節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les10" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea10" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea22" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom10" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第十一節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les11" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea11" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea23" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom11" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="7%">
                                                <HeaderTemplate>第十二節</HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Label ID="Les12" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea12" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Tea24" runat="server"></asp:Label><br>
                                                    <asp:Label ID="Rom12" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn Visible="False" DataField="Vacation" HeaderText="Vacation"></asp:BoundColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Label ID="Show_NoData" runat="server" Visible="False"></asp:Label></td>
                            </tr>
                            <tr id="Btn_TR2" runat="server">
                                <td align="center" class="whitecol">
                                    <asp:Button ID="Button3" runat="server" Text="審核確認" CommandName="ResultY" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button4" runat="server" Text="取消審核" CommandName="ResultN" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
