<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_018.aspx.vb" Inherits="WDAIIP.SD_14_018" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練計畫專職/工作人員名冊</title>
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
    </script>
    <script type="text/javascript">
        function LoalbtnSearch() {
            document.getElementById('btnSearch').click();
        }

        function GETvalue() {
            document.getElementById('Button3').click();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;訓練計畫專職/工作人員名冊</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" class="table_sch">
                        <tr>
                            <td class="bluecol_need" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button2" type="button" value="..." runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" runat="server">
                                <input id="orgid_value" type="hidden" name="orgid_value" runat="server">
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">年度 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_year" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">申請階段 </td><%--年度區間--%>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rblFSQ1_S" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <%--<asp:ListItem Value="00" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="01">上半年</asp:ListItem>
                                    <asp:ListItem Value="02">下半年</asp:ListItem>
                                    <asp:ListItem Value="03">政策性產業</asp:ListItem>--%>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">排序方式 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rblSORT_TYPE1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">依時間</asp:ListItem>
                                    <asp:ListItem Value="2">依序號</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M" /></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Label ID="lMsg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td colspan="4">&nbsp;
							<input id="hidsOMID" type="hidden" name="hidsOMID" runat="server">
                                <table class="table_sch" id="DataGrid1Table" runat="server">
                                    <tr>
                                        <td class="bluecol_sub_need" width="5%">排序 </td>
                                        <td class="bluecol_sub_need" width="20%">職稱 </td>
                                        <td class="bluecol_sub_need" width="20%">姓名 </td>
                                        <td class="bluecol_sub_need" width="20%">聯絡電話 </td>
                                        <td class="bluecol_sub_need" width="15%">申請階段 </td><%--年度別--%>
                                        <td class="whitecol" width="10%"></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tROWNUM1" runat="server" MaxLength="5" Width="90%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tTitle" runat="server" MaxLength="30" Width="90%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tName" runat="server" MaxLength="15" Width="90%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tPhone" runat="server" MaxLength="20" Width="90%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlFSQ1_A" runat="server" Width="80%">
                                                <%--<asp:ListItem Value="01">上半年</asp:ListItem>
                                                <asp:ListItem Value="02">下半年</asp:ListItem>
                                                <asp:ListItem Value="03">政策性產業</asp:ListItem>--%>
                                            </asp:DropDownList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="btnUPD" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="6" class="table_title">訓練計畫專職/工作人員 </td>
                                    </tr>
                                    <tr>
                                        <td colspan="6" align="center">
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <ItemStyle />
                                                <Columns>
                                                    <asp:BoundColumn DataField="ROWNUM1" HeaderText="排序" ItemStyle-Width="5%"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="Title1" HeaderText="職稱" ItemStyle-Width="20%"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="CName" HeaderText="姓名" ItemStyle-Width="20%"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="Phone1" HeaderText="聯絡電話" ItemStyle-Width="20%"></asp:BoundColumn>
                                                    <asp:TemplateColumn>
                                                        <HeaderTemplate>申請階段</HeaderTemplate><%--年度別--%>
                                                        <HeaderStyle Width="15%" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="LabFSQ1_DG" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn>
                                                        <HeaderTemplate>功能</HeaderTemplate>
                                                        <HeaderStyle Width="10%" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="hidOMID" type="hidden" runat="server" />
                                                            <input id="HidROWNUM1" type="hidden" runat="server" />
                                                            <asp:Button ID="btEdit" runat="server" Text="修改" CommandName="edt" CssClass="asp_button_M"></asp:Button>
                                                            <asp:Button ID="btDel" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                            <asp:Button ID="btLOCK" runat="server" Text="解鎖" CommandName="lock1" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <%--<input id="button5" type="button" value="列印空白表單" name="button5" runat="server" class="asp_button_M" onclick="return button5_onclick()">--%>
                                <%--<input id="Button1" type="button" value="列印" name="Button1" runat="server" class="asp_button_S">--%>
                                <asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Label ID="lMsg2" runat="server" ForeColor="Red">※依選擇的「排序方式」列印專職人員順序</asp:Label></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
