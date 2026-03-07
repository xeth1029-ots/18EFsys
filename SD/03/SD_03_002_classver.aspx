<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002_classver.aspx.vb" Inherits="WDAIIP.SD_03_002_classver" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料審核</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>

    <script type="text/javascript" language="javascript">
        //<!--
        var cst_supplyID = 13//預算別
        var cst_AppliedResult = 17//審核
        //function ChangeAll(col, j) {
        //    debugger;
        //    var MyTable = document.getElementById('DataGrid1');
        //    for (i = 1; i < MyTable.rows.length; i++) {
        //        MyTable.rows[i].cells[col].children[0].selectedIndex = j;
        //    }
        //}

        //全部置換
        function ChangeAll(type, objAllID) {
            //debugger;
            $("select[Name$='" + type + "']").each(function () {
                $(this).val($("select#" + objAllID).val());
            });
        }
        //-->
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" border="0" width="100%">
            <tbody>
                <%--<tr>
                    <td colspan="2">
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;<font color="#990000"> 學員資料審核作業(產學訓專用)</font></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>--%>
                <tr>
                    <td colspan="2" class="whitecol">
                        <div align="center">
                            <asp:Button ID="Button4" runat="server" Text="列印" CssClass="asp_button_M" Style="display: none"></asp:Button>
                            <asp:Button ID="Button11" runat="server" Text="儲存" ToolTip="學員資料審核" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="Button1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="Button2" runat="server" Text="儲存" ToolTip="選擇還原審核" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label1" runat="server" ForeColor="Blue">#表示該學員審核結果尚未選取為「請選擇」狀態</asp:Label><br>
                        <asp:Label ID="BIEPTBL" runat="server" ForeColor="Red">* 表示為該學員已申請失業給付，可點選檢示功能查詢</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <div align="center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                    </td>
                </tr>
                <tr>
                    <td align="left" width="50%">
                        <asp:Label ID="OrgName" runat="server" ForeColor="Blue">報名機構:</asp:Label>&nbsp;<asp:Label ID="OrgName1" runat="server"></asp:Label></td>
                    <td align="left">
                        <asp:Label ID="OCIDName" runat="server" ForeColor="Blue">報名班級:</asp:Label>&nbsp;<asp:Label ID="OCIDName1" runat="server"></asp:Label></td>
                </tr>
            </tbody>
        </table>
        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowCustomPaging="True" AutoGenerateColumns="False" AllowSorting="True" Width="100%" CellPadding="8">
            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
            <%--<ItemStyle BackColor="#ECF7FF"></ItemStyle>--%>
            <HeaderStyle HorizontalAlign="Center" ForeColor="White" CssClass="head_navy"></HeaderStyle>
            <Columns>
                <asp:TemplateColumn HeaderText="序號">
                    <HeaderStyle Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                    <ItemTemplate>
                        <asp:Label ID="star2" runat="server" ForeColor="Blue">#</asp:Label>
                        <asp:Label ID="star1" runat="server" ForeColor="Red">*</asp:Label>
                        <asp:Label ID="IDLab" runat="server"></asp:Label>
                    </ItemTemplate>
                    <FooterStyle Wrap="False"></FooterStyle>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="姓名">
                    <HeaderStyle Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" Width="60px"></ItemStyle>
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue" CommandName="link"></asp:LinkButton>
                    </ItemTemplate>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:TemplateColumn>
                <asp:BoundColumn DataField="IDNO" SortExpression="IDNO" HeaderText="身分證號碼">
                    <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="IdentityName" HeaderText="身分別">
                    <HeaderStyle Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" Width="120px"></ItemStyle>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="ActNo" HeaderText="投保單位<br>保險證號">
                    <HeaderStyle Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn HeaderText="預算別">
                    <HeaderStyle CssClass="whitecol" />
                    <ItemStyle CssClass="whitecol" />
                    <HeaderTemplate>
                        預算別<br>
                        <asp:DropDownList ID="HBudID_all" runat="server">
                            <asp:ListItem Value="">請選擇</asp:ListItem>
                            <asp:ListItem Value="01">公務</asp:ListItem>
                            <asp:ListItem Value="02">就安</asp:ListItem>
                            <asp:ListItem Value="03">就保</asp:ListItem>
                            <asp:ListItem Value="99">不補助</asp:ListItem>
                        </asp:DropDownList>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:DropDownList ID="HBudID" runat="server">
                            <asp:ListItem Value="">請選擇</asp:ListItem>
                            <asp:ListItem Value="01">公務</asp:ListItem>
                            <asp:ListItem Value="02">就安</asp:ListItem>
                            <asp:ListItem Value="03">就保</asp:ListItem>
                            <asp:ListItem Value="99">不補助</asp:ListItem>
                        </asp:DropDownList>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:BoundColumn DataField="supplyID" HeaderText="補助比例">
                    <HeaderStyle Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn HeaderText="備註">
                    <HeaderStyle Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <ItemTemplate>
                        <asp:TextBox ID="signUpMemo" Width="120px" runat="server" TextMode="MultiLine" MaxLength="128"></asp:TextBox>
                    </ItemTemplate>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="審核結果">
                    <HeaderStyle Wrap="False" CssClass="whitecol"></HeaderStyle>
                    <ItemStyle Wrap="False" CssClass="whitecol"></ItemStyle>
                    <HeaderTemplate>
                        審核結果<br>
                        <asp:DropDownList ID="SelectAll" runat="server">
                            <asp:ListItem Value="M">=請選擇=</asp:ListItem>
                            <asp:ListItem Value="Y">審核通過</asp:ListItem>
                            <asp:ListItem Value="N">不補助</asp:ListItem>
                            <asp:ListItem Value="R">退件修正</asp:ListItem>
                        </asp:DropDownList>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:DropDownList ID="AppliedResult2" runat="server">
                            <asp:ListItem Value="M">=請選擇=</asp:ListItem>
                            <asp:ListItem Value="Y">審核通過</asp:ListItem>
                            <asp:ListItem Value="N">不補助</asp:ListItem>
                            <asp:ListItem Value="R">退件修正</asp:ListItem>
                        </asp:DropDownList>
                        <input id="KeyValue" type="hidden" name="KeyValue" runat="server">
                        <asp:HiddenField ID="Hid_SOCID1" runat="server" />
                        <asp:Button ID="Btn_VIEW" runat="server" Text="檢視" CommandName="view" CssClass="asp_button_M"></asp:Button>
                    </ItemTemplate>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="還原審核">
                    <HeaderStyle Wrap="False" CssClass="whitecol"></HeaderStyle>
                    <ItemStyle Wrap="False" CssClass="whitecol"></ItemStyle>
                    <HeaderTemplate>
                        還原審核<br>
                        <asp:DropDownList ID="SelectAllR" runat="server">
                            <asp:ListItem Value="M">=請選擇=</asp:ListItem>
                            <asp:ListItem Value="R">還原審核</asp:ListItem>
                        </asp:DropDownList>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:DropDownList ID="AppliedResult2_R" runat="server">
                            <asp:ListItem Value="M">=請選擇=</asp:ListItem>
                            <asp:ListItem Value="R">還原審核</asp:ListItem>
                        </asp:DropDownList>
                        <input id="KeyValueR" type="hidden" name="KeyValueR" runat="server" />
                        <asp:HiddenField ID="Hid_SOCID2" runat="server" />
                    </ItemTemplate>
                    <FooterStyle Wrap="False"></FooterStyle>
                </asp:TemplateColumn>
            </Columns>
            <PagerStyle Visible="False"></PagerStyle>
        </asp:DataGrid>
    </form>
</body>
</html>
