<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="TC_01_016.aspx.vb" Inherits="WDAIIP.TC_01_016" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構計畫調動</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function Search() {
            //debugger;
            var city_code = document.getElementById('city_code');
            var TB_ComIDNO = document.getElementById('TB_ComIDNO');
            var TB_OrgName = document.getElementById('TB_OrgName');

            if (city_code.value == '' && TB_ComIDNO.value == '' && TB_OrgName.value == '') {
                alert('請輸入機構名稱、統一編號或者是縣市代碼');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;訓練機構計畫調動</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td id="td6" runat="server" class="bluecol" width="15%">機構名稱</td>
                <td class="whitecol" width="35%">
                    <asp:TextBox ID="TB_OrgName" runat="server" MaxLength="150" Columns="30" Width="95%"></asp:TextBox></td>
                <td id="td7" runat="server" class="bluecol" width="15%">統編</td>
                <td class="whitecol" width="35%">
                    <asp:TextBox ID="TB_ComIDNO" runat="server" MaxLength="20" Width="40%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol" width="15%">縣市</td>
                <td colspan="3" width="85%">
                    <div align="left" class="whitecol">
                        <asp:TextBox ID="TBCity" runat="server" onfocus="this.blur()" Columns="30" Width="45%"></asp:TextBox>
                        <input id="city_zip" onclick="getZip('../../js/Openwin/zipcode_search.aspx', 'TBCity', 'zip_code', 'city_code')" type="button" value="..." name="city_zip" runat="server" class="button_b_Mini" />
                        <input id="zip_code" type="hidden" name="zip_code" runat="server" />
                        <input id="city_code" type="hidden" name="city_code" runat="server" />
                    </div>
                </td>
            </tr>
            <tr id="DISTTR">
                <td class="bluecol" width="15%">轄區分署</td>
                <td colspan="3" class="whitecol" width="85%">
                    <asp:DropDownList ID="DistID" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" width="15%">機構別</td>
                <td colspan="3" class="whitecol" width="85%">
                    <asp:DropDownList ID="OrgKindList" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">年度</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="Years" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" width="15%">計畫別</td>
                <td colspan="3" class="whitecol" width="85%">
                    <asp:DropDownList ID="drpPlan" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" width="15%">計畫審核狀態</td>
                <td colspan="3" class="whitecol" width="85%">
                    <asp:DropDownList ID="drpAppliedResult" runat="server"></asp:DropDownList></td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol">
                    <div align="center">
                        <asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <br>
        <asp:Panel ID="Panel" runat="server" Visible="False" Width="100%">
            <%--<table id="search_tbl" class="font" border="1" cellspacing="0" cellpadding="0" width="100%" runat="server"></table>--%>
            <asp:DataGrid ID="DG_Org" runat="server" Width="100%" CssClass="font" Visible="False" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="編號">
                        <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區中心">--%>
                    <asp:BoundColumn DataField="DistName" HeaderText="轄區分署">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PlanName" HeaderText="計畫名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ComIDNO" HeaderText="統編">
                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Address" HeaderText="地址">
                        <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ActNo" HeaderText="保險證號">
                        <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ContactName" HeaderText="聯絡人姓名">
                        <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ContactEmail" HeaderText="聯絡人E-Mail">
                        <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                        <ItemTemplate>
                            <asp:LinkButton ID="lbtModify" runat="server" Text="計畫調動" CommandName="modify" CssClass="linkbutton"></asp:LinkButton>
                            <%--
                                <asp:Button id="share_but" runat="server" Text="共用" CommandName="share"></asp:Button>
                                <asp:Button id="edit_but" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                <asp:Button id="del_but" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                <asp:Button id="chk_but" runat="server" Text="審核" CommandName="chk"></asp:Button>
                                <asp:Button id="year_btn" runat="server" Width="90px" Text="年度對應功能" Visible="False" CommandName="year"></asp:Button>
                            --%>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <div align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </div>
        </asp:Panel>
        <%--
        <input id="check_add" type="hidden" name="check_add" runat="server">
        <input id="check_del" type="hidden" name="check_del" runat="server">
        <input id="check_mod" type="hidden" name="check_mod" runat="server">
        --%>
    </form>
</body>
</html>
