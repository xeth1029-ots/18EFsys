<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_012.aspx.vb" Inherits="TIMS.TC_01_012" %>

<%@ Register TagPrefix="uc1" TagName="PageControler" Src="../../PageControler.ascx" %>
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
        /*
        function but_edit(orgid,planid,rid,distid,id) {
        location.href = 'TC_01_012_add.aspx?orgid='+orgid+'&planid='+planid+'&rid='+rid+'&ProcessType=Update&distid='+distid+'&ID='+id;
        }
        */
        function but_del(orgid, account, classid, rid, planid, is_parent, id) {
            if (is_parent) {
                alert("此機構尚有下層單位,不可刪除!!");
                return;
            }

            if (classid == "" & account == "") {
                if (window.confirm("此動作會刪除機構資料，是否確定刪除?")) {
                    location.href = 'TC_01_012_del.aspx?orgid=' + orgid + '&rid=' + rid + '&planid=' + planid + '&ID=' + id;
                }
            } else if (classid != "") {
                alert('此機構已有開班資料，不可以刪除!!');
            } else {
                alert('此機構已有帳號資料，不可以刪除!!');
            }
        }

        function Search() {
            //alert(document.form1.yearlist.value);
            if (document.form1.yearlist.value == '' || (document.form1.OrgKindList.value == '' && document.form1.city_code.value == '' && document.form1.TB_ComIDNO.value == '' && document.form1.TB_OrgName.value == '')) {
                alert('請輸入年度、機構名稱、統一編號或者是縣市代碼');
                return false;
            }
        }

        /*			
        function but_share(orgid,planid,rid,distid,id) {
        location.href = 'TC_01_012_add.aspx?orgid='+orgid+'&ProcessType=Share&distid='+distid+'&planid='+planid+'&rid='+rid+'&ID='+id;
        }
        */
			
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" width="608" style="width: 608px; height: 22px">
        <tr>
            <td class="font">
                首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">訓練機構評鑑設定</font>
            </td>
        </tr>
    </table>
    <table class="table_nw" width="740" cellpadding="1" cellspacing="1">
        <tr>
            <td id="Td1" width="100" runat="server" class="bluecol_need">
                年度
            </td>
            <td colspan="3" class="whitecol">
                <asp:DropDownList ID="yearlist" runat="server">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td id="td6" width="100" runat="server" class="bluecol">
                機構名稱
            </td>
            <td class="whitecol">
                <asp:TextBox ID="TB_OrgName" runat="server" MaxLength="30" Columns="30"></asp:TextBox>
            </td>
            <td id="td7" width="100" runat="server" class="bluecol">
                統編
            </td>
            <td class="whitecol">
                <asp:TextBox ID="TB_ComIDNO" runat="server" Width="88px" MaxLength="10"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td class="bluecol">
                縣市
            </td>
            <td colspan="3" class="whitecol">
                <asp:TextBox ID="TBCity" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                <input id="city_zip" onclick="getZip('../../js/Openwin/zipcode_search.aspx', 'TBCity', 'zip_code','city_code')" type="button" value="..." name="city_zip" runat="server" class="button_b_Mini">
            </td>
        </tr>
        <tr>
            <td class="bluecol">
                機構別
            </td>
            <td colspan="3" class="whitecol">
                <asp:DropDownList ID="OrgKindList" runat="server">
                </asp:DropDownList>
            </td>
        </tr>
    </table>
    <table width="740">
        <tr>
            <td class="whitecol" align="center">
                <asp:Label ID="labPageSize" runat="server" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label>
                <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>
                <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                <asp:Button ID="bt_add" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;
                <asp:Button ID="bt_save" runat="server" Text="匯出" CssClass="asp_button_S" Visible="False"></asp:Button>&nbsp;
                <input id="city_code" style="width: 26px; height: 22px" type="hidden" name="city_code" runat="server">
                <input id="check_add" style="width: 40px; height: 22px" type="hidden" size="1" name="check_add" runat="server">
                <input id="check_del" style="width: 48px; height: 22px" type="hidden" size="2" name="check_del" runat="server">
                <input id="check_mod" style="width: 45px; height: 22px" type="hidden" size="2" name="check_mod" runat="server">
                <input id="zip_code" style="width: 26px; height: 22px" type="hidden" name="zip_code" runat="server">
                <asp:Button ID="Button3" runat="server" Text="匯入年度評鑑" CssClass="asp_button_M"></asp:Button>
                <div align="center">
                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></div>
            </td>
        </tr>
    </table>
    <asp:Panel ID="Panel" runat="server" Width="81.16%" Visible="False">
        <table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="600" border="1" runat="server">
        </table>
        <asp:DataGrid ID="DG_Org" runat="server" Width="100%" Visible="False" CssClass="font" AllowPaging="True" AutoGenerateColumns="False">
            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
            <HeaderStyle CssClass="head_navy"></HeaderStyle>
            <Columns>
                <asp:BoundColumn HeaderText="編號">
                    <HeaderStyle HorizontalAlign="Center" Width="30px"></HeaderStyle>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="PlanYear" HeaderText="年度">
                    <HeaderStyle HorizontalAlign="Center" Width="60px"></HeaderStyle>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="name" HeaderText="轄區中心">
                    <HeaderStyle HorizontalAlign="Center" Width="100px"></HeaderStyle>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱"></asp:BoundColumn>
                <asp:BoundColumn DataField="CHARNAME" HeaderText="訓練性質">
                    <HeaderStyle Width="120px"></HeaderStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn HeaderText="功能">
                    <HeaderStyle HorizontalAlign="Center" Width="150px"></HeaderStyle>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    <ItemTemplate>
                        <asp:Button ID="add_but" runat="server" Text="新增" CommandName="edit"></asp:Button>
                        <asp:Button ID="edit_but" runat="server" Text="修改" CommandName="edit"></asp:Button>
                        <asp:Button ID="del_but" runat="server" Text="刪除" CommandName="del"></asp:Button>
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
            <PagerStyle Visible="False"></PagerStyle>
        </asp:DataGrid>
        <font face="新細明體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
        <div align="center">
            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
        </div>
    </asp:Panel>
    </form>
</body>
</html>
