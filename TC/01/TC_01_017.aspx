<%@ Page Language="vb" AutoEventWireup="false" EnableEventValidation="true" CodeBehind="TC_01_017.aspx.vb" Inherits="WDAIIP.TC_01_017" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構屬性設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <%--<script type="text/javascript" src="../../js/selectControl.js.aspx" charset="UTF-8"></script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        /**
       
		function SelPlan(selectedID) {
		var selValue = '0';
		if (document.getElementById('PlanPoint_0').checked) selValue = '0';
		if (document.getElementById('PlanPoint_1').checked) selValue = '1';
		if (document.getElementById('PlanPoint_2').checked) selValue = '2';
		if (selValue != '0') {
		var parms = "[['TypeID1','" + selValue + "']]";      // 透過 selectControl 傳遞給 SQLMap 的年度查詢條件, 格式請參考 selectControl 定義說明
		selectControl('ajaxQueryKeyOrgTypeS2', 'dl_typeid2', 'TypeID2Name', 'TypeID2', '請選擇', selectedID, parms);
		}
		else {
		var obj = document.getElementById('dl_typeid2');
		obj.length = 1;
		}
		}
		**/

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" width="100%">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">訓練機構屬性設定</font></asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%">
            <tr>
                <td width="15%" class="bluecol">機構名稱
                </td>
                <td width="35%" class="whitecol">
                    <asp:TextBox ID="tb_orgname" runat="server" MaxLength="30" Columns="30"></asp:TextBox>
                </td>
                <td width="15%" class="bluecol">統編
                </td>
                <td width="35%" class="whitecol">
                    <asp:TextBox ID="tb_comidno" runat="server" Width="88px" MaxLength="10"></asp:TextBox>
                </td>
            </tr>
            <tr id="TRPlanPoint28" runat="server">
                <td class="bluecol">計畫
                </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" AutoPostBack="True">
                        <asp:ListItem Value="0" Selected="True">不區分</asp:ListItem>
                        <asp:ListItem Value="1">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">機構別</td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="dl_typeid2" runat="server" Width="150px">
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol">
                    <div align="center">
                        <asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_S"></asp:Button>&nbsp;
                    </div>
                    <div align="center">
                        &nbsp;
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <br>
        <asp:Panel ID="Panel" runat="server" Width="100%">
            <table id="search_tbl" class="font" border="0" cellspacing="0" cellpadding="0" width="740" runat="server">
            </table>
            <asp:DataGrid ID="dg1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="編號">
                        <HeaderStyle HorizontalAlign="Center" Width="30px"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="orgname" HeaderText="機構名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="40%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="comidno" HeaderText="統編">
                        <HeaderStyle HorizontalAlign="Center" Width="60px"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="plantype" HeaderText="計畫別">
                        <HeaderStyle HorizontalAlign="Center" Width=""></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="orgtype" HeaderText="機構別">
                        <HeaderStyle HorizontalAlign="Center" Width="100px"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <ItemTemplate>
                            <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
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
