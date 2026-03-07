<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_003.aspx.vb" Inherits="WDAIIP.TC_01_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班別代碼設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;班別代碼設定</asp:Label>
                </td>
            </tr>
        </table>
        <input id="check_del" style="width: 64px; height: 22px" type="hidden" size="5" name="check_del" runat="server">
        <input id="check_mod" style="width: 64px; height: 22px" type="hidden" size="5" name="check_mod" runat="server">
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" width="20%">訓練計畫</td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="Plan_List" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">計畫年度</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlYears" runat="server"></asp:DropDownList>
                </td>
                <td class="bluecol">轄區分署 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" id="td6" runat="server" class="bluecol">班別代碼</td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="TB_classid" runat="server" MaxLength="50" Width="99%"></asp:TextBox></td>
                <td width="20%" id="td7" runat="server" class="bluecol">班別名稱</td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="TB_ClassName" runat="server" MaxLength="100" Width="99%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">
                    <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="38%"></asp:TextBox>
                    <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server">
                    <input id="trainValue" type="hidden" name="trainValue" runat="server">
                    <input id="jobValue" type="hidden" name="jobValue" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">
                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="38%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>&nbsp;
				<asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
				<asp:Button ID="bt_add" Text="新增" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
				<asp:Button ID="print" runat="server" Text="列印-班別代碼表" CssClass="asp_Export_M"></asp:Button>&nbsp;
				<input id="TPlanID" type="hidden" name="TPlanID" runat="server">
                    <asp:Button ID="Button1" runat="server" Text="匯入班別代碼" CssClass="asp_Export_M"></asp:Button>
                    <div align="center">&nbsp;</div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <br>
        <asp:Panel ID="Panel" runat="server" Width="100%" Visible="False">
            <%--<table id="search_tbl" class="font" border="1" cellspacing="0" cellpadding="0" width="90%" runat="server"></table>--%>
            <asp:DataGrid ID="DG_Class" runat="server" Width="100%" CssClass="font" Visible="False" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="序號">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassID" HeaderText="班別代碼">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類／業別">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CJOB_NAME" HeaderText="通俗職類">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                        <ItemTemplate>
                            <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtCopy" runat="server" Text="複製" CommandName="copy" CssClass="linkbutton"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <div align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </div>
        </asp:Panel>
    </form>
</body>
</html>
