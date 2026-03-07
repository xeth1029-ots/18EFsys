<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_014.aspx.vb" Inherits="WDAIIP.SD_05_014" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤作業(產投)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">

        function GETvalue() { document.getElementById('Button6').click(); }

        function SetOneOCID() { document.getElementById('Button7').click(); }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            var OCID1 = document.getElementById('OCID1');
            var DataGridTable = document.getElementById('DataGridTable');

            DataGridTable.style.display = 'none';
            if (OCID1.value == '') { document.getElementById('Button7').click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學員出缺勤作業</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
<%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
<tr><td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員出缺勤作業</font> </td></tr></table>--%>
                    <div>
                        <table class="table_sch" id="Table2">
                            <tr>
                                <td class="bluecol" width="20%">訓練機構 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="55%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini">
                                    <asp:Button ID="Button7" Style="display: none" runat="server"></asp:Button>
                                    <asp:Button ID="Button6" Style="display: none" runat="server" Text="Button6"></asp:Button>
                                    <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">職類/班別 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input onclick="choose_class();" type="button" value="..." class="asp_button_Mini">
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                    <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span></td>
                            </tr>
                            <tr>
                                <td class="bluecol">通俗職類 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini">
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">未出席日期 </td>
                                <td colspan="3" class="whitecol">
                                    <span id="span1" runat="server">
                                        <asp:TextBox ID="start_date" Width="15%" runat="server" MaxLength="20"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
								    <asp:TextBox ID="end_date" Width="15%" runat="server" MaxLength="20"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;<asp:Button ID="Button5" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" class="whitecol">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div>
                        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="17%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班別" HeaderStyle-Width="17%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="StudentID" HeaderText="學號" HeaderStyle-Width="17%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Name" HeaderText="姓名" HeaderStyle-Width="17%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="LeaveDate" HeaderText="未出席日期" DataFormatString="{0:d}" HeaderStyle-Width="17%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="時數(不列入)" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labHours" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:Button ID="Button3" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="Button4" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
