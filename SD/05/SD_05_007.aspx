 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_007.aspx.vb" Inherits="WDAIIP.SD_05_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_05_007</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">

        function GETvalue() {
            document.getElementById('Button7').click();
        }
        function SetOneOCID() {
            document.getElementById('Button8').click();
        }
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            if (document.getElementById('OCID1').value == '')
            { document.getElementById('Button8').click(); }
            openClass('../02/SD_02_ch.aspx?special=2&RID=' + RID);
        }
        function search() {
            var msg = '';
            if (document.form1.start_date.value != '' && !checkDate(document.form1.start_date.value)) msg += '起始日期時間格式不正確\n';
            if (document.form1.end_date.value != '' && !checkDate(document.form1.end_date.value)) msg += '終至日期時間格式不正確\n';

            if (msg != '') {
                alert(msg);
                return false
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員獎懲作業</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="100" class="bluecol">
                                訓練機構
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
                                <input id="Button6" type="button" value="..." name="Button6" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <asp:Button ID="Button8" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">
                                職類/班別
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                <input onclick="choose_class();" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                通俗職類
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">
                                獎懲日期
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="75px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ～<asp:TextBox ID="end_date" runat="server" Width="75px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                            <td width="100" class="bluecol">
                                獎懲別
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="SanID" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <p align="center">
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                                    <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button></p>
                            </td>
                        </tr>
                    </table>
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SanID" HeaderText="獎懲">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Times" HeaderText="獎懲數目">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SanDate" HeaderText="獎懲日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="Button3" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                <asp:Button ID="Button4" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                <asp:Button ID="Button5" runat="server" Text="查詢" CommandName="view"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="SOCID" HeaderText="SOCID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="SeqNo" HeaderText="SeqNo"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
                </td>
            </tr>
        </table>
    </font>
    </form>
</body>
</html>
