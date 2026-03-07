<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_002.aspx.vb" Inherits="WDAIIP.SD_08_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_08_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('GetClass.aspx?RID=' + RID + '&SubmitBtn=Button1&OCIDField=OCIDValue');
        }

        function search() {
            if (document.form1.OCIDValue.value == '') {
                alert('必須選擇職類班別!');
                return false;
            }
        }
        function SelectAll() {
            var MyTable = document.getElementById('DataGrid2');

            for (i = 1; i < MyTable.rows.length; i++) {
                if (!MyTable.rows(i).cells(10).children(0).disabled)
                    MyTable.rows(i).cells(10).children(0).value = MyTable.rows(0).cells(10).children(0).value;
            }
        }
    </script>
</head>
<body>
    <font face="新細明體"></font>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="600" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;學員動態管理&gt;&gt;職業訓練生活津貼&gt;&gt;<font color="#990000">職業訓練生活津貼核准</font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="SearchTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td width="100" bgcolor="#2aafc0">
                                            <font color="#ffffff">&nbsp;&nbsp;&nbsp; 訓練機構</font>
                                        </td>
                                        <td bgcolor="#ecf7ff">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server" size="1"><input type="button" value="..." id="Button6" name="Button6" runat="server">
                                            <asp:Button ID="Button7" runat="server" Text="查詢上一次的列表" Style="display: none"></asp:Button><br>
                                            <span id="HistoryList2" style="position: absolute; display: none">
                                                <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                                </asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="100" bgcolor="#2aafc0">
                                            <font color="#ffffff">&nbsp;&nbsp;&nbsp; 職類/班別</font><font color="#ff0000">*</font>
                                        </td>
                                        <td bgcolor="#ecf7ff">
                                            <input id="OCIDValue" type="hidden" name="OCIDValue" runat="server" size="1"><input onclick="choose_class()" type="button" value="挑選班級">
                                            <asp:Button ID="Button1" runat="server" Text="查詢"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
                                                <ItemStyle BackColor="White"></ItemStyle>
                                                <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="初審通過數"></asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="待複審數"></asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="複審通過數"></asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="複審未通過數"></asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <ItemTemplate>
                                                            <asp:Button ID="Button4" runat="server" Text="查核" CommandName="view"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                </table>
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="DetailTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:Label ID="ClassName" runat="server"></asp:Label><input id="NowOCID" type="hidden" name="NowOCID" runat="server" size="1">
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
                        <ItemStyle BackColor="White"></ItemStyle>
                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="學員">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TrainingMonth" HeaderText="受訓月數">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SumOfMoney" HeaderText="金額">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ApplyDate" HeaderText="申請日" DataFormatString="{0:d}">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="UnitCode" HeaderText="單位代碼">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TrainCode" HeaderText="職類代碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="FailReasonF" HeaderText="初審備註">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="申請狀態">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="審核">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    審核
                                <asp:DropDownList ID="AppliedAll" runat="server">
                                    <asp:ListItem>請選擇</asp:ListItem>
                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                </asp:DropDownList>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:DropDownList ID="AppliedStatusS" runat="server">
                                        <asp:ListItem>請選擇</asp:ListItem>
                                        <asp:ListItem Value="Y">通過</asp:ListItem>
                                        <asp:ListItem Value="N">不通過</asp:ListItem>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="複審備註">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="FailReasonS" runat="server" Width="150px" TextMode="MultiLine"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button2" runat="server" Text="存檔"></asp:Button><asp:Button ID="Button3" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button><asp:Button ID="Button5" runat="server" Text="回上一頁"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
