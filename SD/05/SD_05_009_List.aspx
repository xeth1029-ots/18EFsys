<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_009_List.aspx.vb" Inherits="WDAIIP.SD_05_009_List" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>鍾點費試算-查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript">
        function print() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇班級職類\n';
            if (document.form1.years.selectedIndex == 0) msg += '請選擇年度\n';
            if (document.form1.months.selectedIndex == 0) msg += '請選擇月份\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function check() {
            var msg = '';
            if (getValue('RB_Teacher_List') == 0) msg += '請選取講師名稱\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //'檢查日期格式-Melody(2005/3/28)
        function check_date() {
            if (!checkDate(form1.Text_Date.value)) {
                alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }
        }
		
    </script>
    <style type="text/css">
        .style1
        {
            color: #ffffff;
        }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <asp:Panel ID="Panel1" runat="server" CssClass="font" Visible="False" Width="168px">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
    </asp:Panel>
    <table id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">鍾點費試算-查詢</font>
                        </td>
                    </tr>
                </table>
                <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td style="width: 75px" width="75" bgcolor="#2aafc0">
                            &nbsp;<font color="#ffffff">&nbsp; 月份</font><font color="red">*</font>
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3">
                            <asp:DropDownList ID="years" runat="server" Width="96px">
                            </asp:DropDownList>
                            <font face="新細明體">年</font>
                            <asp:DropDownList ID="months" runat="server" Width="88px">
                            </asp:DropDownList>
                            <font face="新細明體">月</font>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 75px" bgcolor="#2aafc0">
                            <font color="#ffffff">&nbsp;&nbsp; 職類/班別</font><font color="#ff0000">*</font>
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3">
                            <asp:TextBox ID="TMID1" runat="server" Width="210px" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" Width="210px" onfocus="this.blur()"></asp:TextBox><input onclick="window.open('../02/SD_02_ch.aspx','','width=540,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');" type="button" value="..."><input id="TMIDValue1" style="width: 35px; height: 22px" type="hidden" name="Hidden2" runat="server"><input id="OCIDValue1" style="width: 40px; height: 22px" type="hidden" name="Hidden1" runat="server">
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 75px; height: 21px" bgcolor="#2aafc0">
                            &nbsp;<span class="style1">&nbsp; 講師姓名</span>
                        </td>
                        <td style="width: 216px; height: 21px" bgcolor="#ecf7ff">
                            <asp:TextBox ID="TeacherName" runat="server" Width="142px"></asp:TextBox>
                        </td>
                        <td style="width: 91px; height: 21px" bgcolor="#2aafc0">
                            <span class="style1">&nbsp;&nbsp; 講師代碼</span>
                        </td>
                        <td style="height: 21px" bgcolor="#ecf7ff">
                            <asp:TextBox ID="TeacherID" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 75px" bgcolor="#2aafc0">
                            <span class="style1">&nbsp;&nbsp; 內外聘</span>&nbsp;
                        </td>
                        <td style="width: 216px" bgcolor="#ecf7ff">
                            <asp:DropDownList ID="DropDownList4" runat="server" AutoPostBack="True">
                                <asp:ListItem Value="0">--請選擇--</asp:ListItem>
                                <asp:ListItem Value="1">內聘</asp:ListItem>
                                <asp:ListItem Value="2">外聘</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                        <td style="width: 91px" bgcolor="#2aafc0">
                            <span class="style1">&nbsp;&nbsp; 主要職類</span>
                        </td>
                        <td bgcolor="#ecf7ff">
                            <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()"></asp:TextBox><input id="trainValue" style="width: 16px; height: 22px" type="hidden" name="trainValue" runat="server"><input onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="...">
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4">
                            <p align="center">
                                <asp:Button ID="Button1" runat="server" Text="查詢"></asp:Button>&nbsp;&nbsp;
                                <asp:Button ID="Button2" runat="server" Text="回試算頁"></asp:Button></p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br>
    <br>
    <br>
    <asp:Panel ID="Panel" runat="server" CssClass="font" Visible="False" Width="600px" Height="93px">
        <font face="新細明體">請選擇講師名稱，查看授課月份的授課時數</font>
        <asp:RadioButtonList ID="RB_Teacher_List" runat="server" Width="100%" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="7">
        </asp:RadioButtonList>
        <asp:Button ID="count_Button" runat="server" Text="查詢"></asp:Button></p>
    </asp:Panel>
    <font face="新細明體">
        <asp:Label ID="msg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
    </font>
    <asp:Panel ID="Panel2" runat="server" CssClass="font" Visible="False" Width="224px">
        <table id="Table4" style="width: 208px; height: 228px" cellspacing="1" cellpadding="1" width="208" border="1">
            <tr>
                <td colspan="3">
                    講師:
                    <asp:Label ID="Label1" runat="server"></asp:Label>
                    <input id="Add_Button" type="button" value="新增" name="Add_Button" runat="server">
                    <input id="openwin" type="hidden" value="0" name="openwin" runat="server">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:DataGrid ID="DG_Teacher" runat="server" CssClass="font" BorderColor="Black" AutoGenerateColumns="False" ShowFooter="True" ShowHeader="False">
                        <Columns>
                            <asp:TemplateColumn>
                                <HeaderStyle BackColor="#99FFCC"></HeaderStyle>
                                <ItemStyle BackColor="#CCFFFF"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="Label3" runat="server">單價</asp:Label>
                                    <asp:TextBox ID="Text_Price" runat="server" Columns="6"></asp:TextBox>
                                    <asp:DataGrid ID="DG_Prices" runat="server" CssClass="font" ShowFooter="True" AutoGenerateColumns="False" OnDeleteCommand="DG_Prices_DeleteCommand" OnItemDataBound="DG_Prices_ItemDataBound">
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="日期">
                                                <HeaderStyle BackColor="#2aafc0"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="Text_Date" onfocus="this.blur()" Columns="9" runat="server" Text='<%# Common.FormatDate(DataBinder.Eval(Container.DataItem, "TeachDate")) %>'>
                                                    </asp:TextBox>
                                                    <a href="" id="linkDate" runat="server">
                                                        <img border="0" alt="" src="../../images/show-calendar.gif" width="30" height="30"></a>
                                                </ItemTemplate>
                                                <FooterTemplate>
                                                    小計:
                                                    <asp:Label ID="lblSum" runat="server"></asp:Label>
                                                </FooterTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="時數">
                                                <HeaderStyle BackColor="#2aafc0"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="Text_hour" Columns="3" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "UnitHour") %>'>
                                                    </asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:ButtonColumn Text="刪除" ButtonType="PushButton" CommandName="Delete">
                                                <HeaderStyle BackColor="#2aafc0"></HeaderStyle>
                                            </asp:ButtonColumn>
                                            <asp:BoundColumn Visible="False" DataField="UnitPrice"></asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="TeachDate"></asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="Teach_Pay_ID"></asp:BoundColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </ItemTemplate>
                                <FooterStyle BackColor="#CCFFCC"></FooterStyle>
                                <FooterTemplate>
                                    <asp:Label ID="Label2" runat="server">總計:</asp:Label>
                                    <asp:Label ID="lblAllSum" runat="server"></asp:Label>
                                </FooterTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="save_Button" runat="server" Text="儲存"></asp:Button>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <input id="Year_str" style="width: 48px; height: 22px" type="hidden" size="2" name="Year_str" runat="server">
    <input id="TechID_str" style="width: 48px; height: 22px" type="hidden" size="2" name="TechID_str" runat="server">
    <input id="Month_str" style="width: 40px; height: 22px" type="hidden" name="Month_str" runat="server">
    </form>
</body>
</html>
