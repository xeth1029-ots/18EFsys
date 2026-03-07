<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_06_004.aspx.vb" Inherits="WDAIIP.SD_06_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_06_004</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //檢查匯出
        function chkExport() {
            var msg = '';
            var txtSDate = document.getElementById('txtSDate');
            var txtEDate = document.getElementById('txtEDate');

            if (txtSDate.value == '') msg += '請輸入日期起日!\n';
            else if (!checkDate(txtSDate.value)) msg += '日期起日輸入格式有誤!\n';

            if (txtEDate.value == '') msg += '請輸入日期迄日!\n';
            else if (!checkDate(txtEDate.value)) msg += '日期迄日輸入格式有誤!\n';

            if (msg == '') {
                if (compareDate(txtSDate.value, txtEDate.value) > 0) msg += '日期起日不可大於日期迄日!\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;學員動態管理&gt;&gt;加退保管理&gt;&gt;<font color="#990000">投保資料查詢</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol_need">日期區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtSDate" runat="server" Width="100px"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('txtSDate','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                ~
									<asp:TextBox ID="txtEDate" runat="server" Width="100px"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('txtEDate','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">狀態</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblType" RepeatColumns="6" CssClass="font" runat="server">
                                    <asp:ListItem Value="">不區分</asp:ListItem>
                                    <asp:ListItem Value="0">加保</asp:ListItem>
                                    <asp:ListItem Value="1">退保</asp:ListItem>
                                    <asp:ListItem Value="2">無異動</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="btnExport" Text="匯出" runat="server" CssClass="asp_Export_M"></asp:Button>
                    <br />
                    <br style="line-height: 5px" />
                    <asp:Label ID="labMsg" Style="color: red" CssClass="font" Visible="False" runat="server">查無資料</asp:Label>
                </td>
            </tr>
        </table>
        <div id="div1" runat="server">
            <table width="100%" border="0" cellpadding="0" cellspacing="0" class="font">
                <tr>
                    <td align="center" style="font-size: 18px">投保資料</td>
                </tr>
                <tr>
                    <td>日期區間:<asp:Label ID="labPDataArea" runat="server"></asp:Label>, 狀態:<asp:Label ID="labPType" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:DataGrid ID="DataGrid1" BorderColor="black" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                            AllowPaging="False">
                            <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn HeaderText="班級">
                                    <ItemTemplate>
                                        <asp:Label ID="labClassName" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="姓名" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labStdName" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="開訓日期" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labSTDate" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="結訓日期" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labETDate" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="身分證字號" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labIDNO" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="生日" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labBirth" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="預算別" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labBud" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="異動別" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labType" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="加保日" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labAppDate" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="退保日" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labOutDate" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:DataGrid>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
