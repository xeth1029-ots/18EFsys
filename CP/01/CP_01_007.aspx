<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_01_007.aspx.vb" Inherits="WDAIIP.CP_01_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>不預告(電話)抽訪學員紀錄表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        <%--
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);--%>

        function CheckAdd() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請選擇職類班別!')
                return false;
            }
        }

        /*function search() {var msg = '';if (document.form1.start_date.value == '') msg += '請選擇起始日期\n';if (document.form1.end_date.value == '') msg += '請選擇終至日期\n';if (msg!=''){alert(msg);return false;}}*/
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;不預告(電話)抽訪學員紀錄表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td align="center">
                        <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol" width="20%">機構</td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input type="button" value="..." id="Button5" name="Button5" runat="server">
                                    <input id="RIDValue" type="hidden" name="Hidden1" runat="server" size="1"><br>
                                    <span id="HistoryList2" style="position: absolute; display: none">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">職類/班別</td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input onclick="javascript: openClass('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value);" type="button" value="..." id="Button6" name="Button6" runat="server">
                                    <input id="TMIDValue1" type="hidden" name="Hidden2" runat="server" size="1">
                                    <input id="OCIDValue1" type="hidden" name="Hidden1" runat="server" size="1"><br>
                                    <span id="HistoryList" style="position: absolute; display: none; left: 30%">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">抽訪日期</td>
                                <td class="whitecol" width="80%">
                                    <span id="span01" runat="server">
                                        <asp:TextBox ID="start_date" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server">
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
                                        <asp:TextBox ID="end_date" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server">
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">抽訪狀態</td>
                                <td class="whitecol" width="80%">
                                    <asp:RadioButtonList ID="VisitItem" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1" Selected="True">未抽訪</asp:ListItem>
                                        <asp:ListItem Value="2">已抽訪</asp:ListItem>
                                        <asp:ListItem Value="3">全部</asp:ListItem>
                                    </asp:RadioButtonList>
                                    &nbsp; (選取未抽訪時，抽訪日期選項不列入查詢項目) </td>
                            </tr>
                            <tr>
                                <td colspan="2" class="whitecol" width="100%" align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>&nbsp;
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="PrintBlank" runat="server" Text="列印空白表單" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </td>
                </tr>
            </tbody>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <ItemStyle HorizontalAlign="Center" Width="6%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ApplyDate" HeaderText="抽訪日期">
                                <ItemStyle HorizontalAlign="Center" Width="12%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="抽訪結果">
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="8%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="原因及後續追蹤">
                                <ItemStyle HorizontalAlign="Center" Width="16%" VerticalAlign="Middle"></ItemStyle>
                                <ItemTemplate>
                                    <textarea id="Reason" style="width: 90%; height: 90px;" name="VerReason" rows="3" cols="15" runat="server"></textarea>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <ItemStyle HorizontalAlign="Center" Width="16%"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="Button3" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button4" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button8" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="Button7" runat="server" Text="新增" CommandName="Add" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="PrintBlank2" runat="server" Text="列印空白表單" CssClass="asp_Export_M"></asp:Button>
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
        <input id="Years" type="hidden" name="Years" runat="server" />
    </form>
</body>
</html>
