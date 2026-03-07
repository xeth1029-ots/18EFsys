<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_026.aspx.vb" Inherits="WDAIIP.SD_05_026" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>學員報名查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATO" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function chkdata() {
            //alert('111');
            var msg = '';
            //if(document.form1.IDNO.value=='') msg+='請輸入身分證號碼\n';
            if (document.form1.IDNO.value == '' && document.form1.Name.value == "") {
                msg += '請輸入身分證號碼或姓名\n';
            }
            //else if(checkId(document.form1.IDNO.value)) msg+='身分證號碼錯誤!\n';
            if (msg != '') {
                alert(msg)
                return false;
            }
        }

        function GetMode() {
            document.form1.center.value = '';
            document.form1.RIDValue.value = '';
            document.form1.OCIDValue.value = '';
            document.form1.PlanID.value = '';
            for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
                document.form1.OCID.options[i] = null;
            }
            document.form1.OCID.options[0] = new Option('請選擇機構');
            if (document.form1.DistID.selectedIndex != 0 && document.form1.TPlanID.selectedIndex != 0) {
                document.form1.Button3.disabled = false;
            }
            else {
                document.form1.Button3.disabled = true;
            }
        }

        function ShowPersonData(obj) {
            document.getElementById(obj).style.display = 'inline';
        }

        function HidPersonData(obj) {
            document.getElementById(obj).style.display = 'none';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學員報名查詢</asp:Label>--%>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;學員報名查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Panel ID="Searchtable" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">身分證號碼</td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="IDNO" runat="server" Width="50%"></asp:TextBox></td>
                                <td class="bluecol" style="width: 20%">姓名</td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="Name" runat="server" Width="50%" placeholder="請輸入全名"></asp:TextBox></td>
                            </tr>
                            <tr id="tr01e" runat="server">
                                <td class="bluecol">報名區間</td>
                                <td class="whitecol" colspan="3">
                                    <font color="#ffffff">
                                        <asp:TextBox ID="RelEnterDate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('RelEnterDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～</font>
                                        <asp:TextBox ID="RelEnterDate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('RelEnterDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    </font>
                                    &nbsp;&nbsp;<asp:Label ID="Note" runat="server" ForeColor="Red">搜尋條件【身分證號碼】與【姓名】，請擇一輸入</asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">已結訓班級</td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="rblFTDate" Style="z-index: 0" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="Y">是</asp:ListItem>
                                        <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                                        <asp:ListItem Value="A">全部</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" Font-Size="9pt">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <table class="font" id="ShowDatatable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <%--<tr>
                <td align="right">
                    <table class="font" id="table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td align="right">(為避免消耗主機效能，最大搜尋筆數為1000筆) </td>
                        </tr>
                    </table>
                </td>
            </tr>--%>
            <tr>
                <td align="center" class="whitecol">
                    <input id="Button5" type="button" value="回上一頁" name="Button4" runat="server" class="asp_button_M" /></td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AllowSorting="true" AllowPaging="true" PageSize="20" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" SortExpression="IDNO" HeaderText="身分證號碼">
                                <HeaderStyle></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Birthday" SortExpression="Birthday" HeaderText="出生日期" DataFormatString="{0:d}">
                                <HeaderStyle></HeaderStyle>
                            </asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="DistName" SortExpression="DistName" HeaderText="轄區&lt;BR&gt;中心">--%>
                            <asp:BoundColumn DataField="DistName" SortExpression="DistName" HeaderText="轄區&lt;BR&gt;分署">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Years" HeaderText="年度">
                                <HeaderStyle></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                <HeaderStyle></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
                                <HeaderStyle></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="訓練職類">
                                <ItemTemplate>
                                    <asp:Label ID="LTMID" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="classcname" SortExpression="ClassName" HeaderText="班別名稱">
                                <HeaderStyle></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SFdate" SortExpression="tround" HeaderText="開/結訓期間">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Eterdate" SortExpression="tround" HeaderText="報名日期">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="報名管道">
                                <ItemTemplate>
                                    <asp:Label ID="LEnterChannel" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="報名狀態">
                                <ItemTemplate>
                                    <asp:Label ID="LEnterType" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="錄取狀態">
                                <ItemTemplate>
                                    <asp:Label ID="LAdmission" runat="server"></asp:Label>
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
            <tr>
                <td align="center" class="whitecol">
                    <input id="Button4" type="button" value="回上一頁" name="Button4" runat="server" class="asp_button_M" /></td>
            </tr>
        </table>
    </form>
</body>
</html>
