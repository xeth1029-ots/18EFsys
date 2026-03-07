<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_01_001.aspx.vb" Inherits="WDAIIP.CP_01_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>實地訪查紀錄表</title>
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
    </script>
    <script type="text/javascript" language="javascript">
        function search() {
            var msg = '';
            var start_date = document.getElementById("start_date");
            var end_date = document.getElementById("end_date");
            if (start_date.value == '') msg += '請選擇起始日期\n';
            if (end_date.value == '') msg += '請選擇終至日期\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;查核績效管理&gt;&gt;實地訪查紀錄表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button5" type="button" value="..." name="Button5" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="Hidden1" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button6" onclick="javascript: openClass('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value);" type="button" value="..." name="Button6" runat="server" class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="Hidden2" runat="server">
                                <input id="OCIDValue1" type="hidden" name="Hidden1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訪查日期</td>
                            <td class="whitecol">
                                <span id="span01" runat="server">
                                    <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    ～
                                    <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">抽訪情況</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="interview" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">全部</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">己抽訪</asp:ListItem>
                                    <asp:ListItem Value="3">未抽訪 (訪查日期改為開訓日期)</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯入訪查紀錄表</td>
                            <td class="whitecol">
                                <input id="File1" type="file" name="File1" runat="server" size="66" accept=".xls,.ods" />
                                <asp:Button ID="Btn_XlsImport" runat="server" Text="匯入名冊" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)<br>
                                <asp:HyperLink ID="Hyperlink1" runat="server" CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>

                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
								    <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>&nbsp;
								    <asp:Button ID="Button10" runat="server" Text="列印空白實地訪查紀錄表" CssClass="asp_Export_M"></asp:Button>&nbsp;
								    <asp:Button ID="Button7" runat="server" Text="回查核頁" CssClass="asp_button_M"></asp:Button>&nbsp;
								    <asp:Button ID="btnExport1" runat="server" Text="匯出本署訪查資料" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                    <%--<asp:Button ID="btnExport2" runat="server" Text="匯出縣市政府訪查資料" CssClass="asp_Export_M" Visible="false"></asp:Button>--%>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <div id="Div1" runat="server">
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                                <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="CYCLTYPE" HeaderText="期別">
                                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="COrgName" HeaderText="訪視單位">
                                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDATE" HeaderText="開訓日期">
                                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDATE" HeaderText="結訓日期">
                                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ApplyDate" HeaderText="訪查日期">
                                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="VISITORNAME" HeaderText="訪查人員">
                                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="訪查結果">
                                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="LabITEM32" runat="server"></asp:Label></ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="綜合建議">
                                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lItem31Note" runat="server"></asp:Label></ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="Button3" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="Button4" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="Button8" runat="server" Text="列印" CommandName="prt" CssClass="asp_Export_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <span style="color: #FF0000">
                        <asp:Label ID="Labmsg2" runat="server" Text=""></asp:Label>
                    </span>
                </td>
            </tr>
        </table>
        <input id="HidThours" type="hidden" name="HidThours" runat="server">
        <input id="HIDVISCOUNT" type="hidden" name="HIDVISCOUNT" runat="server">
    </form>
</body>
</html>
