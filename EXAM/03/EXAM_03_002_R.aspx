<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EXAM_03_002_R.aspx.vb" Inherits="WDAIIP.EXAM_03_002_R" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>甄試考卷列印</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
    <script language="javascript">
        function choose_class() {
            openClass('../../SD/02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;甄試管理&gt;&gt;甄試班級考題列印</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="tab_title" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"><FONT face="新細明體">首頁&gt;&gt;招生甄試設定管理&gt;&gt;甄試班級考題列印&gt;&gt;</FONT></asp:Label><asp:Label ID="TitleLab2" runat="server"><font color="#990000">列印試卷</font></asp:Label>
						</td>
					</tr>
				</table>--%>
                    <asp:Panel ID="tab_sch" runat="server" Visible="True">
                        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">訓練機構
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="Button8" value="..." type="button" name="Button8" runat="server">
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                    <span style="position: absolute; display: none" id="HistoryList2" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%">
                                        </asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">職類/班級
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input id="Button5" value="..." type="button" name="Button5" runat="server">
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                    <span style="position: absolute; display: none; left: 270px" id="HistoryList">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%">
                                        </asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">甄試日期
                                </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txt_examdate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txt_examdate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btn_sch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><br>
                                    <asp:Label ID="msg" runat="server" Visible="False" ForeColor="Red">查無資料!!</asp:Label>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="tab_view" runat="server" Visible="False">
                        <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td align="center">
                                    <asp:DataGrid ID="dg_Sch" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="ocid" HeaderText="ocid"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="classcname" HeaderText="班級名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="41%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="isonline" HeaderText="試卷類型">
                                                <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="examdate" HeaderText="甄試日期">
                                                <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="28%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="btn_prt_que" runat="server" Text="列印試卷" CommandName="que" CssClass="asp_Export_M"></asp:Button>
                                                    <asp:Button ID="btn_prt_ans" runat="server" Text="列印解答" CommandName="ans" CssClass="asp_Export_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <asp:Panel runat="server" Visible="false" ID="rpt_view">
            <table border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td>
                        <div id="div1" align="center" runat="server">
                            <div align="center">
                                <asp:Label ID="lbl_distname" runat="server" Font-Size="Large" Font-Bold="True"></asp:Label>
                            </div>
                            <div align="center">
                                <asp:Label ID="lbl_classname" runat="server" Font-Size="Large" Font-Bold="True"></asp:Label>
                            </div>
                            <div align="center">
                                <asp:Label ID="lbl_title" runat="server" Font-Size="Large" Font-Bold="True">學科甄試題目</asp:Label>
                            </div>
                            <div align="center">
                                <asp:Label ID="Label1" runat="server" Font-Size="Large" Font-Bold="True">&nbsp;</asp:Label>
                            </div>
                            <div align="left">
                                <asp:Label ID="lbl_examtime" runat="server" Font-Bold="True"></asp:Label>
                            </div>
                            <asp:DataGrid ID="dg_view" runat="server" Width="100%" AutoGenerateColumns="False" ShowHeader="False" GridLines="None">
                                <Columns>
                                    <asp:BoundColumn DataField="title">
                                        <ItemStyle HorizontalAlign="Right" Width="10%" VerticalAlign="Top"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="data">
                                        <ItemStyle HorizontalAlign="Left" Width="90%" VerticalAlign="Top"></ItemStyle>
                                    </asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div align="center" class="whitecol">
                            <asp:Button ID="btn_lev" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            &nbsp;<asp:Button ID="btn_prt" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
