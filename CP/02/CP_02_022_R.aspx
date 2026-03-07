<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_02_022_R.aspx.vb" Inherits="WDAIIP.CP_02_022_R" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員身份及自負費用統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table style="width: 760px" cellspacing="0" cellpadding="0" width="760" border="0">
        <tr>
            <td class="font">
                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                <asp:Label ID="TitleLab2" runat="server">
						首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;<FONT color="#990000">中長期失業者職前訓練計畫-學員身份及自負費用統計表</FONT>
                </asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                <table class="font" style="width: 760px; height: 128px" cellspacing="1" cellpadding="1" width="760" border="0">
                    <tr>
                        <td class="font" align="left" width="100" bgcolor="#cc6666">
                            <font face="新細明體"><font color="#ffffff">&nbsp;&nbsp;&nbsp;&nbsp;年度</font></font>
                        </td>
                        <td style="width: 276px">
                            <font face="新細明體">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </font>
                        </td>
                        <td align="left" width="100" bgcolor="#cc6666">
                            <font color="#ffffff">&nbsp;&nbsp;&nbsp; 訓練職類</font>
                        </td>
                        <td>
                            <font face="新細明體">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="152px"></asp:TextBox><input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server">&nbsp;<input id="trainValue" type="hidden" name="trainValue" runat="server">
                            </font>
                        </td>
                    </tr>
                    <tr>
                        <td class="font" align="left" width="100" bgcolor="#cc6666">
                            <div class="font" align="left">
                                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 訓練機構</font></div>
                        </td>
                        <td colspan="3">
                            <font face="新細明體">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox><input id="Org" type="button" value="..." name="Org" runat="server"><input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server"><input id="TPlanid" style="width: 32px; height: 22px" type="hidden" name="TPlanid" runat="server"><input id="Re_ID" style="width: 32px; height: 22px" type="hidden" name="Re_ID" runat="server"><br>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span></font>
                        </td>
                    </tr>
                    <tr>
                        <td class="font" align="left" width="100" bgcolor="#cc6666">
                            <font color="#ffffff">&nbsp;&nbsp;&nbsp; 班級名稱</font>
                        </td>
                        <td style="width: 276px">
                            <asp:TextBox ID="ClassName" runat="server"></asp:TextBox>
                        </td>
                        <td bgcolor="#cc6666">
                            <font color="#ffffff">&nbsp;&nbsp;&nbsp; 期別</font>
                        </td>
                        <td>
                            <asp:TextBox ID="CyclType" runat="server" Columns="5"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4">
                            <div align="center">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox><asp:Button ID="btnQuery" runat="server" Text="查詢"></asp:Button>&nbsp;
                            </div>
                        </td>
                    </tr>
                </table>
                <table id="Table1" cellspacing="1" cellpadding="1" width="760" border="0">
                    <tr>
                        <td align="center">
                            <p>
                                <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="757" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <font face="新細明體">訓練計畫：
                                                <asp:Label ID="TPlanName" runat="server" CssClass="font"></asp:Label></font>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" OnItemCommand="dtPlan_ItemCommand" AllowPaging="True" PagerStyle-Mode="NumericPages" PagerStyle-HorizontalAlign="Left" AllowSorting="True">
                                                <ItemStyle HorizontalAlign="Center" BackColor="#FFF9E1"></ItemStyle>
                                                <HeaderStyle HorizontalAlign="Center" BackColor="#FFCCCC"></HeaderStyle>
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="PlanYear" HeaderText="計畫年度"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="AppliedDate" SortExpression="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}">
                                                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn Visible="False" HeaderText="訓練職類"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="訓練起日" DataFormatString="{0:d}">
                                                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="FDDate" SortExpression="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}">
                                                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="機構名稱">
                                                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="ClassName" SortExpression="ClassName" HeaderText="班名">
                                                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Left" Width="130px"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <ItemTemplate>
                                                            <asp:Button ID="Button1" runat="server" Text="列印" CommandName="btnEdit"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                                                </Columns>
                                                <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                        </td>
                                    </tr>
                                </table>
                            </p>
                            <p>
                                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
