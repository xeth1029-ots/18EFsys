<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_003_add.aspx.vb" Inherits="WDAIIP.CP_04_003_add" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <style type="text/css">
        .class_link A { color: #000000; }
            .class_link A:link { color: #0000ff; }
            .class_link A:hover { color: #0000ff; }
        A:visited { color: #0000ff; }
        A:active { color: #0000ff; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" width="100%">
            
            <tr>
                <td>
                    <div>
                        <table class="table_sch" id="Table6" cellspacing="0" cellpadding="2" width="100%" border="0">
                            <tr>
                                <td width="10%" class="bluecol">
                                    <asp:Label ID="Year" runat="server" ForeColor="Red" CssClass="font">年度：</asp:Label>
                                    <asp:Label ID="YearLabel" runat="server" CssClass="font"></asp:Label>
                                </td>
                                <td width="30%" class="bluecol">
                                    <font size="2">
                                        <asp:Label ID="District" runat="server" ForeColor="Red" CssClass="font">轄區：</asp:Label>
                                        <asp:Label ID="DistrictLabel" runat="server" CssClass="font"></asp:Label>
                                    </font>
                                </td>
                                <td width="10%" class="bluecol">
                                    <font size="2">
                                        <asp:Label ID="Count" runat="server" ForeColor="Red" CssClass="font">總班數：</asp:Label>
                                        <asp:Label ID="CountLabel" runat="server"></asp:Label>
                                    </font>
                                </td>
                                <td width="12%" class="bluecol">
                                    <asp:Label ID="Label1" runat="server" ForeColor="Red" CssClass="font">招生總人數：</asp:Label>
                                    <asp:Label ID="STNum" runat="server"></asp:Label>
                                </td>
                                <td width="12%" class="bluecol">
                                    <asp:Label ID="Label3" runat="server" ForeColor="Red" CssClass="font">開訓總人數：</asp:Label>
                                    <asp:Label ID="SSNum" runat="server"></asp:Label>
                                </td>
                                <td width="12%" class="bluecol">
                                    <asp:Label ID="Label5" runat="server" ForeColor="Red" CssClass="font">結訓總人數：</asp:Label>
                                    <asp:Label ID="SESNum" runat="server"></asp:Label>
                                </td>
                                <%--<td width="12%" class="bluecol" style="display: none;">
                                    <asp:Label ID="Label2" runat="server" ForeColor="Red" CssClass="font">就業總人數：</asp:Label>
                                    <asp:Label ID="SGSNum" runat="server"></asp:Label>
                                </td>--%>
                            </tr>
                        </table>
                    </div>
                    <div>
                        <table class="font" id="Table5" cellspacing="0" cellpadding="0" border="0" width="100%">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" PageSize="20" AllowSorting="True" Width="100%" AutoGenerateColumns="False" AllowPaging="True" DataKeyField="OCID">
                                        <AlternatingItemStyle></AlternatingItemStyle>
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle Width="3%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="DistName" SortExpression="DistID" HeaderText="轄區">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="PlanName" SortExpression="PlanID" HeaderText="訓練計畫">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構名稱">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="CityName" SortExpression="CityName" HeaderText="縣市別">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:ButtonColumn DataTextField="ClassCName2" SortExpression="ClassCName2" HeaderText="班別名稱">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                                <ItemStyle></ItemStyle>
                                            </asp:ButtonColumn>
                                            <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                                <HeaderStyle Width="4%" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="TPropertyIDN" HeaderText="訓練性質">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="HourRanName" HeaderText="訓練時段">
                                                <HeaderStyle Width="4%" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="開訓日期">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="CNum" HeaderText="招生人數">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="THours" HeaderText="時數">
                                                <HeaderStyle Width="4%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="SNum" HeaderText="開訓人數">
                                                <HeaderStyle Width="3%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ESNum" HeaderText="結訓人數">
                                                <HeaderStyle Width="3%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <%--  <asp:TemplateColumn Visible="False" HeaderText="功能">
                                                <HeaderStyle Width="3%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Button runat="server" Text="詳細" CausesValidation="false" ID="Button1"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>--%>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <div align="center">
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Label ID="NoData" runat="server" CssClass="font"></asp:Label></td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btnPrint" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="Button2" runat="server" Text="回上頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                            <tr>
                                <td align="left">
                                    <asp:Label ID="description" runat="server" CssClass="font">排序說明：以轄區、訓練計畫、開訓日期做排序</asp:Label></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
