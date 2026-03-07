<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_002_add_01.aspx.vb" Inherits="WDAIIP.CP_04_002_add_01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_002_add_01</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
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
                    <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;</font><font class="font" color="#800000" size="2">計畫資料</font>
                </td>
            </tr>
        </table>

        <table class="font" id="Table5" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <div align="center">
                        <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" AllowPaging="True" Width="100%" CssClass="font">
                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                                <asp:BoundColumn DataField="PlanName" HeaderText="計畫名稱"></asp:BoundColumn>
                                <asp:BoundColumn DataField="Class1" HeaderText="已開班">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Class2" HeaderText="未開班">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Class3" HeaderText="不開班">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Class4" HeaderText="總班數">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="student1" HeaderText="計畫總人數">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="student2" HeaderText="在訓總人數">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="student3" HeaderText="結訓總人數">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </div>
                </td>
            </tr>
            <tr>
                <td style="height: 16px" align="center">
                    <asp:Label ID="NoData" runat="server" CssClass="font"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="Button3" runat="server" Text="匯出" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="Button2" runat="server" Text="回上頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td align="left">
                    <p>
                        &nbsp;
                    </p>
                    <p>
                        <asp:Label ID="description" runat="server" Width="368px" CssClass="font">統計說明：依訓練計畫統計 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							訓練總人數=>目前總參訓人數 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							在訓總人數=>目前學員狀態為在訓的總人數&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
							結訓總人數=>目前學員狀態為結訓的總人數</asp:Label>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
