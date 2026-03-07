<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_002_01.aspx.vb" Inherits="WDAIIP.CP_04_002_01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_002_01</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <p>
            <table id="Table7" cellspacing="0" cellpadding="0" width="100%" border="0">
                <tr>
                    <td style="width: 20%">
                        <font face="新細明體">
                            <asp:Label ID="Label2" runat="server" CssClass="font">訓練年度:</asp:Label><asp:Label
                                ID="train_year" runat="server" CssClass="font"></asp:Label></font>
                    </td>
                    <td style="width: 40%">
                        <asp:Label ID="Label9" runat="server" CssClass="font">訓練計畫:</asp:Label><asp:Label
                            ID="train_plan" runat="server" CssClass="font"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <font face="新細明體">
                            <asp:Label ID="Label10" runat="server" CssClass="font">訓練職類:</asp:Label><asp:Label
                                ID="train_name" runat="server" CssClass="font"></asp:Label></font>
                    </td>
                </tr>
            </table>
        </p>
        <font face="新細明體"></font>
        <asp:Label ID="Label1" runat="server" CssClass="font">目標</asp:Label><br>
        <table class="table_sch" cellpadding="1" cellspacing="1">
            <tr class="font">
                <td class="bluecol">
                    <div>
                        緣由
                    </div>
                </td>
                <td>
                    <asp:TextBox ID="PlanCause" runat="server" TextMode="MultiLine" MaxLength="100" Width="100%"></asp:TextBox>
                </td>
            </tr>
            <tr class="font">
                <td class="bluecol">
                    <div>
                        學科
                    </div>
                </td>
                <td width="90%">
                    <asp:TextBox ID="PurScience" runat="server" TextMode="MultiLine" MaxLength="100"
                        Width="100%"></asp:TextBox>
                </td>
            </tr>
            <tr class="font">
                <td class="bluecol">
                    <div>
                        技能
                    </div>
                </td>
                <td>
                    <asp:TextBox ID="PurTech" runat="server" TextMode="MultiLine" MaxLength="100" Width="100%"></asp:TextBox>
                </td>
            </tr>
            <tr class="font">
                <td class="bluecol">
                    <div>
                        品德
                    </div>
                </td>
                <td>
                    <asp:TextBox ID="PurMoral" runat="server" TextMode="MultiLine" MaxLength="100" Width="100%"></asp:TextBox>
                </td>
            </tr>
        </table>
        <p>
            <asp:Label ID="Label3" runat="server" CssClass="font">受訓資格</asp:Label>
            <table id="Table1" class="table_sch" cellpadding="1" cellspacing="1">
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            學歷
                        </div>
                    </td>
                    <td style="height: 22px" width="93%" class="whitecol">
                        <asp:Label ID="CapDegree" runat="server"></asp:Label>&nbsp;(含以上)
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            年齡
                        </div>
                    </td>
                    <td style="height: 25px" class="whitecol">
                        <asp:Label ID="CapAge1" runat="server"></asp:Label>&nbsp;~
                    <asp:Label ID="CapAge2" runat="server"></asp:Label>歲
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            性別
                        </div>
                    </td>
                    <td style="height: 17px" class="whitecol">
                        <asp:Label ID="CapSex" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            兵役
                        </div>
                    </td>
                    <td style="height: 39px" class="whitecol">
                        <asp:Label ID="CapMilitary" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            其他ㄧ
                        </div>
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="CapOther1" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            其他二
                        </div>
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="CapOther2" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            其他三
                        </div>
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="CapOther3" runat="server"></asp:Label>
                    </td>
                </tr>
            </table>
        </p>
        <asp:Label ID="Label12" runat="server" CssClass="font">訓練方式</asp:Label>
        <table id="Table2" class="table_sch" cellpadding="1" cellspacing="1">
            <tr class="font">
                <td class="bluecol">
                    <div>
                        學科
                    </div>
                </td>
                <td width="93%" class="whitecol">
                    <asp:Label ID="TMScience" runat="server" CssClass="font"></asp:Label>
                </td>
            </tr>
            <tr class="font">
                <td class="bluecol">
                    <div>
                        術科
                    </div>
                </td>
                <td class="whitecol">
                    <asp:Label ID="TMTech" runat="server" CssClass="font"></asp:Label>
                </td>
            </tr>
        </table>
        <p>
        </p>
        <font face="新細明體">
            <asp:Label ID="Label50" runat="server" CssClass="font">課程編配</asp:Label>
            <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1">
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            學科
                        </div>
                    </td>
                    <td class="whitecol" style="height: 43px" width="19%">
                        <asp:Label ID="SciHours" runat="server" CssClass="font"></asp:Label>小時
                    </td>
                    <td class="bluecol">
                        <div>
                            1. 一般學科
                        </div>
                    </td>
                    <td class="whitecol" style="height: 43px" width="64%">
                        <asp:Label ID="GenSciHours" runat="server" CssClass="font"></asp:Label>小時
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol"></td>
                    <td height="21" class="whitecol"></td>
                    <td class="bluecol">
                        <div>
                            2. 專業學科
                        </div>
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="ProSciHours" runat="server" CssClass="font"></asp:Label>小時
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            術科
                        </div>
                    </td>
                    <td class="whitecol" style="height: 24px" colspan="3">
                        <asp:Label ID="ProTechHours" runat="server" CssClass="font"></asp:Label>小時
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            其他時數
                        </div>
                    </td>
                    <td class="whitecol" style="height: 25px" colspan="3">
                        <asp:Label ID="OtherHours" runat="server" CssClass="font"></asp:Label>小時
                    </td>
                </tr>
                <tr class="font">
                    <td class="bluecol">
                        <div>
                            總計
                        </div>
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="TotalHours" runat="server" CssClass="font"></asp:Label>小時
                    </td>
                </tr>
            </table>
            <p>
                <asp:Label ID="Label20" runat="server" CssClass="font">班別資料</asp:Label>
                <table class="table_sch" id="Table4" cellpadding="1" cellspacing="1">
                    <tr class="font">
                        <td style="height: 17px" width="13%" class="bluecol">
                            <div>
                                班別名稱
                            </div>
                        </td>
                        <td style="height: 17px" colspan="3" class="whitecol">
                            <asp:Label ID="ClassName" runat="server" CssClass="font"></asp:Label>
                        </td>
                    </tr>
                    <tr class="font">
                        <td style="height: 35px" class="bluecol">
                            <div>
                                訓練人數
                            </div>
                        </td>
                        <td style="height: 35px" class="whitecol">
                            <asp:Label ID="TNum" runat="server" CssClass="font"></asp:Label>人
                        </td>
                        <td style="height: 35px" class="bluecol">
                            <div>
                                訓練時數
                            </div>
                        </td>
                        <td style="height: 35px" width="51%" class="whitecol">
                            <asp:Label ID="THours" runat="server" CssClass="font"></asp:Label>小時
                        </td>
                    </tr>
                    <tr class="font">
                        <td class="bluecol">
                            <div>
                                訓練起日
                            </div>
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="STDate" runat="server" CssClass="font"></asp:Label>
                        </td>
                        <td class="bluecol">
                            <div>
                                訓練迄日
                            </div>
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="FDDate" runat="server" CssClass="font"></asp:Label>
                        </td>
                    </tr>
                    <tr class="font">
                        <td style="height: 27px" class="bluecol">
                            <div>
                                期別(二碼)
                            </div>
                        </td>
                        <td style="height: 27px" class="whitecol">
                            <asp:Label ID="CyclType" runat="server" CssClass="font"></asp:Label>
                        </td>
                        <td class="bluecol">
                            <div>
                                班數
                            </div>
                        </td>
                        <td style="height: 27px" class="whitecol">
                            <asp:Label ID="ClassCount" runat="server" CssClass="font">1</asp:Label>
                        </td>
                    </tr>
                </table>
                <p>
                </p>
            <asp:Label ID="Label11" runat="server" CssClass="font">訓練費用</asp:Label>&nbsp;
            <asp:Label ID="TrainCostStatus" runat="server" CssClass="font"></asp:Label><asp:Panel
                ID="Panel1" runat="server">
                <table class="font" id="DataGrid1Table" cellspacing="1" cellpadding="1" width="100%"
                    border="1" runat="server">
                    <tr>
                        <td style="height: 17px" class="bluecol_sub_left">費用列表
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" BorderColor="Gray"
                                AutoGenerateColumns="False">
                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="CostName" HeaderText="項目" FooterText="總計"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="OPrice" HeaderText="單價"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="Itemage" HeaderText="數量"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="ItemCost" HeaderText="計價單位"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="AllSmallSum" HeaderText="小計">
                                        <HeaderStyle Width="50px"></HeaderStyle>
                                    </asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                    <tr id="AdmGrantTR" runat="server">
                        <td>
                            <table class="table_sch" id="AdmTable" runat="server" cellpadding="1" cellspacing="1">
                                <tr>
                                    <td width="80" class="bluecol">行政管理費
                                    </td>
                                    <td class="whitecol">
                                        <asp:Label ID="AdmCost" runat="server"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                            <table class="table_sch" id="Table10" cellpadding="1" cellspacing="1">
                                <tr>
                                    <td width="80" class="bluecol">總計
                                    </td>
                                    <td class="whitecol">
                                        <asp:Label ID="TotalCost1" runat="server"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="Panel2" runat="server" Height="244px">
                <table class="font" id="TableCost2" cellspacing="1" cellpadding="1" width="600" border="1"
                    runat="server">
                    <tr>
                        <td>
                            <table class="font" id="Table11" cellspacing="1" cellpadding="1" border="1">
                                <tr>
                                    <td>每人每時單價計價
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table id="DataGrid2Table" cellspacing="1" cellpadding="1" width="100%" border="1"
                                runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" BorderColor="Gray" AutoGenerateColumns="False">
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                            <Columns>
                                                <asp:BoundColumn DataField="OPrice" HeaderText="單價"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Itemage" HeaderText="人數"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="ItemCost" HeaderText="時數"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="AllSmallSum" HeaderText="小計">
                                                    <HeaderStyle Width="80px"></HeaderStyle>
                                                </asp:BoundColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table class="table_sch" id="Table12" cellpadding="1" cellspacing="1">
                                            <!-- add by nick 060519 加入行政管理費 -->
                                            <tr>
                                                <td width="80" class="bluecol">行政管理費
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="AdmCost2" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td width="80" class="bluecol">總計
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="TotalCost2" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="Panel3" runat="server">
                <table class="font" id="TableCost3" cellspacing="1" cellpadding="1" width="600" border="1"
                    runat="server">
                    <tr>
                        <td>
                            <table class="font" id="Table13" cellspacing="1" cellpadding="1" border="1">
                                <tr>
                                    <td>每人輔助單價計費
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table id="DataGrid3Table" cellspacing="1" cellpadding="1" width="100%" border="1"
                                runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" BorderColor="Gray" AutoGenerateColumns="False">
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                            <Columns>
                                                <asp:BoundColumn DataField="OPrice" HeaderText="單價"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Itemage" HeaderText="人數"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="AllSmallSum" HeaderText="小計">
                                                    <HeaderStyle Width="80px"></HeaderStyle>
                                                </asp:BoundColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table class="table_sch" id="Table8" cellpadding="1" cellspacing="1">
                                            <!-- add by nick 060519 加入行政管理費 -->
                                            <tr>
                                                <td width="80" class="bluecol">行政管理費
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="AdmCost3" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td width="80" class="bluecol">總計
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="TotalCost3" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="Panel4" runat="server">
                <table class="font" id="TableCost4" cellspacing="1" cellpadding="1" width="600" border="1"
                    runat="server">
                    <tr>
                        <td>
                            <table id="DataGrid4Table" cellspacing="1" cellpadding="1" width="100%" border="1"
                                runat="server">
                                <tr>
                                    <td style="height: 135px">
                                        <asp:DataGrid ID="DataGrid4" runat="server" CssClass="font" Width="100%" BorderColor="Gray"
                                            AutoGenerateColumns="False">
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                            <Columns>
                                                <asp:BoundColumn DataField="CostName" HeaderText="項目"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="OPrice" HeaderText="單價"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Itemage" HeaderText="數量"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="AllSmallSum" HeaderText="小計">
                                                    <HeaderStyle Width="80px"></HeaderStyle>
                                                </asp:BoundColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table class="table_sch" id="Table9" cellpadding="1" cellspacing="1">
                                            <!-- add by nick 060519 加入行政管理費 -->
                                            <tr>
                                                <td width="80" class="bluecol">行政管理費
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="AdmCost4" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td width="80" class="bluecol">總計
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="TotalCost4" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td width="80" class="bluecol">費用/人
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="PerCost" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <p>
                <asp:Label ID="Label7" runat="server" CssClass="font">經費來源</asp:Label>
                <table class="table_sch" id="Table5" cellpadding="1" cellspacing="1">
                    <tr class="font">
                        <td class="bluecol" width="9%">
                            <div>
                                經費來源
                            </div>
                        </td>
                        <td width="91%" class="whitecol">
                            <p>
                                &nbsp;&nbsp;&nbsp;<span class="font">政府負擔新臺幣&nbsp;
                                    <asp:Label ID="MainCost" runat="server" CssClass="font"></asp:Label>元 </span>
                            </p>
                            <p>
                                &nbsp;&nbsp;&nbsp;<span class="font">企業負擔新臺幣&nbsp;
                                    <asp:Label ID="CenterCost" runat="server" CssClass="font"></asp:Label>元</span>
                            </p>
                            <p>
                                &nbsp;&nbsp;&nbsp;<span class="font">學員負擔新臺幣&nbsp;
                                    <asp:Label ID="UnitCost" runat="server" CssClass="font"></asp:Label>元</span>
                            </p>
                        </td>
                    </tr>
                </table>
            </p>
        </font>
        <asp:Label ID="Label8" runat="server" CssClass="font">備註</asp:Label>
        <table id="Table6" class="table_sch" cellpadding="1" cellspacing="1">
            <tr class="font">
                <td width="9%" class="bluecol">
                    <div>
                        備註
                    </div>
                </td>
                <td width="83%" class="whitecol">
                    <textarea id="Note" disabled name="UNIT_NOTE" rows="9" cols="50" runat="server"></textarea>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
