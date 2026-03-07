<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_004.aspx.vb" Inherits="WDAIIP.SD_02_004" %>


<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>甄試通知單設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
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
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試通知單設定</asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <table id="table3" class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">年度</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:DropDownList ID="ddlyears" runat="server" AutoPostBack="true">
                                </asp:DropDownList>
                                <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server" />
                            </td>
                            <td class="bluecol" style="width: 20%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server" class="asp_button_Mini" />
                                <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:HiddenField ID="hidDistID" runat="server" />
                                <asp:HiddenField ID="hidPlanID" runat="server" />
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Btn_OrgSet" runat="server" Width="140px" Text="甄試通知單設定" CssClass="asp_button_L"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                                <input id="TPlanid" style="width: 32px; height: 22px" type="hidden" name="TPlanid" runat="server" />
                                <input id="Re_ID" style="width: 32px; height: 22px" type="hidden" name="Re_ID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">班級名稱</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="ClassName" runat="server" Width="40%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 20%">期別</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="30%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr style="display: none">
                            <td class="bluecol" id="dt_datatype" runat="server">資料類型</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="IsApprPaper" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" Visible="False">
                                    <asp:ListItem Value="Y" Selected="true">正式</asp:ListItem>
                                    <asp:ListItem Value="N">草稿</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_audit1" style="display: none" runat="server">
                            <td class="bluecol" id="td_audit" runat="server">審核狀態
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="audit" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" Visible="False">
                                    <asp:ListItem Value="A" Selected="true">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y">己審核</asp:ListItem>
                                    <asp:ListItem Value="N">審核中</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>訓練計畫：<asp:Label ID="TPlanName" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" OnItemCommand="dtPlan_ItemCommand" AllowPaging="true" PagerStyle-Mode="NumericPages" PagerStyle-HorizontalAlign="Left" AllowSorting="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="5%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="計畫年度">
                                            <HeaderStyle Width="5%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" SortExpression="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#B0E2FF" Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="訓練職類">
                                            <HeaderStyle Width="5%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Stdate" SortExpression="Stdate" HeaderText="訓練起日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#B0E2FF" Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" SortExpression="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#B0E2FF" Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="管控單位">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="機構名稱">
                                            <HeaderStyle ForeColor="#B0E2FF" Width="30%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" SortExpression="ClassName" HeaderText="班名">
                                            <HeaderStyle ForeColor="#B0E2FF" Width="30%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="審核狀態"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="VerReason" HeaderText="未通過原因"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="已轉班"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="but1" runat="server" Text="設定" CommandName="update" CssClass="linkbutton"></asp:LinkButton><br>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="DistID" HeaderText="轄區"></asp:BoundColumn>
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
                        <tr>
                            <td align="center"></td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
