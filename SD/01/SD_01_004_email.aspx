<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_004_email.aspx.vb" Inherits="WDAIIP.SD_01_004_email" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>e 網審核郵件設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
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
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;e 網審核郵件設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table_nw" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <asp:Panel ID="table3" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">年度 </td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="ddlyears" runat="server"></asp:DropDownList>
                                    <input id="RIDValue" style="width: 24%;" type="hidden" name="RIDValue" runat="server" class="asp_button_M" />
                                </td>
                                <td class="bluecol" width="20%"><asp:Label ID="labtmid" runat="server">訓練職類</asp:Label></td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server" class="asp_button_Mini" />&nbsp;
                                    <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                    <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%"><asp:Label ID="labcjob_unkey" runat="server">通俗職類 </asp:Label></td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">訓練機構 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Columns="45" Width="60%"></asp:TextBox>
                                    <input id="Org" type="button" value="..." name="Org" runat="server" class="asp_button_Mini" />&nbsp;
                                    <input id="TPlanid" style="width: 10%" type="hidden" name="TPlanid" runat="server" />
                                    <input id="Re_ID" style="width: 10%" type="hidden" name="Re_ID" runat="server" />
                                    <asp:Button ID="btn_orgset" runat="server" Text="機構e網郵件設定" CssClass="asp_button_L"></asp:Button>
                                    <span id="HistoryList2" style="display: none; position: absolute"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱 </td>
                                <td class="whitecol"><asp:TextBox ID="ClassName" runat="server" Width="40%"></asp:TextBox></td>
                                <td class="bluecol">期別 </td>
                                <td class="whitecol"><asp:TextBox ID="CyclType" runat="server" Columns="5" Width="30%"></asp:TextBox></td>
                            </tr>
                            <tr style="display: none">
                                <td class="bluecol">資料類型 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="IsApprPaper" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow" Visible="false">
                                        <asp:ListItem Value="Y" Selected="true">正式</asp:ListItem>
                                        <asp:ListItem Value="N">草稿</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr id="tr_audit1" style="display: none" runat="server">
                                <td class="bluecol" id="td_audit" runat="server">審核狀態 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="audit" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow" Visible="false">
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
                                        <asp:Label ID="labpagesize" runat="server" ForeColor="slateblue">顯示列數</asp:Label>
                                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>訓練計畫：<asp:Label ID="TPlanName" runat="server" CssClass="font"></asp:Label></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" OnItemCommand="dtPlan_ItemCommand" AllowPaging="true" PagerStyle-Mode="numericpages" PagerStyle-HorizontalAlign="left" AllowSorting="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <ItemStyle HorizontalAlign="center" Width="4%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="planyear" HeaderText="計畫年度">
                                            <ItemStyle HorizontalAlign="center" Width="6%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="applieddate" SortExpression="applieddate" HeaderText="申請日期" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="center" Width="8%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" HeaderText="訓練職類"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="stdate" SortExpression="stdate" HeaderText="訓練起日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="center" Width="8%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="fddate" SortExpression="fddate" HeaderText="訓練迄日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="center" Width="8%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="管控單位">
                                            <ItemStyle HorizontalAlign="center" Width="14%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="orgname" SortExpression="orgname" HeaderText="機構名稱">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="center" Width="14%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSNAME2" SortExpression="CLASSNAME2" HeaderText="班名">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="center" Width="14%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" HeaderText="審核狀態">
                                            <ItemStyle HorizontalAlign="center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" DataField="verreason" HeaderText="未通過原因">
                                            <ItemStyle HorizontalAlign="center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" HeaderText="已轉班"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Width="6%" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="but1" runat="server" Text="設定" CommandName="update" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:BoundColumn Visible="false" DataField="distid" HeaderText="轄區"></asp:BoundColumn>--%>
                                    </Columns>
                                    <PagerStyle Visible="false" HorizontalAlign="left" ForeColor="blue" Position="top" Mode="numericpages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td>
                        </tr>
                        <tr>
                            <td align="center"></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <div align="center" style="width: 100%"><asp:Label ID="msg" runat="server" CssClass="font" ForeColor="red"></asp:Label></div>
    </form>
</body>
</html>