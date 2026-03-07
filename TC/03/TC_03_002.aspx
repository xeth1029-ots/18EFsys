<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_03_002.aspx.vb" Inherits="WDAIIP.TC_03_002" %>

<html>
<head>
    <title>班級複製作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級複製作業</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" style="width: 20%">年度 </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="PlanYear" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr id="trddlDIST" runat="server">
                <td class="bluecol" style="width: 20%">轄區分署 </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="ddlDIST" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練機構 </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%">
                        </asp:Table>
                    </span></td>
            </tr>
            <tr>
                <td class="bluecol">
                    <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                    <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server" class="button_b_Mini">
                    <input id="trainValue" type="hidden" name="trainValue" runat="server">
                    <input id="jobValue" type="hidden" name="jobValue" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">
                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">班級名稱 </td>
                <td class="whitecol" style="width: 30%">
                    <asp:TextBox ID="ClassName" runat="server" Columns="30" Width="60%"></asp:TextBox>
                </td>
                <td class="bluecol" style="width: 20%">期別 </td>
                <td class="whitecol" style="width: 30%">
                    <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="40%"></asp:TextBox>
                </td>
            </tr>
            <tr id="trCOPYSUB" runat="server">
                <td class="bluecol" style="width: 20%">是否複製 </td>
                <td class="whitecol" colspan="3">
                    <asp:CheckBoxList ID="CBL_COPYSUB" runat="server"></asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol" colspan="4">
                    <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td><%--ForeColor="#B0E2FF" ForeColor="#E7F6FF"  --%>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowSorting="True" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear" SortExpression="PlanYear" HeaderText="計畫年度">
                                            <HeaderStyle ForeColor="#E7F6FF" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" SortExpression="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#E7F6FF" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="訓練起日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#E7F6FF" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" SortExpression="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#E7F6FF" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="機構名稱">
                                            <HeaderStyle ForeColor="#E7F6FF" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" SortExpression="ClassCName" HeaderText="班別名稱">
                                            <HeaderStyle Width="40%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="PointYN" SortExpression="PointYN" HeaderText="學分班"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="Button3" runat="server" Text="複製" CommandName="copy" CssClass="linkbutton"></asp:LinkButton>
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

                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol" colspan="4">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <%--<tbody><tr><td align="center"><table class="table_sch" id="Table2" cellspacing="1" cellpadding="1"><tbody></tbody></table><table width="100%"></table></td></tr></tbody>--%>
        <%--<input type="hidden" id="hidTC03002PlanID" runat="server" />--%>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server" />

    </form>
</body>
</html>
