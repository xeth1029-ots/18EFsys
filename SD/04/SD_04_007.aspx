<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_007.aspx.vb" Inherits="WDAIIP.SD_04_007" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>作息時間設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function checkaudit1() {
            //debugger;
            var myaudit1 = document.getElementById('tr_audit1');
            if (myaudit1) { myaudit1.style.display = 'none'; }
            //if (document.form1.IsApprPaper.item(0).checked == true) { myaudit1.style.display = 'inline'; }
        }
    </script>
</head>
<body>
    <table cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
            <td>
                <form id="form1" method="post" runat="server">
                    <%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0"><tr><td colspan="4">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;<font color="#990000">作息時間設定</font></td></tr></table>--%>
                    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;作息時間設定</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table cellpadding="1" cellspacing="1" class="table_nw">
                        <tr>
                            <td class="bluecol" style="width: 20%">年度
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>

                            </td>
                            <td class="bluecol" style="width: 20%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TB_career_id" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server" class="button_b_Mini">&nbsp;
							<input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">&nbsp;
							    <asp:Button ID="Btn_OrgSet" runat="server" Text="機構作息時間設定" CssClass="asp_button_M"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%">
                                    </asp:Table>
                                </span>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="TPlanid" type="hidden" name="TPlanid" runat="server" />
                                <input id="Re_ID" type="hidden" name="Re_ID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" Width="40%"></asp:TextBox>
                            </td>
                            <td class="bluecol">期別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="30%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr style="display: none">
                            <td class="bluecol" id="dt_datatype" runat="server">資料類型
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="IsApprPaper" runat="server" Visible="False" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="Y" Selected="True">正式</asp:ListItem>
                                    <asp:ListItem Value="N">草稿</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_audit1" style="display: none" runat="server">
                            <td class="bluecol" id="td_audit" runat="server">審核狀態
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="audit" runat="server" Visible="False" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
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
                            <td>訓練計畫：
							<asp:Label ID="TPlanName" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left" PagerStyle-Mode="NumericPages" AllowPaging="True" OnItemCommand="dtPlan_ItemCommand" AutoGenerateColumns="False" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="5%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="計畫年度">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" SortExpression="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#E3F8FD" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="訓練職類">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="訓練起日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#E3F8FD" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" SortExpression="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}">
                                            <HeaderStyle ForeColor="#E3F8FD" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="管控單位"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="機構名稱">
                                            <HeaderStyle ForeColor="#E3F8FD" Width="20%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" SortExpression="ClassName" HeaderText="班名">
                                            <HeaderStyle ForeColor="#E3F8FD" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="審核狀態"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="VerReason" HeaderText="未通過原因"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="已轉班"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="but1" runat="server" Text="設定" CommandName="update" CssClass="asp_button_M"></asp:Button><br>
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
                            <td align="center">
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </form>
            </td>
        </tr>
    </table>
</body>
</html>
