<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_020.aspx.vb" Inherits="WDAIIP.TC_01_020" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>修改班級聯絡資訊(產業人才專用)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
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
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;修改班級聯絡資訊</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="SchPanel2" runat="server" Width="100%">
            <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="bluecol" width="16%">訓練機構</td>
                    <td colspan="3" class="whitecol" width="84%">
                        <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                        <input id="Org" type="button" value="..." name="Org" runat="server" />
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                        <%--<input id="Orgidvalue" type="hidden" name="Orgidvalue" runat="server" />--%>
                        <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                        </span>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">班級名稱</td>
                    <td class="whitecol">
                        <asp:TextBox ID="ClassName" runat="server" Columns="44" MaxLength="55" Width="88%"></asp:TextBox></td>
                    <td class="bluecol" width="16%">期別</td>
                    <td class="whitecol" width="34%">
                        <asp:TextBox ID="CyclType" runat="server" Columns="10" MaxLength="3" Width="33%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">課程代碼</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="s_OCID" runat="server" Columns="20" MaxLength="10" Width="22%"></asp:TextBox></td>
                </tr>
                <tr id="tr_AppStage_TP28" runat="server">
                    <td class="bluecol">申請階段</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td colspan="4" class="whitecol" width="100%">
                        <div align="center">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                            <asp:Button ID="BtnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        </div>
                        <div align="center">
                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                        </div>
                    </td>
                </tr>
            </table>
        </asp:Panel>

        <asp:Panel ID="SchPanel" runat="server" Width="100%" Visible="False">
            <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" Visible="False" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5" />
                <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="編號">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Width="6%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ORGNAME" HeaderText="機構名稱">
                        <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別名稱">
                        <ItemStyle HorizontalAlign="Center" Width="22%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="STDATE" HeaderText="開訓日期">
                        <ItemStyle HorizontalAlign="Center" Width="16%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="FTDATE" HeaderText="結訓日期">
                        <ItemStyle HorizontalAlign="Center" Width="16%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Width="16%" Font-Size="Small"></ItemStyle>
                        <ItemTemplate>
                            <%--
                                <asp:HiddenField ID="Hid_PCS" runat="server" />
                            <asp:HiddenField ID="Hid_OCID" runat="server" />
                            <asp:HiddenField ID="Hid_MSD" runat="server" />
                            --%>
                            <asp:LinkButton ID="BTNUNLOCK" runat="server" Text="解鎖" CommandName="BTNUNLOCK" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="BTNREVISE" runat="server" Text="修改" CommandName="BTNREVISE" CssClass="linkbutton"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <div align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </div>
        </asp:Panel>
        <asp:Panel ID="EdtPanel1" runat="server" Width="100%" Visible="False">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" width="15%">訓練機構</td>
                    <td class="whitecol" width="35%">
                        <asp:TextBox ID="TB_OrgName" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox></td>
                    <td class="bluecol" width="15%">課程代碼</td>
                    <td class="whitecol" width="35%">
                        <asp:TextBox ID="TB_OCID" runat="server" onfocus="this.blur()" Columns="10" Width="60%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol" align="center">班級名稱</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_ClassCName" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox></td>
                    <td class="bluecol" align="center">期別</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_CYCLTYPE" runat="server" onfocus="this.blur()" Columns="3" MaxLength="2" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">聯絡人</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_ContactName" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                    <td class="whitecol"></td>
                    <td class="whitecol"></td>
                    <%-- <td class="bluecol">電話</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_ContactPhone" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>--%>
                </tr>
                <tr id="trContactPhone_2024_N1" runat="server">
                    <td class="bluecol_need">辦公室電話</td>
                    <td class="whitecol">
                        <asp:TextBox ID="ContactPhone_1" runat="server" MaxLength="10" Width="18%" ToolTip="區碼(2~4碼)" placeholder="區碼(0開頭)"></asp:TextBox>-
                                <asp:TextBox ID="ContactPhone_2" runat="server" MaxLength="10" Width="30%" ToolTip="電話(8碼內)" placeholder="電話(8碼)"></asp:TextBox>#
                                <asp:TextBox ID="ContactPhone_3" runat="server" MaxLength="10" Width="18%" ToolTip="分機(8碼內)" placeholder="分機"></asp:TextBox>
                    </td>
                    <td class="bluecol_need">行動電話</td>
                    <td class="whitecol">
                        <asp:TextBox ID="ContactMobile_1" runat="server" MaxLength="10" Width="18%" ToolTip="手機號碼前4碼" placeholder="手機前4碼(0開頭)"></asp:TextBox>-
                                <asp:TextBox ID="ContactMobile_2" runat="server" MaxLength="10" Width="30%" ToolTip="手機號碼後6碼" placeholder="手機後6碼"></asp:TextBox>
                    </td>
                </tr>
                <tr id="trContactPhone_2024_N2" runat="server">
                    <td class="whitecol"></td>
                    <td class="whitecol">
                        <asp:Label ID="lab_ContactPhone_m1" runat="server" Text="(【辦公室電話】、【行動電話】至少須擇一填寫)" ForeColor="Red"></asp:Label>
                    </td>
                    <td class="whitecol"></td>
                    <td class="whitecol">
                        <asp:Label ID="lab_ContactMobile_m2" runat="server" Text="(【辦公室電話】、【行動電話】至少須擇一填寫)" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">電子郵件 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_ContactEmail" runat="server" MaxLength="64" Width="80%"></asp:TextBox></td>
                    <td class="bluecol">傳真 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_ContactFax" runat="server" MaxLength="64" Width="60%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="center" colspan="4" class="whitecol">
                        <asp:Button ID="BtnSAVEDATA1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnBack1" runat="server" Text="回查詢頁面" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>

        <asp:HiddenField ID="Hid_OCID" runat="server" />
        <asp:HiddenField ID="Hid_PlanID" runat="server" />
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_SeqNo" runat="server" />
    </form>
</body>
</html>
