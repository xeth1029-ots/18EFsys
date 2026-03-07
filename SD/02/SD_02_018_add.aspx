<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_018_add.aspx.vb" Inherits="WDAIIP.SD_02_018_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>錄訓名單審核</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function savedataCHK1() {
            var rst1 = true; //正常再次檢核。
            var vMsgchk1 = "解鎖學員錄訓作業：首頁>>學員動態管理>>招生作業>>錄訓作業，已審核錄訓名單，該班可以再進行學員錄訓!!";
            rst1 = confirm(vMsgchk1);
            return rst1;
        }

        function savedataCHK2() {
            var rst1 = true; //正常再次檢核。
            var vMsgchk1 = "審核確認：首頁>>學員動態管理>>招生作業>>錄訓名單審核，已審核錄訓名單，該班不可以再進行學員錄訓!!";
            rst1 = confirm(vMsgchk1);
            return rst1;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					    <tr>
						    <td>
							    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
							    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;<FONT color="#990000">錄訓名單審核</FONT></asp:Label>
						    </td>
					    </tr>
				    </table>
                    --%>
                    <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="LabClassName1" runat="server" CssClass="font"></asp:Label><br />
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <%-- <AlternatingItemStyle BackColor="#F5F5F5" /> <ItemStyle /> <HeaderStyle CssClass="head_navy" />--%>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>准考證號碼</HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="labEXAMNO" runat="server"></asp:Label>
                                                <input id="SETID" type="hidden" runat="server" />
                                                <input id="ENTERDATE" type="hidden" runat="server" />
                                                <input id="SERNUM" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labStdName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證字號">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labIDNO" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:TemplateColumn HeaderText="是否加權">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labEXAMPLUS" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分別">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labEIdentity" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                        <asp:TemplateColumn HeaderText="成績">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labSUMOFGRAD" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="錄訓結果"><%--甄試結果/錄訓結果--%>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labSELRESULT_N" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table width="100%" class="table_nw">
                        <tr>
                            <td align="center" class="whitecol">
                                <br />
                                <asp:Button ID="BtnSave1X" runat="server" Text="解鎖學員錄訓作業" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="BtnSave2X" runat="server" Text="審核確認" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                &nbsp;<asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></td>
                        </tr>
                        <tr id="trODNUMBER" runat="server">
                            <td class="whitecol">&nbsp;&nbsp; <font color="red">(必填) 公文文號：</font><br />
                                <asp:TextBox ID="ODDATE1" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img id="imgODDATE1" runat="server" style="cursor: pointer" onclick="javascript:show_calendar('ODDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                &nbsp;<asp:TextBox ID="ODNUMBER" runat="server" MaxLength="15" Columns="20" Width="20%"></asp:TextBox>
                                字&nbsp;第&nbsp;<asp:TextBox ID="ODNUMBER2" runat="server" MaxLength="15" Columns="20" Width="20%"></asp:TextBox>號&nbsp;函&nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <br />
                                <asp:Button ID="BtnSave3X" runat="server" Text="公告" CssClass="asp_button_M"></asp:Button>&nbsp; </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_CFGUID" runat="server" />
        <asp:HiddenField ID="Hid_OCID" runat="server" />
        <asp:HiddenField ID="Hid_CFSEQNO" runat="server" />
        <asp:HiddenField ID="Hid_NOLOCK" runat="server" />
        <asp:HiddenField ID="Hid_ROVEDDATE" runat="server" />
        <asp:HiddenField ID="Hid_ANNMENTDATE" runat="server" />
    </form>
</body>
</html>
