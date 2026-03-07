<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_006.aspx.vb" Inherits="WDAIIP.SYS_01_006" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>EMAIL通知對象維護</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;EMAIL通知對象維護</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="Panel_Sch" runat="server" Visible="True">
            <table id="table_Sch" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                <tr>
                    <td class="bluecol" style="width: 20%">訓練計畫</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlTPlanSCH" runat="server">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">轄區分署</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddl_DistID" runat="server">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">功能1
                    </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rbl_EMAILCODE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList>
                        <asp:Label ID="labERR_EMAILCODE" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">帳號姓名
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtACCTNAME" runat="server" MaxLength="33" Width="33%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">帳號ID
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtACCTID" runat="server" MaxLength="33" Width="33%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">排除帳號停用
                    </td>
                    <td class="whitecol">
                        <asp:CheckBox ID="CB_ISUSED_N_NOSHOW" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">功能狀態
                    </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="RBL_FUNC_USE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Selected="True" Value="A">不區分</asp:ListItem>
                            <asp:ListItem Value="Y">啟用</asp:ListItem>
                            <asp:ListItem Value="N">停用</asp:ListItem>
                        </asp:RadioButtonList>

                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol" colspan="2">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol" colspan="2">
                        <asp:Label ID="msg1" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
            <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                <tr>
                    <td align="center">
                        <%--序號、轄區分署、功能、寄發對象、啟用(Y/N)、功能(修改鈕)--%>
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AllowSorting="True" CssClass="font"
                            AutoGenerateColumns="False" CellPadding="8">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            <AlternatingItemStyle BackColor="#F5F5F5" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:BoundColumn HeaderText="序號">
                                    <ItemStyle HorizontalAlign="Center" Width="6%" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署">
                                    <ItemStyle HorizontalAlign="Center" Width="16%" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="ECNAME" HeaderText="功能">
                                    <ItemStyle HorizontalAlign="Center" Width="14%" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="ACCTNAME" HeaderText="寄發對象">
                                    <ItemStyle HorizontalAlign="Center" Width="30%" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="FGUSE_N" HeaderText="啟用(Y/N)">
                                    <ItemStyle HorizontalAlign="Center" Width="10%" />
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <ItemStyle HorizontalAlign="Center" Width="8%" />
                                    <ItemTemplate>
                                        <%--<asp:LinkButton ID="btnEDIT" runat="server" Text="修改" CssClass="linkbutton" CommandName="btnEDIT"></asp:LinkButton>--%>
                                        <asp:LinkButton ID="btnUSED" runat="server" Text="啟用" CssClass="linkbutton" CommandName="btnUSED"></asp:LinkButton>
                                        <asp:LinkButton ID="btnNOUSE" runat="server" Text="停用" CssClass="linkbutton" CommandName="btnNOUSE"></asp:LinkButton>
                                        <asp:LinkButton ID="btnDELE" runat="server" Text="刪除" CssClass="linkbutton" CommandName="btnDELE"></asp:LinkButton>
                                        <asp:LinkButton ID="btnEMAIL" runat="server" Text="EMAIL查詢" CssClass="linkbutton" CommandName="btnEMAIL"></asp:LinkButton>
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
        </asp:Panel>
        <asp:Panel ID="Panel_edit" runat="server" Visible="True">
            <table id="table1" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                <tr>
                    <td class="bluecol" style="width: 20%">訓練計畫</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlTPlan" runat="server">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">轄區分署</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddl_DistID2" runat="server" AutoPostBack="True">
                        </asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td class="bluecol">帳號角色
                    </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rdo_role" runat="server" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">帳號
                    </td>
                    <td class="whitecol">
                        <asp:DropDownList ID="DDL_ACCOUNT1" runat="server"></asp:DropDownList>
                        <asp:Button ID="bt_data_sch2" runat="server" CssClass="asp_button_M" Text="重新查詢" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">功能1
                    </td>
                    <td class="whitecol">
                        <asp:CheckBoxList ID="CBL_EMAILCODE2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        </asp:CheckBoxList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">啟用狀態
                    </td>
                    <td class="whitecol">
                        <%--<asp:CheckBox ID="CB_FGUSE" runat="server" />--%>
                        <asp:Label ID="lab_FGUSE" runat="server" Text="(勾選後預設為啟用,不勾選則為停用)"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="2" class="whitecol">
                        <asp:Button ID="bt_backoff" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:HiddenField ID="Hid_EASEQ" runat="server" />
        <asp:HiddenField ID="Hid_ECSEQ" runat="server" />
        <asp:HiddenField ID="Hid_DISTID" runat="server" />
        <asp:HiddenField ID="Hid_ACCOUNT" runat="server" />
        <asp:HiddenField ID="Hid_EMAIL" runat="server" />
        <%--<asp:HiddenField ID="Hid_FGUSE" runat="server" />--%>
    </form>
</body>
</html>
