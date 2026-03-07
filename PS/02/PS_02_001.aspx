<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="PS_02_001.aspx.vb" Inherits="WDAIIP.PS_02_001" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>功能頁面設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        function Check_Num() {
            var errMsg = "";
            var len = $('input:checkbox:checked').length;
            if (len < 3) {
                errMsg += "請選擇至少三個圖表!";
            }
            if (len > 3) {
                errMsg += "只能選擇三個圖表!";
            }
            if (errMsg == "") {
                return true;
            } else {
                alert(errMsg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;個人化設定&gt;&gt;首頁視覺化圖表設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="tb_View" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="auto-style1">請選擇首頁要顯示的<font color="#FF0000">三</font>個視覺化統計圖表:</td>
                                    </tr>
                                </table>
                                <div>
                                    <%--<asp:SqlDataSource runat="server" ID="DataSource" DataSourceMode="DataReader" ConnectionString="<%$ ConnectionStrings:ConnectionString %>" SelectCommand="SELECT CRTID, CNAME, PIC FROM [VISUALCHART] WHERE ISUSED='Y' ORDER BY SORT "/>--%>
                                    <%--DataKeyNames="CRTID" DataSourceID="DataSource"--%>
                                    <asp:ListView ID="ListView1" runat="server" AutoGenerateColumns="False" GroupItemCount="3">
                                        <GroupTemplate>
                                            <tr>
                                                <asp:PlaceHolder runat="server" ID="itemPlaceholder" />
                                            </tr>
                                        </GroupTemplate>
                                        <LayoutTemplate>
                                            <table id="table_chart" runat="server">
                                                <tr id="groupPlaceholder" runat="server"></tr>
                                            </table>
                                        </LayoutTemplate>
                                        <ItemTemplate>
                                            <td>
                                                <asp:CheckBox ID="CHECK" type="checkbox" Style="align-items: center" runat="server" />
                                                <asp:Label ID="CRTID" hidden="true" runat="server" Text='<%# Eval("CRTID") %>' />
                                                <asp:Label ID="CNAME" runat="server" Style="align-items: center" Text='<%# Eval("CNAME") %>' />
                                                <asp:Label ID="myTooltip" ForeColor="Red" runat="server" Text='(預設)' ToolTip='<%# Eval("DEFAULTSHOW") %>' />
                                            </td>
                                            <td>
                                                <asp:Image ID="PIC" runat="server" ImageUrl='<%# "~/PS/02/" + Eval("PIC") + ".jpg" %>' Width="150" />
                                            </td>
                                        </ItemTemplate>
                                    </asp:ListView>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center" colspan="4">
                                <div id="save" class="whitecol">
                                    <%--<asp:Button ID="btn_save" runat="server" OnClientClick="return Check_Num();" OnClick="btn_save_Click" />--%>
                                    <asp:Button ID="btn_save" type="button" OnClientClick="return Check_Num();" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>