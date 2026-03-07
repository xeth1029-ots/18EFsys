<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_031.aspx.vb" Inherits="WDAIIP.SD_14_031" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓證書清冊</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function choose_class() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';

            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?&RID=' + RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;數位結訓證明</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Panel ID="PanelSCH" runat="server" Visible="True">
            <table class="table_sch">
                <tr>
                    <td class="bluecol" style="width: 20%">姓名</td>
                    <td class="whitecol">
                        <asp:TextBox ID="q_CNAME" runat="server" Width="22%" MaxLength="99"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">身分證號碼</td>
                    <td class="whitecol">
                        <asp:TextBox ID="q_IDNO" runat="server" Width="22%" MaxLength="13"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">申請日期區間</td>
                    <td class="whitecol">
                        <asp:TextBox ID="q_APPLIEDDATE1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('q_APPLIEDDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        ～<asp:TextBox ID="q_APPLIEDDATE2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('q_APPLIEDDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </td>
                </tr>
                <tr>
                    <td colspan="2" align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="BtnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
            <table class="table_sch">
                <tr>
                    <td align="center">
                        <div align="center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                        </div>
                        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#f5f5f5" />
                                        <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="CNAME" HeaderText="姓名" />
                                            <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證號碼" />
                                            <asp:BoundColumn DataField="DCASENO" HeaderText="案件編號" />
                                            <asp:BoundColumn DataField="APPLNDATE_F" HeaderText="申請日期時間" />
                                            <asp:TemplateColumn>
                                                <HeaderTemplate>功能</HeaderTemplate>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:Button ID="btnDATASHOW1" runat="server" Text="查看" CssClass="asp_Export_M" CommandName="DATASHOW1"></asp:Button>
                                                    <asp:Button ID="btnPrint1" runat="server" Text="列印封面" CssClass="asp_Export_M" CommandName="Print1"></asp:Button>
                                                    <asp:Button ID="btnPrint2" runat="server" Text="列印資料" CssClass="asp_Export_M" CommandName="Print2"></asp:Button>
                                                    <asp:HiddenField ID="hDCASENO" runat="server" />
                                                    <asp:HiddenField ID="hDCANO" runat="server" />
                                                    <asp:HiddenField ID="hEMVCODE" runat="server" />
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
                            <tr>
                                <td align="center">
                                    <%--<asp:CheckBox ID="AllPrint" runat="server" Font-Size="X-Small" Text="全部列印"></asp:CheckBox>--%>
                                    <%--<asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_button_S" />--%>
                                </td>
                            </tr>
                        </table>

                    </td>
                </tr>
            </table>
            <%--<div align="center" class="whitecol"></div>--%>
            <%--<asp:Label ID="labmsg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>--%>
        </asp:Panel>
        <asp:Panel ID="PanelVIEW" runat="server" Visible="True">
            <table id="tbPanelEdit1" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="table_title" colspan="2">線上申請-填寫申請資料</td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">姓名</td>
                    <td class="whitecol">
                        <asp:Label ID="labCNAME" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">身分證號碼</td>
                    <td class="whitecol">
                        <asp:Label ID="labIDNO" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">案件編號</td>
                    <td class="whitecol">
                        <asp:Label ID="labDCASENO" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">申請日期時間</td>
                    <td class="whitecol">
                        <asp:Label ID="labAPPLNDATE" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right">申請用途</td>
                    <td class="whitecol">
                        <%--<asp:DropDownList ID="ddlPURPOSE" runat="server"></asp:DropDownList>--%>
                        <asp:Label ID="labPURPOSE" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right">申請證明要提供的使用單位 </td>
                    <td class="whitecol">
                        <%--<asp:DropDownList ID="ddlUSAGEUNIT" runat="server"></asp:DropDownList>--%>
                        <asp:Label ID="labUSAGEUNIT" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right">申請日期 </td>
                    <td class="whitecol">
                        <asp:Label ID="labAPPLNDATE_TW" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">選擇班級數</td>
                    <td class="whitecol">
                        <asp:Label ID="labCLSCNT" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">下載次數</td>
                    <td class="whitecol">
                        <asp:Label ID="labDLCNT" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" align="right" width="20%">最後下載時間</td>
                    <td class="whitecol">
                        <asp:Label ID="labLASTDLTIME" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" align="center" class="whitecol">
                        <%--<asp:Button ID="btnSAVE1" runat="server" Text="確認" CssClass="asp_button_M"></asp:Button>&nbsp;--%>
                        <asp:Button ID="btnBACK1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:HiddenField ID="Hid_DCANO" runat="server" />
        <asp:HiddenField ID="Hid_DCASENO" runat="server" />
        <asp:HiddenField ID="Hid_EMVCODE" runat="server" />
        <%--<input id="Years" type="hidden" name="Years" runat="server" />--%>
    </form>
</body>
</html>
