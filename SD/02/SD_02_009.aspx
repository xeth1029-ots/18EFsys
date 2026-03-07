<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_02_009.aspx.vb" Inherits="WDAIIP.SD_02_009" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>交寄大宗函件資料列印</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button12').click();
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁>>學員動態管理>>招生作業>>交寄大宗函件資料列印</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr id="org_tr" runat="server">
                            <td class="bluecol_need" width="100">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button8" type="button" value="..." name="Button5" runat="server" class="button_b_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="DistValue" type="hidden" name="DistValue" runat="server" />
                                <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue();">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="class_tr" runat="server">
                            <td class="bluecol" width="100">班級名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" onclick="choose_class();" type="button" value="..." name="Button5" runat="server" class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">通俗職類
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="100">收件人
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="consignee" runat="server" AutoPostBack="True" RepeatDirection="Horizontal" RepeatLayout="Flow" Font-Size="X-Small" CssClass="font">
                                    <asp:ListItem Value="00" Selected="True">甄試通知對象</asp:ListItem>
                                    <asp:ListItem Value="01">錄取通知對象</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4">
                                <div align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button><br />
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </div>
                            </td>
                        </tr>
                    </table>

                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center" Width="30%" VerticalAlign="Middle"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="30%" VerticalAlign="Middle"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="12%" VerticalAlign="Middle"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="12%" VerticalAlign="Middle"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="16%" VerticalAlign="Middle"></HeaderStyle>
                                            <ItemTemplate>
                                                <input id="OCID" type="hidden" size="2" runat="server" name="OCID" />
                                                <input id="consignee2" type="hidden" size="2" runat="server" name="consignee2" />
                                                <input id="Button3" type="button" value="列印執據" name="Button3" runat="server" class="asp_Export_M" />
                                                <input id="Button4" type="button" value="列印存根" name="Button4" runat="server" class="asp_Export_M" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server" Visible="False"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
