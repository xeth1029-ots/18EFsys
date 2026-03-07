<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_008.aspx.vb" Inherits="WDAIIP.SYS_06_008" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>ECFA廠商名單檢視</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;ECFA廠商名單檢視</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Panel ID="panelEdit" runat="server">
                        <table class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">SeqNo<font color="red">*</font></td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="SeqNo" runat="server" Columns="10" MaxLength="10" Width="40%"></asp:TextBox></td>
                                <td class="bluecol" width="20%">EcfaID</td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox Style="z-index: 0" ID="EcfaID" runat="server" Columns="10" MaxLength="10" Width="40%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">勞保投保證號(公司保險證號)</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Ubno" runat="server" Columns="10" MaxLength="20" Width="30%"></asp:TextBox>
                                    &nbsp;<asp:Label Style="z-index: 0" ID="bUbno" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">工廠登記證號</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="factoryNo" runat="server" Columns="10" MaxLength="20" Width="30%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">認定類別<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:RadioButtonList ID="CATEGORY" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                        <%--<asp:ListItem Value="0" Selected="True">不拘</asp:ListItem>--%>
                                        <asp:ListItem Value="1">加強輔導型產業</asp:ListItem>
                                        <asp:ListItem Value="2">可能受貿易自由化影響產業</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;產（行）業別<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="kName" runat="server" Columns="30" MaxLength="30" Width="40%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;廠商名稱<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="UName" runat="server" Columns="30" MaxLength="30" Width="40%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;統一編號<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="ComIDNO" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;主要產品<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Mproduct" runat="server" Columns="50" MaxLength="1000" Width="60%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;耗用原料</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Consumable" runat="server" Columns="50" MaxLength="1000" Width="60%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;地址<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Address" runat="server" Columns="60" MaxLength="300" Width="60%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;負責人<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="tMaster" runat="server" Columns="50" MaxLength="50" Width="20%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;工廠電話</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="phone" runat="server" Columns="30" MaxLength="100" Width="30%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;員工人數</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="MemNum" runat="server" Columns="3" MaxLength="5" Width="10%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;網址</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Url1" runat="server" Columns="30" MaxLength="200" Width="60%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;產業認定日<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="maintainDate" runat="server" Columns="10" MaxLength="10" Width="16%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">&nbsp;離職判斷日<font color="red">*</font></td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox Style="z-index: 0" ID="judgmentDate" runat="server" Columns="10" MaxLength="10" Width="16%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;是否歇業</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox Style="z-index: 0" ID="isClose" runat="server" Columns="2" MaxLength="1" Width="6%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;異動日期</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="modifyDate" Style="z-index: 0" runat="server" Columns="10" MaxLength="10" Width="16%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center" width="100%">
                                    <asp:Button ID="btnBack" Text="回上頁" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button Style="z-index: 0" ID="btnSave1" Text="存檔" runat="server" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center" width="100%">
                                    <asp:Label ID="Label1" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="panelSearch" runat="server">
                        <table id="Table3" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;序號 或&nbsp;EcfaID</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="EcfaID_s" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;保險證號</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Ubno_s" runat="server" Columns="10" MaxLength="20" Width="30%"></asp:TextBox>&nbsp;(有*號代表光碟檔有公司保險證號)</td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;工廠登記證號</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="factoryNo_s" runat="server" Columns="10" MaxLength="20" Width="30%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">認定類別</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:RadioButtonList ID="CATEGORY_s" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                        <asp:ListItem Value="0" Selected="True">不拘</asp:ListItem>
                                        <asp:ListItem Value="1">加強輔導型產業</asp:ListItem>
                                        <asp:ListItem Value="2">可能受貿易自由化影響產業</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;產（行）業別</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="kName_s" runat="server" Columns="30" MaxLength="30" Width="40%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;廠商名稱</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="UName_s" runat="server" Columns="30" MaxLength="30" Width="40%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;統一編號</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="ComIDNO_s" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">&nbsp;地址</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="Address_s" runat="server" Columns="60" MaxLength="100" Width="60%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">日期區間</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="MDate1" runat="server" Columns="8" Width="16%"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('MDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" height="30" width="30"></span>～
                                    <asp:TextBox ID="MDate2" runat="server" Columns="8" Width="16%"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('MDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" height="30" width="30"></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">匯入ECFA名單</td>
                                <td class="whitecol" colspan="3">
                                    <input id="File1" type="file" size="80" name="File1" runat="server" accept=".xls,.ods" />
                                    <asp:Button ID="BTN_IMP_ECFA_1" runat="server" Text="匯入名單" CssClass="asp_button_M"></asp:Button>
                                    <asp:HyperLink ID="HyperLink2" runat="server" NavigateUrl="../../Doc/ECFA28名單.zip" ForeColor="#8080FF">(匯入檔案必須為ods或xls格式)</asp:HyperLink>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center" width="100%">
                                    <asp:Button ID="btnSearch2" Text="今日資料查詢" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnSearch" Text="查詢" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="BtnAdd" Text="新增" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="BTN_EXP_ECFA_1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center" width="100%">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </table>
                        <table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td align="center">
                                    <div id="Div1" runat="server">
                                        <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" Width="100%" PageSize="20" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
                                            <HeaderStyle CssClass="head_navy" />
                                            <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                                            <Columns>
                                                <asp:BoundColumn DataField="SeqNo" HeaderText="流水號">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="EcfaID" HeaderText="EcfaID">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Ubno" HeaderText="保險證號">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="factoryNo" HeaderText="工廠登記證號">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="kName" HeaderText="產業別內容">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="UName" HeaderText="廠商名稱">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="ComIDNO" HeaderText="統一編號">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Mproduct" HeaderText="主要產品">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Consumable" HeaderText="耗用原料">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Master" HeaderText="負責人">
                                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle Width="16%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="btnView1" runat="server" Text="檢視" CommandName="View1" CssClass="asp_button_M"></asp:Button>
                                                        <asp:Button ID="btnCopy1" runat="server" Text="複製" CommandName="Copy1" CssClass="asp_button_M"></asp:Button>
                                                        <asp:Button ID="btnUPT1" runat="server" Text="修改" CommandName="UPT1" CssClass="asp_button_M"></asp:Button>
                                                        <asp:Button ID="btnDel1" runat="server" Text="刪除" CommandName="Del1" CssClass="asp_button_M"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                    </div>
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <%--btnUPT1--%>
        <asp:HiddenField ID="Hid_SEQNO" runat="server" />
    </form>
</body>
</html>
