<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_08_015.aspx.vb" Inherits="WDAIIP.SD_08_015" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_08_001</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('GetClass.aspx?RID=' + RID + '&SubmitBtn=Button1&OCIDField=OCIDValue');
        }

        function search() {
            var msg = '';

            if (document.form1.center.value == '') msg += '必須選擇 訓練機構\n';
            if (document.form1.LSDate.value == '' || document.form1.LEDate.value == '') {
                msg += '必須填寫 離退日期區間\n';
            }
            else {
                if (!checkDate(document.getElementById('LSDate').value)) {
                    msg += '離退「起始」日期 不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
                if (!checkDate(document.getElementById('LEDate').value)) {
                    msg += '離退「迄止」日期 不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
                if (msg == '') {
                    if (getDiffDay(getAdDate(document.getElementById('LSDate').value), getAdDate(document.getElementById('LEDate').value)) < 0) {
                        msg1 += '離退日期區間(起)日不得大於(迄)日!\n';
                    }
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;學員動態管理&gt;&gt;職業訓練生活津貼&gt;&gt;<font color="#990000">離退清冊</font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="SearchTable" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td>
                                <table class="table_nw" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td width="100" class="bluecol_need">離退日期
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="LSDate" runat="server" onfocus="this.blur()" MaxLength="10" Columns="10"></asp:TextBox>
                                            <img id="ImgLSDate" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" width="30" height="30" align="top" runat="server" />～
                                        <asp:TextBox ID="LEDate" runat="server" onfocus="this.blur()" MaxLength="10" Columns="10"></asp:TextBox>
                                            <img id="ImgLEDate" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" width="30" height="30" align="top" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="100" class="bluecol_need">訓練機構
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox>
                                            <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                            <input id="Button7" type="button" value="..." name="Button7" runat="server" class="asp_button_Mini" />
                                            <asp:Button ID="Button8" Style="display: none" runat="server" Text="查詢上一次的列表"></asp:Button><br>
                                            <span id="HistoryList2" style="position: absolute; display: none">
                                                <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                                </asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">職類/班別
                                        </td>
                                        <td class="whitecol">
                                            <input id="OCIDValue" type="hidden" runat="server" />
                                            <input onclick="choose_class()" type="button" value="挑選班級" class="asp_button_S">
                                        </td>
                                    </tr>
                                </table>
                                <p style="margin-top: 3px; margin-bottom: 3px" align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label><asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                                    <asp:Button ID="Button3" runat="server" Text="列印" Visible="False" CssClass="asp_Export_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>

                    <br style="line-height: 5px">
                    <table class="font" id="msgt" cellspacing="0" bordercolordark="#ffffff" cellpadding="1" width="100%" align="center" bordercolorlight="#666666" border="0" runat="server">
                        <tr>
                            <td id="td1" align="left" colspan="2" runat="server"></td>
                        </tr>
                        <tr>
                            <td id="Td2" align="left" width="50%" runat="server"></td>
                            <td id="Td3" align="left" width="50%" runat="server"></td>
                        </tr>
                    </table>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" BorderColor="Black" AutoGenerateColumns="False">
                        <FooterStyle Wrap="False"></FooterStyle>
                        <ItemStyle BackColor="White"></ItemStyle>
                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="編&lt;br&gt;號">
                                <HeaderStyle Width="2%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="no" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="ClassName" HeaderText="參訓班別">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IdentityName" HeaderText="申請身&lt;br&gt;分別">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                <HeaderStyle Width="8%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="出生日期">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="birth" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證&lt;br&gt;編號">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="受訓起&lt;br&gt;迄日期">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="sd" runat="server"></asp:Label>
                                    <br>
                                    至
                                <asp:Label ID="ed" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="OPayMoney" HeaderText="原核發&lt;br&gt;金額">
                                <HeaderStyle Width="8%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="審核通過&lt;br&gt;後實際已&lt;br&gt;領取金額">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="FinPayMoney" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="RtnMoney" HeaderText="本次退&lt;br&gt;回金額">
                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="退訓日期&lt;br&gt;退訓原因">
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="RtnData" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle HorizontalAlign="Center"></PagerStyle>
                    </asp:DataGrid>
                    <table class="font" id="Memo" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td style="width: 619px" colspan="2">上列人數共計<u><asp:Label ID="PeopleNum" runat="server"></asp:Label></u>人，合計新臺幣<u><asp:Label ID="PeopleMoney" runat="server"></asp:Label></u>元整
                            </td>
                        </tr>
                        <tr>
                            <td width="35%"></td>
                            <td style="width: 365px" align="left" width="365">承辦人員：
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 619px" colspan="2">說明：本清冊請分別填繕一份，並加蓋承辦人員職章
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
