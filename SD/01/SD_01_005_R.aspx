<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_01_005_R.aspx.vb" Inherits="WDAIIP.SD_01_005_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印報名學員郵遞標籤</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button13').click();
        }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?special=2&RID=' + RID);
        }

        function search() {
            var msg = '';
            if (isEmpty(document.form1.STDate1) && isEmpty(document.form1.STDate2) && isEmpty(document.form1.OCID1) && isEmpty(document.form1.IDNO)) {
                msg += '查詢條件請選擇其中之一!\n';
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
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">列印報名學員郵遞標籤</font>
                    </asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="119">報名日期區間
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="STDate1" runat="server" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />~
                            <asp:TextBox ID="STDate2" runat="server" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="100">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input type="button" value="..." onclick="choose_class()" class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="119">身分證號碼
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" Width="200px" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <div style="width: 100%" align="center" class="whitecol">
            <asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button><br />
            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
        </div>
    </form>
</body>
</html>
