<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_010.aspx.vb" Inherits="WDAIIP.SD_01_010" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>報名作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            //openClass('../02/SD_02_ch.aspx?RWClass=1&RID='+RID);openClass('../02/SD_02_ch.aspx?RID='+RID+'&BtnName=Button2');
            openClass('../02/SD_02_ch.aspx?RWClass=1&RID=' + RID + '&BtnName=btnCheckClass');
        }

        function ChkData() {
            var msg = '';
            if (isEmpty('OCIDValue1')) { msg += '請選擇班級\n'; }
            if (isEmpty('IDNO')) { msg += '請輸入身分證字號!!!\n'; }
            if (isEmpty('birthDay')) { msg += '請輸入出生日期\n'; }
            //if (isEmpty('EnterDate')) { msg += '請輸入報名日期\n'; }
            if (msg !== '') {
                alert(msg);
                return false;
            }
            //鎖定按鈕邏輯 
            var btn = document.getElementById('<%= btnAdd1.ClientID %>');
            var btnD = document.getElementById('btnAdd1D');
            if (btn && btnD) { btn.style.display = 'none';btnD.style.display = '';}
            return true;
        }

        function sChkData1() {
            document.form1.hidCheck1.value = '';
        }
    </script>
    <%--<style type="text/css">.auto-style1 {color: #333333;padding: 4px;width: 300px;}</style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;報名作業</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button8" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">班級名稱</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 25%">
                                    <asp:Table ID="Historytable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">身分證號碼</td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" Width="20%"></asp:TextBox>
                                <input id="hidCheck1" type="hidden" name="hidCheck1" runat="server" />
                                <asp:Button ID="btnCheck1" runat="server" Text="台端資料確認" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">出生日期</td>
                            <td class="whitecol">
                                <asp:TextBox ID="birthDay" runat="server" Width="15%"></asp:TextBox>
                                <span id="date1" runat="server">
                                    <img title="點選日期" style="cursor: pointer" onclick="javascript:show_calendar('<%= birthDay.ClientId %>','','','CY/MM/DD');" alt="點選日期" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="btnCheckClass" runat="server" Text="班級確認" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnAdd1" runat="server" Text="報名" CssClass="asp_button_M"></asp:Button>
                                <input id="btnAdd1D" runat="server" type="button" value="處理中" class="asp_button_M" style="display:none" disabled="disabled" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Literal ID="JAVASCRIPT_LITERAL" runat="server"></asp:Literal>
        <input id="isBlack" type="hidden" name="isBlack" runat="server">
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server">
    </form>
</body>
</html>
