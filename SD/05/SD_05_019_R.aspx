<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_019_R.aspx.vb" Inherits="WDAIIP.SD_05_019_R" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓學員成績單</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
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
            document.getElementById('Button3').click();
        }

        function search() {
            var msg = ''
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') msg += '必須選擇班別職類\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function chall() {
            var mytable = document.getElementById('DataGrid1')
            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheck = mytable.rows[i].cells[0].children[0];
                if (mycheck.disabled == false)
                    mycheck.checked = document.form1.Choose1.checked
            }
        }

        function CheckPrint() {
            //Button2
            var RID = document.getElementById('RIDValue').value;
            var OCID = document.getElementById('OCIDValue1').value;
            var MyTable = document.getElementById('DataGrid1');
            var DistID = document.getElementById('DistID').value;
            var TPlanID = document.getElementById('TPlanID').value;
            var H1 = document.getElementById('H1').value;
            var H2 = document.getElementById('H2').value;
            var type = document.getElementById('type').value;
            var SOCID = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (SOCID != '') { SOCID += ',' }
                    SOCID += MyTable.rows[i].cells[0].children[1].value;
                }
            }
            if (SOCID == '') {
                alert('請選擇學員');
                return false;
            } else {
                //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=Member&filename=SD_05_019_R&&path=TIMS&SOCID='+SOCID+'&OCID='+OCID+'&TPlanID='+TPlanID);
                //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=Member&filename=SD_05_019_R&&path=TIMS&SOCID='+SOCID+'&OCID='+OCID+'&TPlanID='+TPlanID+'&DistID='+DistID+'&H1='+H1+'&H2='+H2+'&type='+type);
                /*報表以HTML呈現*/
                window.open('SD_05_019_R_1.aspx?&RID=' + RID + '&SOCID=' + SOCID + '&OCID=' + OCID + '&TPlanID=' + TPlanID + '&DistID=' + DistID + '&H1=' + H1 + '&H2=' + H2 + '&type=' + type);
            }
            //return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;受訓學員成績單</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="Button10" type="button" value="..." name="Button10" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">

                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="javascript: openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 28%; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                                <input id="DistID" type="hidden" name="DistID" runat="server">
                                <input id="TPlanID" type="hidden" name="TPlanID" runat="server">
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <div align="center" style="width: 100%">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tbody>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowCustomPaging="True" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="10%" />
                                                <HeaderTemplate>
                                                    <input onclick="chall();" type="checkbox" checked name="Choose1">
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <input id="Checkbox1" type="checkbox" checked name="student" runat="server" />
                                                    <input id="SOCID" type="hidden" name="SOCID" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="STUDID2" HeaderText="學號" HeaderStyle-Width="25%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="SNAME" HeaderText="姓名" HeaderStyle-Width="20%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼" HeaderStyle-Width="20%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="SEX2_N" HeaderText="性別" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Birthday" HeaderText="出生日期" HeaderStyle-Width="15%"></asp:BoundColumn>
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
                                <td>
                                    <p align="center" class="whitecol">
                                        <asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                    </p>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
        </table>
        <input id="H1" type="hidden" name="H1" runat="server" />
        <input id="H2" type="hidden" name="H2" runat="server" />
        <input id="type" type="hidden" name="type" runat="server" />
    </form>
</body>
</html>
