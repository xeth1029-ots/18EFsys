<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_016_R.aspx.vb" Inherits="WDAIIP.SD_09_016_R" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_09_016_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <%--link href="../../style.css" type="text/css" rel="stylesheet"--%>
    <link href="../../css/style.css" rel="stylesheet" type="text/css" />
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
        //判斷如果有勾申請則補助金額欄位是可輸入的
        function Check1() {
            var Mytable = document.getElementById("DataGrid1");

            for (var i = 1; i < Mytable.rows.length; i++) {
                Mycheckbox = Mytable.rows[i].cells[0].children[0];
                Mytextbox = Mytable.rows[i].cells[6].children[0];
                if (Mycheckbox.checked) { Mytextbox.disabled = false; }
                else {
                    Mytextbox.disabled = true;
                    Mytextbox.value = '';
                }

            }

        }

        //檢查是否只有一個班級用
        function GETvalue() { document.getElementById('Button12').click(); }

        //選擇班級用
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID);
        }

        //查詢及列印時的檢查
        function CheckSPrint() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') {
                msg += '請選擇班級職類\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //全選時
        function ChangeAll(obj) {

            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (obj.checked == true) {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        if (document.form1.elements[iCount].disabled == false) {
                            document.form1.elements[iCount].checked = true;

                        }
                    }
                    Check1();

                }
                else {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = false;

                    }

                    Check1();
                }
            }
        }

        //存檔時的檢查
        function Check_date() {
            var msg = '';
            var Mytable = document.getElementById("DataGrid1");

            for (var i = 1; i < Mytable.rows.length; i++) {
                Mycheckbox = Mytable.rows[i].cells[0].children[0];
                Mytextbox = Mytable.rows[i].cells[6].children[0];
                if (Mycheckbox.checked) {
                    if (Mytextbox.value == '') msg += '補助金額必須填寫(第' + i + '列)\n';
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
        <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;<font color="#990000">學員補助印領清冊</font></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="100">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12"></asp:Button>
                                <span id="HistoryList2" style="display: none; z-index: 100; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="100">班別/職類
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server">
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server">
                                <span id="HistoryList" style="display: none; z-index: 101; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <p align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Query" runat="server" Text="查詢" CssClass="button_b_M"></asp:Button>
                                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <p align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td style="font-size: 8pt; color: red">
                                <font face="新細明體"></font>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="是否補助">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <HeaderTemplate>
                                                是否補助<input id="CheckboxAll" type="checkbox" name="CheckboxAll" runat="server">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" value='<%# DataBinder.Eval(Container.DataItem,"socid")%>' name="Checkbox1" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Sex" HeaderText="姓別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MIdentityID" HeaderText="身分別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證字號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="birthday" HeaderText="出生年月日" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="補助金額">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="OtherJobCost" runat="server"></asp:TextBox>
                                                <asp:Label ID="Pmode2" runat="server" Text="*"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <asp:Button ID="Save" runat="server" Text="儲存" CssClass="button_b_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
