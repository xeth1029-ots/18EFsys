<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_005c.aspx.vb" Inherits="WDAIIP.SD_05_005c" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>查詢主課程</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function check_data() {
            var TextBox1 = document.getElementById('TextBox1');
            if (TextBox1.value != '@@') {
                if (TextBox1.value != '' && !isUnsignedInt(TextBox1.value)) {
                    alert('小時數必須為數字\n');
                    return false;
                }
            }
        }
        function result() {
            if (document.form1.ListBox2.options.length == 0) {
                alert('請選擇課程!');
                return false;
            }
        }

        function MoveItem(num) {
            switch (num) {
                case 1:
                    moveOption(document.form1.ListBox1, document.form1.ListBox2);
                    break;
                case 2:
                    moveOption(document.form1.ListBox2, document.form1.ListBox1);
                    break;
                case 3:
                    moveAllOption(document.form1.ListBox1, document.form1.ListBox2);
                    break;
                case 4:
                    moveAllOption(document.form1.ListBox2, document.form1.ListBox1);
                    break;
            }

            document.form1.course.value = '';
            for (var i = 0; i < document.form1.ListBox2.options.length; i++) {
                if (document.form1.course.value != '') { document.form1.course.value += ','; }
                document.form1.course.value += document.form1.ListBox2.options[i].value;
            }
            //alert(document.form1.course.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="whitecol">
                                <asp:Button ID="Button7" runat="server" Text="重建時間配當表" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">欲搜尋
							<asp:TextBox ID="TextBox1" runat="server" Width="10%" MaxLength="5">1</asp:TextBox>小時的主課程
							<asp:Button ID="Button1" runat="server" Text="查詢" ToolTip="重建時間配當請輸入'@@'" CssClass="asp_button_M"></asp:Button>&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<asp:Label ID="Label1" runat="server">評量學員成績科目 <BR>(*)表示已有成績資料</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td align="center">
                                            <asp:ListBox ID="ListBox1" runat="server" Width="200px" SelectionMode="Multiple" Rows="10"></asp:ListBox>
                                        </td>
                                        <td align="center">
                                            <input id="Button3" style="width: 30px; height: 30px" type="button" value=">" name="Button3" runat="server" class=""><br />
                                            <br />
                                            <input id="Button4" style="width: 30px; height: 30px" type="button" value="<" name="Button4" runat="server" class=""><br />
                                            <br />
                                            <input id="Button5" style="width: 30px; height: 30px" type="button" value=">>" name="Button5" runat="server" class=""><br />
                                            <br />
                                            <input id="Button6" style="width: 30px; height: 30px" type="button" value="<<" name="Button6" runat="server" class="">
                                        </td>
                                        <td align="center">
                                            <asp:ListBox ID="ListBox2" runat="server" Width="200px" SelectionMode="Multiple" Rows="10"></asp:ListBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="3" class="whitecol">
                                            <asp:Button ID="Button2" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <asp:Label ID="Label2" runat="server" ForeColor="Red" CssClass="font">	(請注意!!)<br />
											＊有成績資料，若移至左邊 為 「排除搜尋主課程」中，將在「儲存時」刪除該班該學員該課程成績！<br />
											＊課程資料設定 若有儲存主課程，分數將歸類為主課程！
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="course" type="hidden" runat="server" />
    </form>
</body>
</html>
