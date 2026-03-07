<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_036c.aspx.vb" Inherits="WDAIIP.SD_05_036c" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>查詢 重大災害受災地區</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function result() {
            if (document.form1.ListBox2.options.length == 0) {
                alert('請選擇 重大災害受災地區!');
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

            var HID_ZIPCODES1 = document.form1.HID_ZIPCODES1;
            HID_ZIPCODES1.value = '';
            for (var i = 0; i < document.form1.ListBox2.options.length; i++) {
                if (HID_ZIPCODES1.value != '') { HID_ZIPCODES1.value += ','; }
                HID_ZIPCODES1.value += document.form1.ListBox2.options[i].value;
            }
            //alert(document.form1.course.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="88%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <%--<tr> <td>欲搜尋
							<asp:TextBox ID="TextBox1" runat="server" Width="40px" MaxLength="5">1</asp:TextBox>小時的主課程
							<asp:Button ID="Button1" runat="server" Text="查詢" ToolTip="重建時間配當請輸入'@@'"></asp:Button>&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<asp:Label ID="Label1" runat="server" Width="112px">評量學員成績科目 <BR>(*)表示已有成績資料</asp:Label>
						</td> </tr>--%>
                        <tr>
                            <td align="center">
                                <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                                    <tr>
                                        <td align="center">可挑選：<br />
                                            <asp:ListBox ID="ListBox1" runat="server" Width="90%" SelectionMode="Multiple" Rows="20"></asp:ListBox>
                                        </td>
                                        <td align="center">
                                            <p>
                                                <font face="新細明體">
                                                    <input id="Button3" style="width: 25px; height: 25px" type="button" value=">" name="Button3" runat="server"></font>
                                            </p>
                                            <p>
                                                <font face="新細明體">
                                                    <input id="Button4" style="width: 25px; height: 25px" type="button" value="<" name="Button4" runat="server"></font>
                                            </p>
                                            <p>
                                                <font face="新細明體">
                                                    <input id="Button5" style="width: 25px; height: 25px" type="button" value=">>" name="Button5" runat="server"></font>
                                            </p>
                                            <p>
                                                <font face="新細明體">
                                                    <input id="Button6" style="width: 25px; height: 25px" type="button" value="<<" name="Button6" runat="server"></font>
                                            </p>
                                        </td>
                                        <td align="center">已挑選：<br />
                                            <asp:ListBox ID="ListBox2" runat="server" Width="90%" SelectionMode="Multiple" Rows="20"></asp:ListBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="3">
                                            <asp:Button ID="Button2" runat="server" Text="送出"></asp:Button>

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
                                <%--<asp:Label ID="Label2" runat="server" ForeColor="Red" CssClass="font">	＊(請注意!!)<BR></asp:Label>--%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="HidADID" type="hidden" runat="server" />
        <input id="HID_ZIPCODES1" type="hidden" runat="server" />
    </form>
</body>
</html>
