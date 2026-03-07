<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_008.aspx.vb" Inherits="WDAIIP.SD_04_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_04_008</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        function CheckData() {
            var msg = '';

            if (document.getElementById('TechID').value == '') msg += '請選擇師資\n';
            if (document.getElementById('Start_Date').value == '') msg += '請選擇查詢啟日\n'
            else if (!checkDate(document.getElementById('Start_Date').value)) msg += '查詢啟日不是正確的日期格式\n';
            if (document.getElementById('End_Date').value == '') msg += '請選擇查詢迄日\n'
            else if (!checkDate(document.getElementById('End_Date').value)) msg += '查詢啟日不是正確的日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function CheckData2() {
            var msg = '';

            if (document.getElementById('Start_Date2').value != '')
                if (!checkDate(document.getElementById('Start_Date2').value)) msg += '查詢啟日不是正確的日期格式\n';
            if (document.getElementById('End_Date2').value != '')
                if (!checkDate(document.getElementById('End_Date2').value)) msg += '查詢啟日不是正確的日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function ChangeItem(num) {
            if (num == 1) {
                moveOption(document.getElementById('ListBox1'), document.getElementById('ListBox2'));
            }
            else {
                moveOption(document.getElementById('ListBox2'), document.getElementById('ListBox1'));
            }
            document.getElementById('TechID').value = '';

            MyList = document.getElementById('ListBox2');
            for (var i = 0; i < MyList.options.length; i++) {
                var Index = MyList.options[i].value;
                if (document.getElementById('TechID').value == '')
                    document.getElementById('TechID').value = Index;
                else
                    document.getElementById('TechID').value += ',' + Index;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="600" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <font face="新細明體">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;<font color="#990000">老師衝堂查詢</font>(此功能必須填入師資正確的身分證號碼才可正確判斷結果)</font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Page1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="SD_TD1" width="100">
                                            <font face="新細明體">&nbsp;&nbsp;&nbsp; 訓練機構</font>
                                        </td>
                                        <td class="SD_TD2" colspan="3">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox><input id="Button3" type="button" value="..." name="Button3" runat="server"><input id="RIDValue" type="hidden" runat="server"><br>
                                            <span id="HistoryList1" style="display: none; position: absolute">
                                                <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                                </asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="SD_TD1">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>講師姓名</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <font face="新細明體">
                                                <asp:TextBox ID="TeachName" runat="server"></asp:TextBox></font>
                                        </td>
                                        <td class="SD_TD1" width="100">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>身分證號碼</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:TextBox ID="IDNO" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="SD_TD1">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>講師代碼</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <font face="新細明體">
                                                <asp:TextBox ID="TeacherID" runat="server"></asp:TextBox></font>
                                        </td>
                                        <td class="SD_TD1">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>職稱</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <font face="新細明體">
                                                <asp:DropDownList ID="IVID" runat="server">
                                                </asp:DropDownList>
                                            </font>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="SD_TD1">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>主要職類</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <font face="新細明體">
                                                <asp:TextBox ID="TB_career_id" runat="server"></asp:TextBox><input type="button" value="..." onclick="openTrain(document.getElementById('trainValue').value);"><input id="trainValue" type="hidden" runat="server"></font>
                                        </td>
                                        <td class="SD_TD1">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>任職狀況</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:RadioButtonList ID="WorkStatus" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                <asp:ListItem Value="1" Selected="True">在職</asp:ListItem>
                                                <asp:ListItem Value="2">離職</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="SD_TD1">&nbsp;&nbsp;&nbsp; 內外聘
                                        </td>
                                        <td class="SD_TD2">
                                            <font face="新細明體">
                                                <asp:DropDownList ID="KindEngage" runat="server">
                                                    <asp:ListItem Value="不區分">不區分</asp:ListItem>
                                                    <asp:ListItem Value="1">內聘</asp:ListItem>
                                                    <asp:ListItem Value="2">外聘</asp:ListItem>
                                                </asp:DropDownList>
                                            </font>
                                        </td>
                                        <td class="SD_TD1">
                                            <font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp; </font>師資別</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:DropDownList ID="KindID" runat="server">
                                                <asp:ListItem Value="請先選擇內外聘">請先選擇內外聘</asp:ListItem>
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="4">
                                            <font face="新細明體">
                                                <asp:Button ID="Button1" runat="server" Text="查詢老師"></asp:Button><asp:Button ID="Button6" runat="server" Text="查詢結果"></asp:Button></font>
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <font face="新細明體">檢查日期範圍：
                                            <asp:TextBox ID="Start_Date" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= Start_Date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
                                            <asp:TextBox ID="End_Date" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= End_Date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">

                                            <table class="font" id="Table4" cellspacing="1" cellpadding="1" border="0">
                                                <tr>
                                                    <td>所有師資
                                                    </td>
                                                    <td></td>
                                                    <td>選取的師資
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <asp:ListBox ID="ListBox1" runat="server" Width="200px" Rows="8" SelectionMode="Multiple"></asp:ListBox>
                                                    </td>
                                                    <td>
                                                        <input id="Button4" type="button" value=">>" name="Button4" runat="server"><br>
                                                        <br>
                                                        <input id="Button5" type="button" value="<<" name="Button5" runat="server">
                                                    </td>
                                                    <td>
                                                        <asp:ListBox ID="ListBox2" runat="server" Width="200px" Rows="8" SelectionMode="Multiple"></asp:ListBox>
                                                    </td>
                                                </tr>
                                            </table>

                                            <input id="TechID" type="hidden" runat="server">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <font face="新細明體">
                                                <asp:Button ID="Button2" runat="server" Text="送出"></asp:Button></font>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Page2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="SD_TD1" width="100">&nbsp;&nbsp;&nbsp; 訓練機構
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:TextBox ID="center2" runat="server" onfocus="this.blur()"></asp:TextBox><input id="Button9" type="button" value="..." name="Button3" runat="server"><input id="RIDValue2" type="hidden" name="RIDValue2" runat="server"><br>
                                            <span id="HistoryList2" style="display: none; position: absolute">
                                                <asp:Table ID="HistoryRID2" runat="server" Width="310px">
                                                </asp:Table>
                                            </span>
                                        </td>
                                        <td class="SD_TD1" width="100">&nbsp;&nbsp;&nbsp; 講師代碼
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:TextBox ID="TeacherID2" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="SD_TD1" width="100">
                                            <font face="新細明體">&nbsp;&nbsp;&nbsp; 講師姓名</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:TextBox ID="TeachCName2" runat="server"></asp:TextBox>
                                        </td>
                                        <td class="SD_TD1" width="100">
                                            <font face="新細明體">&nbsp;&nbsp;&nbsp; 身分證號碼</font>
                                        </td>
                                        <td class="SD_TD2">
                                            <asp:TextBox ID="IDNO2" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="SD_TD1">
                                            <font face="新細明體">&nbsp;&nbsp;&nbsp; 查詢日期範圍</font>
                                        </td>
                                        <td class="SD_TD2" colspan="3">
                                            <font face="新細明體">
                                                <asp:TextBox ID="Start_Date2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= Start_Date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                                ～
                                            <asp:TextBox ID="End_Date2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= End_Date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="4">
                                            <font face="新細明體">
                                                <asp:Button ID="Button7" runat="server" Text="查詢"></asp:Button><asp:Button ID="Button8" runat="server" Text="回上一頁"></asp:Button></font>
                                        </td>
                                    </tr>
                                </table>
                                <table id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AllowPaging="True" CssClass="font" AutoGenerateColumns="False">
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle Width="25px"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="DupDate" HeaderText="衝堂日期" DataFormatString="{0:d}">
                                                        <HeaderStyle Width="60px"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼">
                                                        <HeaderStyle Width="60px"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TeachCName" HeaderText="講師姓名">
                                                        <HeaderStyle Width="60px"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="DupPart" HeaderText="節次">
                                                        <HeaderStyle Width="25px"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="DupDesc" HeaderText="說明"></asp:BoundColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                                        </td>
                                    </tr>
                                </table>
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
