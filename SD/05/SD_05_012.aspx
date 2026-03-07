<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_012.aspx.vb" Inherits="WDAIIP.SD_05_012" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級結訓作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function Check_Data() {
            var MyTable = document.getElementById('DataGrid1');
            //var HidTPlanID = document.getElementById('HidTPlanID');
            var CloseClassName = '';     //搜集close
            var OpenClassName = '';      //搜集open
            var cst_Checkbox1 = 0;       // 勾選
            var cst_seq = 1;             // 序號
            var cst_TrainName = 2;       //訓練職類
            var cst_ClassID = 3;         //班級代碼
            var cst_ClassCName = 4;      //班級名稱
            var cst_STDate = 5;          //開訓日期
            var cst_FTDate = 6;          //結訓日期
            var cst_StudentCount = 7;    //學員人數
            var cst_StudentClose = 8;    //結訓人數
            var cst_nodata = 9;          // 未填資料
            var cst_CanCloseResult = 10; // 開放班級結訓理由
            for (i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = MyTable.rows[i].cells[0].children[0]; //Checkbox1
                var MyValue = MyTable.rows[i].cells[0].children[1]; //ChangeFlag
                if (MyValue.value == '1') {
                    if (MyCheck.checked) {
                        CloseClassName += MyTable.rows[i].cells[cst_ClassCName].innerHTML + '\n';
                    }
                    else {
                        OpenClassName += MyTable.rows[i].cells[cst_ClassCName].innerHTML + '\n';
                        //num
                        /*
						if (HidTPlanID.value != '' && MyTable.rows[i].cells(cst_OpneResult).children(0).value == '') {
						msg += '請輸入 開放班級結訓理由 (第' + i + '行:' + MyTable.rows[i].cells(cst_ClassCName).innerHTML + ')\n';
						}
						*/
                    }
                }
            }
            var msg = '';
            if (CloseClassName != '')
                msg += '系統將結訓以下班級\n' + CloseClassName + '\n';
            if (OpenClassName != '')
                msg += '系統將解除結訓以下班級\n' + OpenClassName + '\n';
            if (msg == '') {
                alert('請先選擇班級\n');
                return false;
            }
            else {
                msg += '您確定要執行以上動作?';
                return confirm(msg);
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;班級結訓作業</asp:Label>
                </td>
            </tr>
        </table>
        <div>
            <table class="table_sch" id="Table2" cellpadding="1" cellspacing="1">
                <tr>
                    <td class="bluecol" style="width: 15%">訓練機構</td>
                    <td class="whitecol" style="width: 35%">
                        <asp:TextBox ID="center" runat="server" Width="70%" onfocus="this.blur()"></asp:TextBox>
                        <input id="Button7" type="button" value="..." name="Button7" runat="server" class="button_b_Mini">
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                        <asp:Button ID="Button3" Style="display: none" runat="server"></asp:Button>
                        <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                        </span>
                    </td>
                    <td class="bluecol" style="width: 15%">班別代碼</td>
                    <td class="whitecol" style="width: 35%">
                        <asp:TextBox ID="ClassID" runat="server" Columns="15" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">班別名稱</td>
                    <td class="whitecol">
                        <asp:TextBox ID="ClassCName" runat="server" Width="60%"></asp:TextBox></td>
                    <td class="bluecol">期別</td>
                    <td class="whitecol">
                        <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">班級範圍</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="ClassRound" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatLayout="Flow">
                            <asp:ListItem Value="已結訓" Selected="True">已結訓</asp:ListItem>
                            <asp:ListItem Value="未結訓">未結訓</asp:ListItem>
                            <asp:ListItem Value="全部">全部</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol" colspan="4">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </div>
        <div>
            <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                <tr>
                    <td>勾取表示已經結訓，尚未勾取的班級表示尚未結訓&nbsp; <font color="#ff0000" size="2">*表示為該班有必填資料未填，無法執行班級結訓動作</font><br>
                        如該班次尚未超過結訓日期不得執行勾選結訓作業。<br />
                        &nbsp;<asp:Label ID="labMsg28" runat="server" Text="Label"><font color="#ff0000">*開放可做結訓：</font><br>
                                1.開放未填意見調查表可做班級結訓，少數學員因特殊原因未填寫意見調查表，經分署審核，可開放該班進行結訓作業。<br>
                                2.已點選班級結訓之班級，無法使用此功能。
                        </asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn>
                                    <HeaderStyle Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                    <ItemTemplate>
                                        <input id="Checkbox1" type="checkbox" runat="server">
                                        <input id="ChangeFlag" type="hidden" value="0" runat="server">
                                        <asp:Label ID="star" Visible="False" runat="server"><FONT color="#ff0000">*</FONT></asp:Label>
                                        <input id="HidCanClose" type="hidden" runat="server">
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn HeaderText="序號">
                                    <HeaderStyle Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="TrainName" SortExpression="TrainID" HeaderText="訓練職類">
                                    <HeaderStyle ForeColor="#00ffff" Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="ClassID" SortExpression="ClassID" HeaderText="班級代碼">
                                    <HeaderStyle ForeColor="#00ffff" Width="8%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="CLASSCNAME2" SortExpression="ClassCName" HeaderText="班級名稱">
                                    <HeaderStyle ForeColor="#00ffff" Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                    <HeaderStyle ForeColor="#00ffff" Width="8%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="FTDate" SortExpression="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                    <HeaderStyle ForeColor="#00ffff" Width="8%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="StudentCount" HeaderText="學員人數">
                                    <HeaderStyle Width="8%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="StudentClose" HeaderText="結訓人數">
                                    <HeaderStyle Width="8%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="未填資料">
                                    <HeaderStyle Width="10%" />
                                </asp:TemplateColumn>
                                <asp:TemplateColumn>
                                    <HeaderStyle Width="20%" HorizontalAlign="Justify" />
                                    <HeaderTemplate>開放可做結訓</HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:Button ID="BtnCanClose" runat="server" Text="開放" CommandName="OpenCmd" CssClass="asp_button_M" />理由:<br>
                                        <asp:TextBox ID="CanCloseResult" runat="server" Width="250px" MaxLength="150" TextMode="MultiLine"></asp:TextBox><br>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <%-- <asp:TemplateColumn HeaderText="開放班級結訓理由*">
                                            <HeaderStyle Width="109px" />
                                            <ItemTemplate><asp:TextBox ID="OpenResult" runat="server" Width="100px" MaxLength="150" TextMode="MultiLine"></asp:TextBox><br></ItemTemplate>
                                            </asp:TemplateColumn>--%>
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
                    <td align="center" class="whitecol">
                        <asp:Button ID="Button2" runat="server" Text="班級結訓" CssClass="asp_button_M"></asp:Button></td>
                </tr>
            </table>
        </div>
        <div>
            <table cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center" class="whitecol">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>
        <input id="HidTPlanID" type="hidden" name="HidTPlanID" runat="server">
    </form>
</body>
</html>
