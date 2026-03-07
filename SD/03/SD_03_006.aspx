<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_03_006.aspx.vb" Inherits="TIMS.SD_03_006" %>

<%@ Register TagPrefix="uc1" TagName="PageControler" Src="../../PageControler.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓學員資料維護</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">

        function GETvalue() {
            document.getElementById('Button10').click();
        }

        function search() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請選擇職類班別!')
                return false;
            }
        }
        function choose_class(num) {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            document.form1.Button2.disabled = true;
            document.getElementById('ImportTable').style.display = 'none';
            document.getElementById('DataGridTable').style.display = 'none';
            document.getElementById('msg').innerHTML = '';

            openClass('../02/SD_02_end.aspx?RWClass=1&RID=' + RID);
        }

        function CheckPrint() {
            var flag = false;
            var MyTable = document.getElementById('DataGrid1');
            var StudentID = '';
            var OCID = document.getElementById('OCIDValue1').value;

            for (var i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = MyTable.rows(i).cells(0).children(0);

                if (MyCheck.checked) {
                    flag = true;
                    if (StudentID == '') {
                        StudentID = '\'' + MyCheck.value + '\'';
                    }
                    else {
                        StudentID += ',\'' + MyCheck.value + '\'';
                    }
                }
            }

            if (document.getElementById('PrintValue').value == '') {
                alert('請勾選要列印的學員!')
                return false;
            }
            else {
                url = '../../SQControl.aspx?&SQ_AutoLogout=true&sys=list&filename=in_class_stud&path=TIMS&'
                window.open(url + 'StudentID=' + document.getElementById('PrintValue').value + '&OCID=' + OCID, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
            }
        }

        function InsertValue(Flag, MyValue) {
            if (Flag) {
                if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\'') == -1) {
                    if (document.getElementById('PrintValue').value == '') {
                        document.getElementById('PrintValue').value = '\'' + MyValue + '\''
                    }
                    else {
                        document.getElementById('PrintValue').value += ',\'' + MyValue + '\''
                    }
                }
            }
            else {
                if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\'') != -1) {
                    if (document.getElementById('PrintValue').value.indexOf(',\'' + MyValue + '\'') != -1) {
                        document.getElementById('PrintValue').value = document.getElementById('PrintValue').value.replace(',\'' + MyValue + '\'', '');
                    }
                    if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\',') != -1) {
                        document.getElementById('PrintValue').value = document.getElementById('PrintValue').value.replace('\'' + MyValue + '\',', '');
                    }
                    if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\'') != -1) {
                        document.getElementById('PrintValue').value = document.getElementById('PrintValue').value.replace('\'' + MyValue + '\'', '');
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;<font color="#990000">結訓學員資料維護</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <table class="table_nw" id="SearchTable" cellspacing="1" width="100%">
                                <tr>
                                    <td width="100" class="bluecol">
                                        訓練機構
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                        <input id="RIDValue" type="hidden" size="1" name="RIDValue" runat="server" />
                                        <input id="Button8" type="button" value="..." name="Button8" runat="server" class="button_b_Mini" />
                                        <asp:Button ID="Button10" Style="display: none" runat="server" Text="Button10"></asp:Button>
                                        <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                            <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                            </asp:Table>
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="100" class="bluecol_need">
                                        職類/班級
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                        <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                        <input id="Button5" type="button" value="..." name="Button5" runat="server" class="button_b_Mini" />
                                        <input id="OCIDValue1" type="hidden" size="1" name="OCIDValue1" runat="server" />
                                        <input id="TMIDValue1" type="hidden" size="1" name="TMIDValue1" runat="server" />
                                        <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                            <asp:Table ID="HistoryTable" runat="server" Width="310">
                                            </asp:Table>
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" class="whitecol" align="center">
                                                                            <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                                        <asp:Button ID="Button2" runat="server" Text="新增" Enabled="False" CssClass="asp_button_S"></asp:Button>

                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table class="table_nw" id="ImportTable" cellspacing="1" width="100%" runat="server">
                                <tr>
                                    <td width="100" class="bluecol">
                                        匯入學員名冊
                                    </td>
                                    <td class="whitecol">
                                        <input id="File1" type="file" name="File1" runat="server" size="50" />
                                        <asp:Button ID="Button7" runat="server" Text="匯入名冊" CssClass="asp_button_S"></asp:Button>(必須為csv格式)
                                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/ClassStudent.zip" ForeColor="#8080FF" CssClass="font">下載整批上載格式檔</asp:HyperLink>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="100" class="bluecol">
                                        匯出學員名冊
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="shiftsort" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1" Selected="True">以代號匯出</asp:ListItem>
                                            <asp:ListItem Value="2">以名稱匯出</asp:ListItem>
                                        </asp:RadioButtonList>
                                        <asp:Button ID="Button6" runat="server" Text="匯出名冊" CssClass="asp_button_S"></asp:Button>
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
                        <td>
                            <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <table class="font" id="Table4" cellspacing="1" cellpadding="1" border="0">
                                            <tr>
                                                <td width="65">
                                                    受訓日期：
                                                </td>
                                                <td width="150">
                                                    <asp:Label ID="DateRound" runat="server" CssClass="font"></asp:Label>
                                                </td>
                                                <td width="40">
                                                    導師：
                                                </td>
                                                <td width="80">
                                                    <asp:Label ID="CTName" runat="server" CssClass="font"></asp:Label>
                                                </td>
                                                <td width="65">
                                                    開班人數：
                                                </td>
                                                <td width="50">
                                                    <asp:Label ID="Tnum" runat="server" CssClass="font"></asp:Label>
                                                </td>
                                                <td width="65">
                                                    學員人數：
                                                </td>
                                                <td>
                                                    <asp:Label ID="StdNum" runat="server" CssClass="font"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                        <font color="red" size="2">*表示為該學員有必填資料未填</font>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowCustomPaging="True" AllowPaging="True" AutoGenerateColumns="False">
                                            <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                            <HeaderStyle CssClass="head_navy" />
                                            <Columns>
                                                <asp:TemplateColumn HeaderText="選取">
                                                    <HeaderStyle Width="25px"></HeaderStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="star1" runat="server" Visible="False">
																	<FONT color="#ff0000">*</FONT></asp:Label>
                                                        <input id="Checkbox2" type="checkbox" runat="server">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Sex" HeaderText="性別"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Birthday" HeaderText="出生日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="學員狀態"></asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="Button3" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                                                        <asp:LinkButton ID="Button9b" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
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
                                    <td align="center">
                                        <asp:Button ID="Button9" runat="server" Width="102px" Text="查詢參訓歷史" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="Button4" runat="server" Text="列印資料卡" CssClass="asp_button_M"></asp:Button>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <input id="PrintValue" type="hidden" size="1" runat="server" />
    </form>
</body>
</html>
