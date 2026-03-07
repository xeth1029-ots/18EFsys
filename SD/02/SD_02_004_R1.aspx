<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_02_004_R1.aspx.vb" Inherits="WDAIIP.SD_02_004_R1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>甄試通知單1</title>
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
        function chall() {
            var mytable = document.getElementById('DataGrid1');
            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheck = mytable.rows[i].cells[0].children[0];
                if (mycheck.disabled == false)
                    mycheck.checked = document.form1.Choose1.checked; //name=Choose1
            }
        }

        function CheckPrint() {
            var MyTable = document.getElementById('DataGrid1');
            if (!MyTable) { return false; }
            var Mailtype = document.getElementById('Mailtype');
            if (!Mailtype) { return false; }
            var chkvalue = document.getElementById('chkvalue');
            if (!chkvalue) { return false; }

            var vExamID = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (vExamID != '') { vExamID += ','; }
                    vExamID += MyTable.rows[i].cells[0].children[1].value;
                }
            }
            if (vExamID == '') {
                alert('請選擇學員');
                return false;
            }
            else {
                var reportFN1 = "Maintest_list2";
                var DistID = document.getElementById('DistID');
                var OCID = document.getElementById('OCID');
                //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=Maintest_list&&path=TIMS&ExamNO1='+ExamID+'&DistID='+document.getElementById('DistID').value+'&OCID1='+document.getElementById('OCID').value+'&Mailtype1='+Mailtype);
                openPrint('../../SQControl.aspx?filename=' + reportFN1 + '&ExamNO1=' + vExamID + '&DistID=' + DistID.value + '&OCID1=' + OCID.value + Mailtype.value + chkvalue.value);
            }
        }
    </script>
    <style type="text/css">
        .auto-style1 { height: 25px; width: 873px; }
        .auto-style2 { color: #333333; padding: 4px; width: 873px; }
        .auto-style3 { width: 873px; }
        .auto-style4 { font-size: 1rem; width: 100%; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <%--
                    <table class="font" id="table2" cellspacing="1" cellpadding="1" border="0">
					    <tr>
						    <td>
							    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
							    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試通知單</asp:Label>
						    </td>
					    </tr>
				    </table>
                    --%>
                    <table id="table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td style="height: 25px">
                                <asp:Label ID="label1" runat="server" CssClass="font"></asp:Label>
                                <input id="DistID" type="hidden" name="DistID" runat="server" />
                                <input id="OCID" type="hidden" name="OCID" runat="server" />
                                <input id="Mailtype" type="hidden" name="Mailtype" runat="server" />
                                <input id="chkvalue" type="hidden" name="chkvalue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="false" AllowCustomPaging="true" AllowPaging="true" PageSize="20" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderTemplate>
                                                <input onclick="chall();" type="checkbox" checked="checked" name="Choose1" />
                                            </HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <input id="checkbox1" type="checkbox" checked="checked" name="student" runat="server" />
                                                <input id="ExamNO" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ExamNO" HeaderText="准考證號碼">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="relenterdate" HeaderText="報名日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="middle"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="列印">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="Button1" runat="server" Text="列印" CommandName="print" CssClass="asp_Export_M"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="准考証重複">
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:Label ID="DoubleExamNo" runat="server"></asp:Label>
                                                <asp:LinkButton ID="UpdateBtn" runat="server" Text="修正准考證號" CommandName="update" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="郵寄類型">
                                            <ItemTemplate>
                                                <asp:CheckBoxList ID="Mailtype3" runat="server" CssClass="font" RepeatDirection="horizontal">
                                                    <asp:ListItem Value="1">印刷品</asp:ListItem>
                                                    <asp:ListItem Value="2">平信</asp:ListItem>
                                                    <asp:ListItem Value="3">限時</asp:ListItem>
                                                    <asp:ListItem Value="4">掛號</asp:ListItem>
                                                    <asp:ListItem Value="5">雙掛號</asp:ListItem>
                                                </asp:CheckBoxList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="false"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="button3" runat="server" Text="群組列印" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="button2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button4" runat="server" Text="自動修正准考試號" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
