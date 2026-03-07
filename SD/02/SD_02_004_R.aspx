<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_02_004_R.aspx.vb" Inherits="WDAIIP.SD_02_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>甄試通知單</title>
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
    <script type="text/javascript">
        function checkselect(obj) {
            document.form1.M6.value = 1;
            if (document.getElementById(obj + '_' + 0).checked)
            { document.form1.M1.value = 1; } else { document.form1.M1.value = 0; }
            if (document.getElementById(obj + '_' + 1).checked)
            { document.form1.M2.value = 1; } else { document.form1.M2.value = 0; }
            if (document.getElementById(obj + '_' + 2).checked)
            { document.form1.M3.value = 1; } else { document.form1.M3.value = 0; }
            if (document.getElementById(obj + '_' + 3).checked)
            { document.form1.M4.value = 1; } else { document.form1.M4.value = 0; }
            if (document.getElementById(obj + '_' + 4).checked)
            { document.form1.M5.value = 1; } else { document.form1.M5.value = 0; }
        }

        function GETvalue() {
            document.getElementById('Button7').click();
        }

        /*
		function RedictPage(CommandArgument, ID) {
		var M1 = document.getElementById('M1').value;
		var M2 = document.getElementById('M2').value;
		var M3 = document.getElementById('M3').value;
		var M4 = document.getElementById('M4').value;
		var M5 = document.getElementById('M5').value;
		getCheckBoxListChecked();
		location.href = 'SD_02_004_R1.aspx' + CommandArgument + '&ID=' + ID + '&Mailtype1=' + M1 + '&Mailtype2=' + M2 + '&Mailtype3=' + M3 + '&Mailtype4=' + M4 + '&Mailtype5=' + M5 + document.getElementById('chkvalue').value;
		}

		function print_rpt(SMpath, CommandArgument, DistID) {
		var M1 = document.getElementById('M1').value;
		var M2 = document.getElementById('M2').value;
		var M3 = document.getElementById('M3').value;
		var M4 = document.getElementById('M4').value;
		var M5 = document.getElementById('M5').value;
		getCheckBoxListChecked();
		openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=Maintest_list_org&path=' + SMpath + CommandArgument + '&DistID=' + DistID + '&Mailtype1=' + M1 + '&Mailtype2=' + M2 + '&Mailtype3=' + M3 + '&Mailtype4=' + M4 + '&Mailtype5=' + M5 + document.getElementById('chkvalue').value);
		}

		function search() {
		//if(document.form1.OCIDValue1.value==''){
		//	alert('請選擇職類班別!')
		//return false;
		//}			
		//	if (document.form1.start_date.value=='' || document.form1.end_date.value==''){
		//	window.alert('請選擇開訓日期範圍');
		//	return false;
		//}				
		}

		function chall(num) {
		if (num == 1) {
		document.form1.OCID_Grade.checked = document.form1.Choose1.checked
		for (var i = 0; i < document.form1.OCID_Grade.length; i++)
		document.form1.OCID_Grade[i].checked = document.form1.Choose1.checked
		}
		else {
		document.form1.OCID_Sort.checked = document.form1.Choose2.checked
		for (var i = 0; i < document.form1.OCID_Grade.length; i++)
		document.form1.OCID_Sort[i].checked = document.form1.Choose2.checked
		}
		}
		*/

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        //限定textbox的欄位長度
        function checkTextLength(obj, xlong) {
            var maxlength = new Number(xlong);
            if (obj.value.length > maxlength) {
                obj.value = obj.value.substring(0, maxlength);
            }
        }

        /*
		function getCheckBoxListChecked() {
		var elementref = document.getElementById('cblist_info');
		var checkBoxArray = elementref.getElementsByTagName('input');
		var checkedValues = '';
		for (var i = 0; i < checkBoxArray.length; i++) {
		var checkBoxRef = checkBoxArray[i];
		if (checkBoxRef.checked == true) {
		if (checkedValues.length > 0) checkedValues += '';
		checkedValues += '&chk' + String(i + 1) + '=1';
		}
		else {
		if (checkedValues.length > 0) checkedValues += '';
		checkedValues += '&chk' + String(i + 1) + '=0';
		}
		}
		document.getElementById('chkvalue').value = checkedValues;
		}
		*/
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試通知單</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button6" Style="display: none" runat="server" Text="Button5" CssClass="asp_button_S"></asp:Button>
                                <span id="HistoryList2" style="z-index: 100; position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班級 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class();" type="button" value="..." class="asp_button_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="z-index: 102; position: absolute; display: none; left: 28%">
                                    <asp:Table ID="Historytable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">開訓日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                ～
                                <asp:TextBox ID="end_date" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" id="td_3" runat="server">列印課程資訊 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblist_info" runat="server" Width="400px" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1" Selected="true">計劃</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="true">班名</asp:ListItem>
                                    <asp:ListItem Value="3" Selected="true">年度</asp:ListItem>
                                    <asp:ListItem Value="4" Selected="true">期別</asp:ListItem>
                                    <asp:ListItem Value="5" Selected="true">准考證號</asp:ListItem>
                                </asp:CheckBoxList>
                                <input id="chkvalue" type="hidden" name="chkvalue" runat="server" />
                                <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button5"></asp:Button>
                            </td>
                        </tr>
                        <tr id="Trwork2013a" runat="server">
                            <td class="bluecol">就服單位協助報名 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                                <%--
							    <asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
								    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
								    <asp:ListItem Value="CH4">一般推介單</asp:ListItem>
								    <asp:ListItem Value="EPW">免試推介單</asp:ListItem>
								    <asp:ListItem Value="EP2P">專案核定報名</asp:ListItem>
							    </asp:RadioButtonList>
                                --%>
                            </td>
                        </tr>
                    </table>
                    <table style="width: 100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" Font-Size="X-Small" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="Button4" runat="server" Text="設定通知單內容" Visible="False" CssClass="asp_button_M"></asp:Button>&nbsp;
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="table11" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol">甄試通知單內容 </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <asp:TextBox ID="Item_Note" onblur="checkTextLength(this,512)" onkeyup="checkTextLength(this,512)" onpropertychange="checkTextLength(this,512);" Width="100%" Rows="3" Columns="80" TextMode="MultiLine" runat="server" Height="160px" onChange="checkTextLength(this,512)"></asp:TextBox></td>
                        </tr>
                    </table>
                    <table style="width: 100%">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="true" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別" HeaderStyle-Width="15%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="total" HeaderText="人數" HeaderStyle-Width="5%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="郵寄類別" HeaderStyle-Width="30%">
                                            <ItemTemplate>
                                                <asp:CheckBoxList ID="Mailtype1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                    <asp:ListItem Value="1">印刷品</asp:ListItem>
                                                    <asp:ListItem Value="2">平信</asp:ListItem>
                                                    <asp:ListItem Value="3">限時</asp:ListItem>
                                                    <asp:ListItem Value="4">掛號</asp:ListItem>
                                                    <asp:ListItem Value="5">雙掛號</asp:ListItem>
                                                </asp:CheckBoxList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="列印" HeaderStyle-Width="20%">
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="Button2" runat="server" Text="列印全部" CommandName="all" CssClass="asp_Export_M"></asp:LinkButton>
                                                <asp:LinkButton ID="Button3" runat="server" Text="個別列印" CommandName="only" CssClass="asp_Export_M"></asp:LinkButton>
                                                <asp:LinkButton ID="btnExport1" runat="server" Text="匯出全部" CommandName="exp1" CssClass="asp_Export_M"></asp:LinkButton>
                                                <input type="hidden" id="hidOCID" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="備註" HeaderStyle-Width="25%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="LPSMEMO1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--
                                        <asp:BoundColumn Visible="False" DataField="CyclType" HeaderText="CyclType"></asp:BoundColumn>
									    <asp:BoundColumn Visible="False" DataField="LevelType" HeaderText="LevelType"></asp:BoundColumn>
									    <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
									    <asp:BoundColumn Visible="False" DataField="PlanID" HeaderText="PlanID"></asp:BoundColumn>
                                        --%>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td style="height: 27px" align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button5" runat="server" Text="儲存" Visible="False" CssClass="asp_button_M"></asp:Button><br />
                                <input id="M1" type="hidden" name="M1" runat="server" />
                                <input id="M2" type="hidden" name="M2" runat="server" />
                                <input id="M3" type="hidden" name="M3" runat="server" />
                                <input id="M4" type="hidden" name="M4" runat="server" />
                                <input id="M5" type="hidden" name="M5" runat="server" />
                                <input id="M6" type="hidden" name="M6" runat="server" />
                                <br />
                                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
