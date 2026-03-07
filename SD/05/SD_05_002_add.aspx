<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_002_add.aspx.vb" Inherits="WDAIIP.SD_05_002_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤作業</title>
    <meta http-equiv="Content-Language" content="zh-tw">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        /*.head_navy {  color: White; background-color: #0eabd6 !important; line-height: 30px; text-align: center; border-right:1px solid #FFF; }*/
        .FixedTitleRow { z-index: 10; color: white; position: relative; top: expression(this.offsetParent.scrollTop); background-color: #0eabd6; }
        /*.FixedTitleRow { z-index: 10; color: white; position: relative; top: expression(this.offsetParent.scrollTop); background-color: #2aafc0; }*/
        .FixedTitleColumn { left: expression(this.parentElement.offsetParent.scrollLeft); position: relative; background-color: blue; }
        .FixedDataColumn { left: expression(this.parentElement.offsetParent.parentElement.scrollLeft); position: relative; }
        .DivWidth { display: inline; overflow: auto; width: 600px; cursor: default; position: static; height: 350px; }
        .DivHeight { overflow-y: scroll; display: inline; cursor: default; position: static; height: 350px; }
        .style1 { font-size: 12px; color: #FF0000; line-height: 22px; text-align: center; background-color: #CCD8EE; padding: 2px; height: 26px; }
        .style2 { font-size: 12px; color: Black; line-height: 22px; background-color: #FFFFFF; padding: 2px; height: 26px; }
    </style>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function search() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請選擇職類班別!')
                return false;
            }
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            //'BtnName
            //openClass('../02/SD_02_ch.aspx?special=2&RID=' + RIDValue.value);
            openClass('../02/SD_02_ch.aspx?BtnName=Button5&RID=' + RIDValue.value);
            //document.form1.Hidden1.value = 1;
        }

        function chkdata() {
            var msg = ''
            if (document.form1.LeaveDate.value == '') msg += '請填寫申請日\n'; //判斷是否寫申請日期

            if (document.form1.LeaveDate.value != '' && !checkDate(document.form1.LeaveDate.value)) msg += '申請日日期格式不正確\n';
            var mytable = document.getElementById('DataGrid1');
            var allstudent = 0;

            // 令i = datagrid1表格裡的第一位學員;i< DataGrid1的全部行數;i=i+1
            for (var i = 1; i < mytable.rows.length; i++) {
                var mydrop = mytable.rows[i].cells[2].children[0]; //mydrop = 第i行的[假別]那個欄位
                //if 假別 不等於 '請選擇'
                if (mydrop.selectedIndex != 0) {
                    var all = 0;
                    //判斷DataGrid1裡的第i行的第1節到第12節有無勾選 開始-----
                    for (var j = 7; j < mytable.rows[i].cells.length - 1; j++) {
                        if (j < 19) {
                            var mycheck = mytable.rows[i].cells[j].children[0];
                            if (mycheck) {
                                if (!mycheck.checked) all++; // 若沒有勾選 all =all+1
                            }
                            else {
                                break;
                            }
                        }
                    }
                    if (all == (mytable.rows[i].cells.length - 8)) msg += '請勾選請假的節數(學號' + i + '學員)\n'; //第i行的全部欄位減8會等於從第1節~第12節間的節數
                    //如果第1節~第12節間的欄位沒有勾選 就msg+ =請勾選請假的節數
                }                                                                                           // ----結束
                else {
                    allstudent++;
                }

                var chk = 0;
                //判斷情況和以上相反
                for (var j = 7; j < mytable.rows[i].cells.length - 1; j++) {
                    if (j < 19) {
                        var mycheck = mytable.rows[i].cells[j].children[0];
                        if (mycheck) {
                            if (mycheck.checked) chk++;
                        }
                        else {
                            break;
                        }
                    }
                }
                if (chk != 0 && mydrop.selectedIndex == 0) msg += '請勾選假別(學號' + i + '學員)\n';
                //  如果第1節~第12節間的欄位有勾選 and 假別 等於 '請選擇'  msg+=請勾選假別 
            }
            if (allstudent == mytable.rows.length - 1) msg += '至少要選擇一名學生\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function GetClassTime1(num) {
            var mytable = document.getElementById('DataGrid1');
            for (var j = 7; j < 11; j++) {
                if (mytable.rows[num].cells[3].children[0].checked) {
                    mytable.rows[num].cells[j].children[0].checked = true;
                    if (mytable.rows[num].cells[4].children[0].checked)
                        mytable.rows[num].cells[6].children[0].checked = true;
                }
                else {
                    mytable.rows[num].cells[j].children[0].checked = false;
                    mytable.rows[num].cells[6].children[0].checked = false;
                }
            }
        }
        function GetClassTime2(num) {
            var mytable = document.getElementById('DataGrid1');
            for (var j = 11; j < 15; j++) {
                if (mytable.rows[num].cells[4].children[0].checked) {
                    mytable.rows[num].cells[j].children[0].checked = true;
                    if (mytable.rows[num].cells[3].children[0].checked)
                        mytable.rows[num].cells[6].children[0].checked = true;
                }
                else {
                    mytable.rows[num].cells[j].children[0].checked = false;
                    mytable.rows[num].cells[6].children[0].checked = false;
                }
            }
        }
        function GetClassTime3(num) {
            var mytable = document.getElementById('DataGrid1');
            for (var j = 15; j < 19; j++) {
                if (mytable.rows[num].cells[5].children[0].checked == true) {
                    mytable.rows[num].cells[j].children[0].checked = true;
                }
                else {
                    mytable.rows[num].cells[j].children[0].checked = false;
                }
            }
        }
        function GetClassTime4(num) {
            var mytable = document.getElementById('DataGrid1');
            for (var j = 7; j < 15; j++) {
                if (mytable.rows[num].cells[6].children[0].checked == true) {
                    mytable.rows[num].cells[j].children[0].checked = true;
                    mytable.rows[num].cells[3].children[0].checked = true;
                    mytable.rows[num].cells[4].children[0].checked = true;
                }
                else {
                    mytable.rows[num].cells[j].children[0].checked = false;
                    mytable.rows[num].cells[3].children[0].checked = false;
                    mytable.rows[num].cells[4].children[0].checked = false;
                }
            }
        }
        function GetClassTime5(num) {
            var mytable = document.getElementById('DataGrid1');
            for (var j = 7; j < 11; j++) {
                if (mytable.rows[num].cells[5].children[0].checked == true) {
                    mytable.rows[num].cells[j].children[0].checked = true;
                }
                else {
                    mytable.rows[num].cells[j].children[0].checked = false;
                }
            }
        }
    </script>
</head>
<%--<body onload="FrameLoad();">--%>
<%--上面試原本的寫法，不知道為什麼要那樣(加上那行會讓寬度限制在740px不能調整成100%)--%>
<body>
    <form id="form1" method="post" runat="server">
        <%--<input id="MySessionSet" type="hidden" name="MySessionSet" runat="server">	--%>
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員出缺勤作業</font> </td>
					</tr>
				</table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="55%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" runat="server" />
                                <%--<input id="Button4" type="button" value="..." runat="server">
							<input id="Hidden1" type="hidden" runat="server">--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button3" type="button" value="..." runat="server">
                                <input id="TMIDValue1" type="hidden" runat="server" />
                                <input id="OCIDValue1" type="hidden" runat="server" />
                                <span id="spanButton5" runat="server" style="display: none">
                                    <asp:Button ID="Button5" runat="server" Text="顯示學員"></asp:Button>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">點名日期 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="LeaveDate" runat="server" Width="15%"></asp:TextBox>
                                <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('LeaveDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                <asp:TextBox ID="StDate" runat="server" Width="33px" Visible="False"></asp:TextBox>
                                <asp:TextBox ID="FtDate" runat="server" Width="33px" Visible="False"></asp:TextBox>
                            </td>
                        </tr>

                        <tr>
                            <td class="bluecol">依分頁</td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="ddl_sSearch2" runat="server" AutoPostBack="True"></asp:DropDownList>
                            </td>
                        </tr>

                    </table>
                    <div>
                        <table id="StudentTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>&nbsp; </td>
                            </tr>
                            <tr>
                                <td>
                                    <div id="scrollDiv" runat="server">
                                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False"   CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="假別">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:DropDownList ID="LeaveID" runat="server">
                                                        </asp:DropDownList>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="1-4節">
                                                    <HeaderStyle BackColor="gold" Width="4%" ForeColor="black"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C14" type="checkbox" runat="server" name="C14">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="5-8節">
                                                    <HeaderStyle BackColor="gold" Width="4%" ForeColor="black"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C58" type="checkbox" runat="server" name="C58">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="9-12節">
                                                    <HeaderStyle BackColor="gold" Width="5%" ForeColor="black"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C912" type="checkbox" runat="server" name="C912">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="1-8節">
                                                    <HeaderStyle BackColor="gold" Width="4%" ForeColor="black"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C18" type="checkbox" runat="server" name="C18">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節1">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C1" type="checkbox" runat="server" name="C1">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節2">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C2" type="checkbox" runat="server" name="C2">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節3">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C3" type="checkbox" runat="server" name="C3">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節4">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C4" type="checkbox" runat="server" name="C4">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節5">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C5" type="checkbox" runat="server" name="C5">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節6">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C6" type="checkbox" runat="server" name="C6">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節7">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C7" type="checkbox" runat="server" name="C7">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節8">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C8" type="checkbox" runat="server" name="C8">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節9">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C9" type="checkbox" runat="server" name="C9">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節10">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C10" type="checkbox" runat="server" name="C10">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節11">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C11" type="checkbox" runat="server" name="C11">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="節12">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="C12" type="checkbox" runat="server" name="C12">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn Visible="False" DataField="LeaveID" HeaderText="LeaveID"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="total" HeaderText="備註" HeaderStyle-Width="12%"></asp:BoundColumn>
                                                <asp:BoundColumn Visible="False" DataField="TPeriod" HeaderText="訓練時段"></asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="不列入&lt;br/&gt;缺曠課" HeaderStyle-Width="5%">
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <input id="TurnoutIgnore" type="checkbox" runat="server" name="TurnoutIgnore" />
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <input id="Button2" type="button" value="回上一頁" name="Button2" runat="server" class="asp_button_M" />
                    </p>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </p>
                </td>
            </tr>
        </table>
        <input id="LeaveDateHidden" type="hidden" runat="server" />
        <asp:HiddenField ID="HidThours" runat="server" />
        <asp:HiddenField ID="HidItemVar1" runat="server" />
        <asp:HiddenField ID="HidItemVar2" runat="server" />
        <asp:HiddenField ID="HidOCID1" runat="server" />
    </form>
</body>
</html>
