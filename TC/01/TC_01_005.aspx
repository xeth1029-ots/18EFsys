<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_005.aspx.vb" Inherits="WDAIIP.TC_01_005" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>課程資料設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <%--<script type="text/javascript" src="../../js/buttonAuth.js.aspx" charset="UTF-8"></script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function aloader2on() {
            var construction2 = document.getElementById("construction2");
            var form1 = document.getElementById("form1");
            //form1.style.display = "none";              //不顯示
            //construction2.style.display = "block";     //顯示     //(20180907 由於此遮罩屬於TIMS功能，因此先將此遮罩拿掉)
        }

        function aloader2off() {
            var construction2 = document.getElementById("construction2");
            //construction2.style.display = "none";        //不顯示
        }

        function but_edit(courid) {
            location.href = 'TC_01_005_add.aspx?courid=' + courid + '&ProcessType=Update';
        }

        function but_del(courid, ch_courid, is_parent, id) {
            if (is_parent) {
                alert("此為主課程,且尚有子課程,不可刪除!!");
                return;
            }
            if (ch_courid == 0) {
                if (window.confirm("此動作會刪除課程資料，是否確定刪除?")) {
                    location.href = 'TC_01_005_del.aspx?courid=' + courid + '&ID=' + id;
                }
            }
            else {
                alert('此課程已有排課資料，不可以刪除!!');
            }
        }

        function ShowFrame(arg) {
            var FrameObj = document.getElementById('FrameObj');
            var HistoryRID = document.getElementById('HistoryRID');
            var HistoryList2 = document.getElementById('HistoryList2');
            FrameObj.height = HistoryRID.rows.length * 20;
            FrameObj.style.display = HistoryList2.style.display;
        }
    </script>
    <style type="text/css">
        .auto-style1 {
            color: Black;
            text-align: right;
            padding: 4px 6px;
            background-color: #f1f9fc;
            border-right: 3px solid #49cbef;
            width: 66px;
        }

        .auto-style2 {
            color: #FF0000;
            text-align: right;
            padding: 4px 6px;
            background-color: #f1f9fc;
            border-right: 3px solid #49cbef;
            width: 66px;
        }

        .auto-style3 {
            color: Black;
            text-align: right;
            padding: 4px 6px;
            background-color: #f1f9fc;
            border-right: 3px solid #49cbef;
            width: 64px;
        }
    </style>
</head>
<body>
    <%--  <div id="construction2" onclick="aloader2off();">
        <table width="100%" height="100%">
            <tr>
                <td align="center" valign="middle">
                    <img id="construction2-img" src="../../images/icon_construction-a.gif" alt="系統正在處理您的需求 請稍候.."></td>
            </tr>
        </table>
    </div>--%>
    <form id="form1" method="post" runat="server">
        <%--
        <table class="font" width="740">
		   <tr>
			   <td class="font">
				   <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                   <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;開班資料設定&gt;&gt;課程資料設定</asp:Label>
                   <input id="check_del" size="6" type="hidden" name="check_del" runat="server">
				   <input id="check_mod" size="6" type="hidden" name="check_mod" runat="server">
			   </td>
		   </tr>
	    </table>
        --%>
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;開班資料設定&gt;&gt;課程資料設定</asp:Label>
                    <input id="check_del" size="6" type="hidden" name="check_del" runat="server">
                    <input id="check_mod" size="6" type="hidden" name="check_mod" runat="server">
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">訓練機構</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                    <input id="Org" value="..." type="button" name="Org" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                    <span style="z-index: 1; position: absolute; display: none" id="HistoryList2">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                    <iframe style="position: absolute; display: none" id="FrameObj" height="52" frameborder="0" width="310" scrolling="no"></iframe>
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">課程代碼 </td>
                <td class="whitecol" style="width: 30%">
                    <asp:TextBox ID="TB_CourseID" runat="server" Width="40%" MaxLength="8"></asp:TextBox></td>
                <td class="bluecol" style="width: 20%">課程名稱 </td>
                <td class="whitecol" style="width: 30%">
                    <asp:TextBox ID="TB_CourseName" runat="server" Width="60%" MaxLength="30"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">學/術科 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Classification1_List" runat="server">
                        <asp:ListItem Value="0">全部</asp:ListItem>
                        <asp:ListItem Value="1">學科</asp:ListItem>
                        <asp:ListItem Value="2">術科</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td class="bluecol">共同/一般/專業 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Classification2_List" runat="server">
                        <asp:ListItem Value="3">請選擇</asp:ListItem>
                        <asp:ListItem Value="0">共同</asp:ListItem>
                        <asp:ListItem Value="1">一般</asp:ListItem>
                        <asp:ListItem Value="2">專業</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">是否有效 </td>
                <td class="whitecol">
                    <asp:CheckBox ID="CB_Valid" runat="server" CssClass="font"></asp:CheckBox></td>
                <td class="bluecol">訓練職類 </td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox><input id="career" onclick="openTrain(document.getElementById('trainValue').value);" value="..." type="button" name="career" runat="server" class="button_b_Mini"><input id="trainValue" type="hidden" name="trainValue" runat="server"></td>
            </tr>
            <tr>
                <td class="bluecol">含排課匯入用代碼 </td>
                <td colspan="3" class="whitecol">
                    <asp:CheckBox ID="cb_CourID" runat="server" CssClass="font"></asp:CheckBox></td>
            </tr>
            <tr>
                <td class="bluecol">隸屬班級 </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="Classid" runat="server" onfocus="this.blur()" Columns="46" Width="40%"></asp:TextBox><input id="Button3" onclick="javascript: wopen('TC_01_005_Classid.aspx', '班級代碼', 1000, 630, 1)" value="選擇" type="button" name="choice_button" runat="server" class="asp_button_M">
                    <asp:Button ID="Button2" runat="server" CausesValidation="False" Text="清除" CssClass="asp_button_S"></asp:Button><input style="width: 32px; height: 22px" id="Classid_Hid" type="hidden" name="CLSID_Hid" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>

            <%--
            <tr>
                <td bgColor="#cccc66" style="HEIGHT: 19px">&nbsp;&nbsp;&nbsp;年度</td>
                <td colSpan="3" style="HEIGHT: 19px"><asp:dropdownlist id="DrpYear" runat="server"></asp:dropdownlist></td>
		    </tr>
            --%>
        </table>
        <table width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_M" AuthType="QRY"></asp:Button>&nbsp;
				<asp:Button ID="bt_add" Text="新增" runat="server" CssClass="asp_button_M" AuthType="ADD"></asp:Button>&nbsp;
				<asp:Button ID="print" runat="server" Text="列印-課程代碼表" CssClass="asp_Export_M" AuthType="PRT"></asp:Button>&nbsp;
				<asp:Button ID="BTN_IMP1" runat="server" Text="匯入" CssClass="asp_button_M" AuthType="ADD"></asp:Button>&nbsp;
				<asp:Button ID="btnExport" runat="server" Text="匯出" CssClass="asp_Export_M" AuthType="QRY"></asp:Button>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <%--<asp:Panel ID="Panel1" runat="server" Width="100%" Visible="False" HorizontalAlign="Center"></asp:Panel>--%>
        <br />
        <%--<asp:Panel ID="Panel" runat="server" Width="100%" Visible="False"></asp:Panel>--%>
        <table id="tb_DG_Course" runat="server" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DG_Course" runat="server" Width="100%" CssClass="font" Visible="False" AllowCustomPaging="True" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CourseID" HeaderText="課程代碼">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CourseName" HeaderText="課程名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Hours" HeaderText="小時數">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Classification1" HeaderText="學/術科">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Classification2" HeaderText="共同/一般/專業">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="主課程">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="Button4" runat="server" Text="修改" CommandName="edit" AuthType="UPD" class="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button5" runat="server" Text="刪除" CommandName="del" AuthType="DEL" class="asp_button_M"></asp:Button>
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
        </table>

    </form>
</body>
</html>
