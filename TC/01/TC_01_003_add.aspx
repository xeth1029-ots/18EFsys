<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_003_add.aspx.vb" Inherits="WDAIIP.TC_01_003_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班別代碼設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function wopen(url, name, width, height, k) {
            LeftPosition = (screen.width) ? (screen.width - width) / 2 : 0;
            TopPosition = (screen.availHeight) ? (screen.availHeight - height - 28) / 2 : 0;
            window.open(url, name, 'top=' + TopPosition + ',left=' + LeftPosition + ',width=' + width + ',height=' + height + ',resizable=0,scrollbars=' + k + ',status=0');
        }

        //function window_onload() {
        //	if (document.getElementById("BypassCheck")) {
        //		if (window.confirm('新增班別代碼重複!!')) {
        //			document.getElementById('bt_save').click();
        //		} else {
        //		document.getElementById('BypassCheck').value='0';
        //		}
        //	}
        //	}
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
                            首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000"> 班別代碼設定</font>
							</asp:Label>
							<font color="#990000">-
								<asp:Label ID="lblProecessType" runat="server"></asp:Label></font> <font color="#000000">( <font color="#ff0000">*</font>為必填欄位 )</font>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>--%>
        <asp:Label ID="lblProecessType" runat="server" Visible="false"></asp:Label>
        <table class="table_nw" id="Table1" runat="server" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" style="width: 20%">班別代碼
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="TB_classid" runat="server" Columns="10" MaxLength="4"></asp:TextBox>
                    <input id="hTB_classid" type="hidden" name="hTB_classid" runat="server" />
                    <asp:RegularExpressionValidator ID="seqno" runat="server" ErrorMessage="班別代碼請填寫任意四碼英數字" Display="None" ControlToValidate="TB_classid" ValidationExpression="[0-9A-Za-z]{4}"></asp:RegularExpressionValidator>
                    <asp:RequiredFieldValidator ID="TB_seqno" runat="server" ErrorMessage="請輸入班別代碼" Display="None" ControlToValidate="TB_classid"></asp:RequiredFieldValidator>(可輸入四碼英數字)
                </td>
            </tr>
            <tr id="tr_cb_USTUDENTID" runat="server">
                <td class="bluecol">修改學員學號</td>
                <td colspan="3" class="whitecol"><asp:CheckBox ID="cb_USTUDENTID" runat="server" />修改學員學號</td>
            </tr>
            <tr>
                <td id="td2" runat="server" class="bluecol_need" style="width: 20%">班別名稱
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:TextBox ID="TBclass_name" runat="server" Columns="100" MaxLength="50" Width="99%"></asp:TextBox><asp:RequiredFieldValidator ID="class_name" runat="server" ErrorMessage="請輸入班別名稱" Display="None" ControlToValidate="TBclass_name"></asp:RequiredFieldValidator>
                </td>
                <td class="bluecol_need" style="width: 20%">英文名稱
				<asp:Label ID="LabEnameStar" runat="server" ForeColor="Red">*</asp:Label>
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:TextBox ID="ClassEName" runat="server" Columns="130" MaxLength="100" Width="99%"></asp:TextBox><asp:RequiredFieldValidator ID="Re_Class_CEName" runat="server" ErrorMessage="請輸入英文名稱" Display="None" ControlToValidate="ClassEName"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練計畫
                </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="Plan_List" runat="server">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="Re_Plan_List" runat="server" ErrorMessage="請選擇訓練計畫" Display="None" ControlToValidate="Plan_List"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td class="bluecol">計畫年度</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlYears" runat="server"></asp:DropDownList>
                </td>
                <td class="bluecol">轄區分署 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td id="Td1" runat="server" class="bluecol_need">
                    <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="38%"></asp:TextBox>
                    <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server">
                    <asp:RequiredFieldValidator ID="Tcareerid" runat="server" ErrorMessage="請選擇訓練職類" Display="None" ControlToValidate="TB_career_id"></asp:RequiredFieldValidator>
                    <input id="trainValue" style="width: 37px; height: 22px" type="hidden" name="trainValue" runat="server">
                    <input id="jobValue" style="width: 43px; height: 22px" type="hidden" name="jobValue" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">
                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="38%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server"><asp:RequiredFieldValidator ID="fill1b" runat="server" ControlToValidate="txtCJOB_NAME" Display="None" ErrorMessage="請選擇通俗職類"></asp:RequiredFieldValidator><input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">課程內容
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="ComSumm" runat="server" Columns="40" Rows="7" TextMode="MultiLine" Width="70%"></asp:TextBox>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    <asp:ValidationSummary ID="Summary" runat="server" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>
                </td>
            </tr>
        </table>
        <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator" Display="Static"></asp:CustomValidator>
        <%--<asp:Literal ID="clientscript" runat="server"></asp:Literal>--%>
        <input id="Re_ID" type="hidden" name="Re_ID" runat="server" />
        <input id="HidClassID1" type="hidden" name="HidClassID1" runat="server" />
        <input id="Hid_CLSID" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_PERC100" runat="server" />
    </form>
</body>
</html>
