<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="SYS_05_002_add.aspx.vb" Inherits="WDAIIP.SYS_05_002_add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>上稿維護-公告維護</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //檢查日期格式1
        function check_date() {
            if (!checkDate(form1.PostDate.value)) {
                document.form1.PostDate.value = '';
                alert('發布日期起始，請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }
            js1_close();
        }
        //檢查日期格式2
        function check_dateF() {
            if (!checkDate(form1.PostFDate.value)) {
                document.form1.PostFDate.value = '';
                alert('發布日期迄止，請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }
        }

        //TypeList
        function change_TypeList() {
            var vTypeList = $("input[id^='<%=TypeList.ClientID%>']:checked").val();
            //var trSubject1 = $("#<%=trSubject1.ClientID%>");
            //var trHtmlx = $("#<%=trHtmlx.ClientID%>");
            //trDoc1//trDoc2
            $("#trSubject1").show();
            $("#trHtmlx").show();
            $("#trDoc1").hide();
            $("#trDoc2").hide();
            var flag_34 = false;
            if (vTypeList == "3") { flag_34 = true; }
            if (vTypeList == "4") { flag_34 = true; }
            if (flag_34) {
                $("#trSubject1").show();
                $("#trHtmlx").show();
                $("#trDoc1").show();
                $("#trDoc2").show();
                if ($("#Subject").val() == "") {
                    $("#trSubject1").hide();
                    $("#trHtmlx").hide();
                    $("#trDoc1").show();
                    $("#trDoc2").show();
                }
            }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        function set_lab() {
            var MyTable = document.getElementById('isShow');
            var radio1 = MyTable.rows[0].cells[0].children[0];
            var radio2 = MyTable.rows[0].cells[0].children[1];

            var spanPostFDate = document.getElementById('spanPostFDate');
            //var MyTable1 = document.getElementById('Table1');
            //var lab = MyTable1.rows(2).cells(1).children(2);
            //var PostFDate = MyTable1.rows(2).cells(1).children(3);
            //var img = MyTable1.rows(2).cells(1).children(4);
            //var radiolength=MyTable1.rows(2).cells(1).children.length;
            //var radiolength=MyTable.rows(0).cells(0).children.length;
            //debugger;
            if (spanPostFDate) {
                spanPostFDate.style.display = '';
                if (radio1.checked) {
                    spanPostFDate.style.display = 'none';
                    //lab.style.display = 'none';PostFDate.style.display = 'none';img.style.display = 'none';
                    //alert('radio1_lab'+':'+radiolength);
                }
            }

        }
        function htmlencode(s) {
            var div = document.createElement('div');
            div.appendChild(document.createTextNode(s));
            return div.innerHTML;
        }
        function htmldecode(s) {
            var div = document.createElement('div');
            div.innerHTML = s;
            return div.innerText || div.textContent;
        }
        function js1_encode() {
            $('#Hid_context_decode1').val("");
            $('#BtnEncode1').removeAttr('disabled');
            $('#BtnDecode1').removeAttr('disabled');
            var sVal = $('#Subject').val();
            if (sVal.indexOf(">") == -1 && sVal.indexOf("<") == -1) {
                $('#BtnEncode1').attr('disabled', true);
                alert("已編碼!");
                return false;
            }
            $('#Subject').val(htmlencode(sVal));
            return false;
        }
        function js1_decode() {
            $('#BtnEncode1').removeAttr('disabled');
            $('#BtnDecode1').removeAttr('disabled');
            var sVal = $('#Subject').val();
            if (sVal.indexOf(">") > -1 || sVal.indexOf("<") > -1) {
                $('#BtnDecode1').attr('disabled', true);
                alert("已解碼!");
                return false;
            }
            $('#Hid_context_decode1').val("1");
            $('#Subject').val(htmldecode(sVal));
            return false;
        }
        function js1_close() {
            var vTypeList = $("input[id^='<%=TypeList.ClientID%>']:checked").val();
            var vHid_context_decode1 = $('#Hid_context_decode1').val();//已解碼
            if (vTypeList == "3" && vHid_context_decode1 == "1") {
                var x = js1_encode();//編碼
            }
            if (vTypeList == "4" && vHid_context_decode1 == "1") {
                var x = js1_encode();//編碼
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="table_nw" id="Table2" width="100%" cellspacing="1" cellpadding="1">
            <tr>
                <td class="bluecol_need" width="20%">項目(代碼) </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="TypeList" runat="server" RepeatLayout="Flow" CssClass="font" RepeatDirection="Horizontal">
                        <asp:ListItem Value="1">News</asp:ListItem>
                        <asp:ListItem Value="2">新功能</asp:ListItem>
                        <asp:ListItem Value="3">文件下載</asp:ListItem>
                        <asp:ListItem Value="4">影音教學</asp:ListItem>
                    </asp:RadioButtonList>
                    <%--<asp:RequiredFieldValidator ID="MustType" runat="server" ErrorMessage="請選擇項目代碼" Display="None" ControlToValidate="TypeList"></asp:RequiredFieldValidator>--%>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">永遠顯示 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="isShow" runat="server" CssClass="font" RepeatDirection="Horizontal">
                        <asp:ListItem Value="Y">是</asp:ListItem>
                        <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">發布日期 </td>
                <td class="whitecol">
                    <asp:TextBox ID="PostDate" runat="server" Columns="10" ToolTip="日期格式:99/01/31" MaxLength="10" Width="15%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('PostDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    <span id="spanPostFDate" runat="server">
                        <asp:Label ID="Label1" runat="server">至</asp:Label>
                        <asp:TextBox ID="PostFDate" runat="server" Columns="10" ToolTip="日期格式:99/01/31" MaxLength="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('PostFDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        <asp:DropDownList ID="HR2" runat="server"></asp:DropDownList>
                        時：
					    <asp:DropDownList ID="MM2" runat="server"></asp:DropDownList>
                        分 
                    </span>
                    <%--<asp:RequiredFieldValidator ID="MustPostDate" runat="server" ErrorMessage="請輸入或選擇發布日期起始" Display="None" ControlToValidate="PostDate"></asp:RequiredFieldValidator>--%>
                </td>
            </tr>
            <tr id="trSubject1" runat="server">
                <td class="bluecol" width="20%">發布主題<br />
                    -(News/新功能)-必填 </td>
                <td class="whitecol">
                    <asp:TextBox ID="Subject" runat="server" Columns="58" Rows="6" TextMode="MultiLine" Width="70%"></asp:TextBox>
                    <%--<asp:RequiredFieldValidator ID="MustSubject" runat="server" ErrorMessage="請發布主題" Display="None" ControlToValidate="Subject"></asp:RequiredFieldValidator>--%>
                </td>
            </tr>

            <tr id="trHtmlx" runat="server">
                <td class="bluecol" width="20%">HTML功能</td>
                <td class="whitecol">
                    <asp:Button ID="BtnEncode1" Text="編碼" runat="server" CssClass="asp_button_M" OnClientClick="return js1_encode();"></asp:Button>
                    <asp:Button ID="BtnDecode1" Text="解碼" runat="server" CssClass="asp_button_M" OnClientClick="return js1_decode();"></asp:Button>
                </td>
            </tr>

            <tr id="trDoc1" runat="server">
                <td class="bluecol" width="20%">檔名<br />
                    -文件下載-必填 </td>
                <td class="whitecol">
                    <asp:TextBox ID="txtDoc0" runat="server" Columns="58" Width="70%" MaxLength="100"></asp:TextBox>
                </td>
            </tr>
            <tr id="trDoc2" runat="server">
                <td class="bluecol" width="20%">文字說明<br />
                    -文件下載-必填 </td>
                <td class="whitecol">
                    <asp:TextBox ID="txtDoc1" runat="server" Columns="58" Width="70%" MaxLength="200"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">提示週數 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="msgweek" runat="server" CssClass="font" RepeatDirection="Horizontal">
                        <asp:ListItem Value="2" Selected="True">2週</asp:ListItem>
                        <asp:ListItem Value="3">3週</asp:ListItem>
                        <asp:ListItem Value="4">4週</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <%--<tr>
                <td class="bluecol" width="20%">檔案上傳 </td>
                <td class="whitecol">
                    <input id="File1" type="file" name="File1" runat="server" size="60" style="width: auto" accept=".zip,.rar,.pdf,.odt,.ods" />
                    <asp:Button ID="btnUpload1" runat="server" Text="檔案上傳後儲存" CssClass="asp_button_M"></asp:Button>
                    (必須為.zip,.rar,.pdf,.odt,.ods等格式) </td>
            </tr>--%>
            <tr>
                <td colspan="2" class="whitecol" align="center">
                    <asp:Button ID="bt_addrow" Text="儲存" runat="server" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button1" runat="server" Text="回上頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <%--<asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowSummary="False" ShowMessageBox="True" DisplayMode="List" Width="280px"></asp:ValidationSummary>--%>
        <asp:HiddenField ID="hid_HNID" runat="server" />
        <input id="hid_gptodo" type="hidden" runat="server" />
        <input id="hid_AcceptSearch" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_context_decode1" runat="server" />
    </form>
</body>
</html>
