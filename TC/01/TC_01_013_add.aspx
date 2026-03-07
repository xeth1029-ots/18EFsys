<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_013_add.aspx.vb" Inherits="WDAIIP.TC_01_013_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>計畫場地設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="JavaScript">
        function checkFile(sizeLimit) {  //sizeLimit單位:byte   
            var f = document.FileForm;
            document.MM_returnValue = false;
            re = /(\.jpg|\.gif)$/i;
            if (!re.test(f.file1.value)) {
                alert("只允許上傳JPG或GIF影像檔");
            } else {
                var img = new Image();
                document.MM_returnValue = false;
                img.sizeLimit = sizeLimit;
                img.src = 'file:///' + f.file1.value;
                img.onload = showImageDimensions;
            }
        }

        function showImageDimensions() {
            if (this.fileSize > this.sizeLimit) {
                alert('您所選擇的檔案大小為 ' + (this.fileSize / 1000) + ' kb，\n超過了上傳上限 ' + (this.sizeLimit / 1000) + ' kb！\n不允許上傳！');
                document.getElementById("File1").outerHTML = '<input type="file" name="file1" size="20" id="file1" accept=".jpg,.gif">';
            }
            else {
                document.MM_returnValue = true;
            }
        }

        //檢查檔案上傳位置
        function CheckAddPIC() {
            var msg = '';
            if (document.getElementById('File1').value == '') msg += '請輸入檔案上傳位置\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function chkdata() {
            var msg = '';
            var hidLID = document.getElementById("hidLID"); //階層代碼
            var ConNum = document.getElementById("ConNum"); //容納人數
            var txtPingNumber = document.getElementById("txtPingNumber"); //坪數
            if (document.form1.PlaceID.value == '') msg = msg + '請輸入場地代碼!\n';
            if (document.form1.PlaceName.value == '') msg = msg + '請輸入場地名稱!\n';
            if (!isChecked(document.form1.IFICation)) msg = msg + '請選擇場地種類!\n';
            if (!isChecked(document.form1.FactMode)) msg = msg + '請選擇場地類型!\n';
            if (getValue('FactMode') == '99') {
                if (document.form1.ModeOther.value == '') msg = msg + '請輸入場地類型其他說明!\n';
            }
            if (!isChecked(document.form1.AreaPoss)) msg = msg + '請選擇場地屬地!\n';
            if (ConNum.value == '') msg += '請輸入容納人數!\n';
            if (ConNum.value != '' && !isUnsignedInt(ConNum.value)) msg += '容納人數必須為數字\n';
            if (txtPingNumber.value == '') msg += '請輸入坪數!\n';
            if (txtPingNumber.value != '' && !isUnsignedInt(txtPingNumber.value) && !isPositiveFloat(txtPingNumber.value)) msg += '坪數 應為數字格式(可含小數點4位)\n';
            if (isPositiveFloat(txtPingNumber.value)) {
                re = /^[0-9]+(\.[0-9]{0,4})?$/;
                if (!re.test(txtPingNumber.value)) {
                    msg += '坪數 小數點 最多只能輸入到小數點第4位\n';
                }
            }

            $('#city_code').val($.trim($('#city_code').val()));
            $('#ZIPB3').val($.trim($('#ZIPB3').val()));
            $('#Address').val($.trim($('#Address').val()));
            if ($('#city_code').val() == '') msg += '請輸入場地地址郵遞區號\n'
            if ($('#ZIPB3').val() == '') { msg += '請輸入場地地址郵遞區號後2碼\n'; }
            if ($('#ZIPB3').val() != '') { msg += checkzip23(true, '場地地址', 'ZIPB3'); }
            //else {
            //    if (!isUnsignedInt($('#ZIPB3').val()) || parseInt($('#ZIPB3').val(), 10) < 1) { msg += '場地地址郵遞區號後2碼或後3碼必須為數字，且不得輸入 00\n'; }
            //    if ($('#ZIPB3').val().length != 2 && $('#ZIPB3').val().length != 3) { msg += '場地地址郵遞區號後2碼或後3碼長度必須為 2碼或3碼(例 01 或 001)\n'; }
            //}
            if ($('#Address').val() == '') msg += '請輸入場地地址\n' //20090520 fix				

            if (document.form1.ContactEMail.value != '' && !checkEmail(document.form1.ContactEMail.value)) { msg += '請輸入正確的E-mail格式\n'; }
            if (checkMaxLen(document.getElementById('Hwdesc').value, 1000)) {
                msg += '【硬體設施說明】長度不可超過1000字元\n';
            }
            if (checkMaxLen(document.getElementById('OtherDesc').value, 1000)) {
                msg += '【其他設施說明】長度不可超過1000字元\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
            /*
            else if (hidLID.value >= 2) {
                if (!confirm("儲存後將無法修改，如要修改資料請洽分署人員")) return false;
            }
            */
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input type="hidden" id="hidLID" runat="server" />
        <%--
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;計畫場地設定</asp:Label>
                    <asp:Label ID="ProcessType" runat="server" ForeColor="#990000"></asp:Label>
                </td>
            </tr>
        </table>
        --%>
        <table class="table_nw" id="Table1" width="100%" runat="server" cellpadding="1" cellspacing="1">
            <tbody>
                <tr>
                    <td class="bluecol" width="20%">訓練機構</td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="labORGNAME" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td id="td1" runat="server" class="bluecol_need" width="20%">場地代號 </td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="PlaceID" MaxLength="10" runat="server" Columns="12" Width="50%"></asp:TextBox>
                        <font color="red">限英數字10字以內</font></td>
                    <td id="Td2" runat="server" class="bluecol_need">場地名稱 </td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="PlaceName" runat="server" MaxLength="50" Columns="26" Width="70%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td id="Td3" runat="server" class="bluecol_need">場地類別 </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="IFICation" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">學科</asp:ListItem>
                            <asp:ListItem Value="2">術科</asp:ListItem>
                            <asp:ListItem Value="3">共用</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                    <td id="Td15" runat="server" class="bluecol_need">場地屬性 </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="AreaPoss" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">自有</asp:ListItem>
                            <asp:ListItem Value="2">借租用</asp:ListItem>
                            <asp:ListItem Value="3">企業專屬地</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td id="Td4" runat="server" class="bluecol_need">場地類型 </td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="FactMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">教室</asp:ListItem>
                            <asp:ListItem Value="2">演講廳</asp:ListItem>
                            <asp:ListItem Value="3">會議室</asp:ListItem>
                            <asp:ListItem Value="99">其他</asp:ListItem>
                        </asp:RadioButtonList>
                        <asp:TextBox ID="ModeOther" runat="server" Columns="26" MaxLength="100" Width="50%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td id="Td5" runat="server" class="bluecol" width="20%">聯絡人姓名 </td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="ContactName" runat="server" MaxLength="50" Columns="26" Width="50%"></asp:TextBox></td>
                    <td id="Td6" runat="server" class="bluecol" width="20%">聯絡人電話 </td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="ContactPHone" runat="server" MaxLength="50" Columns="26" Width="50%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td id="Td7" runat="server" class="bluecol">聯絡人傳真 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="ContactFax" runat="server" MaxLength="64" Columns="26" Width="50%"></asp:TextBox></td>
                    <td id="Td8" runat="server" class="bluecol">電子郵件 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="ContactEMail" runat="server" MaxLength="64" Columns="26" Width="50%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">負責人姓名 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="MasterName" runat="server" MaxLength="30" Columns="26" Width="50%"></asp:TextBox></td>
                    <td class="bluecol_need">訓練容納人數</td>
                    <td class="whitecol">
                        <asp:TextBox ID="ConNum" runat="server" MaxLength="5" Columns="26" Width="50%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need">坪數</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtPingNumber" runat="server" MaxLength="14" Columns="15"></asp:TextBox></td>
                    <td class="bluecol_need">啟用</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblMODIFYTYPE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="Y" Selected="True">啟用</asp:ListItem>
                            <asp:ListItem Value="N">停用</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">場地地址 </td>
                    <td class="whitecol" colspan="3">
                        <input id="city_code" onfocus="this.blur()" maxlength="3" runat="server" />－
                        <input id="ZIPB3" maxlength="3" runat="server" />
                        <input id="hidZIP6W" type="hidden" runat="server" />
                        <asp:Literal ID="Litcity_code" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                        <br />
                        <asp:TextBox ID="TBCity" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                        <input id="findzip_but" runat="server" type="button" value="..." class="asp_button_Mini">
                        <asp:TextBox ID="Address" runat="server" Width="55%" MaxLength="250"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">硬體設施說明 </td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="Hwdesc" Width="30%" runat="server" Columns="20" TextMode="MultiLine" Rows="4" MaxLength="1000" placeholder="(請詳列課堂中所使用之硬體設備項目、數量、單位)"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">其他設施說明 </td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="OtherDesc" Width="30%" runat="server" Columns="20" TextMode="MultiLine" Rows="4" MaxLength="1000"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">場地圖片 </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="depID" runat="server"></asp:DropDownList>
                        <input id="File1" type="file" size="66" name="File1" runat="server" accept=".jpg,.gif,.bmp,.png" />
                        <asp:Button ID="But1" runat="server" Text="確定上傳" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <%--<font color="red">上傳照片檔案需為JPG檔，檔案請小於或等於2M，大小：320X240</font>--%>
                        <font color="red">
                            <asp:Label ID="lab_msg_WH1" runat="server"></asp:Label></font>
                        <table class="font" id="DataGrid3Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="Datagrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle />
                                        <Columns>
                                            <asp:BoundColumn DataField="depID" HeaderText="教室">
                                                <HeaderStyle Width="10%" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="PTID" HeaderText="序號">
                                                <HeaderStyle Width="10%" />
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="圖檔名稱">
                                                <HeaderStyle Width="70%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                                    <input id="HFileName" style="width: 87px; height: 22px" type="hidden" size="9" runat="server">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:Button ID="But4" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="4" class="whitecol">
                        <asp:Button ID="btnAdd" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="But5" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <br />
                        <font color="red">記得按儲存(確定)鈕</font>
                    </td>
                </tr>
            </tbody>
        </table>
        <asp:HiddenField ID="Hid_comidno" runat="server" />
        <asp:HiddenField ID="Hid_PlaceID" runat="server" />
        <asp:HiddenField ID="Hid_TRAINPLACE_GUID1" runat="server" />

    </form>
</body>
</html>
