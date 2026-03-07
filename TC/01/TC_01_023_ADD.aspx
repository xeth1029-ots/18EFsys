<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_023_ADD.aspx.vb" Inherits="WDAIIP.TC_01_023_ADD" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%--<!DOCTYPE html>--%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>遠距課程環境設定</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function CheckExtension(fileType2Val) {
            var extension = new Array("jpg", "jpeg", "png", "gif");
            for (var i = 0; i < extension.length; i++) {
                if (fileType2Val == extension[i]) { return true; }
            }
            return false;
        }

        function CanUseExtension() {
            var rst = "";
            var extension = new Array("jpg", "jpeg", "png", "gif");
            for (var i = 0; i < extension.length; i++) {
                rst += ((rst != "" ? "、" : "") + extension[i])
            }
            return rst;
        }

        //function File2Base64(file) {
        //    let fr = new FileReader();
        //    //如果下面的語句執行失敗,需要放入 setTimeout 非同步處理
        //    fr.readAsDataURL(file);
        //    fr.onload = function (e) {
        //        //console.log(this.result);// base64
        //        //console.log(e.target.result);// base64
        //        let base64 = e.target.result;// data:image/jpeg;base64,/9j/4AAQSkZJ
        //        //console.log(base64.constructor);//f String() { [native code] }
        //        //return getWH(base64);
        //        return base64;
        //    }
        //}
        //function getWH(base64) {
        //    var img = new Image();
        //    img.src = base64;
        //    img.onload = function () {
        //        //圖片尺寸 console.log(img.width, img.height);
        //        //return (img.width >= img.height);
        //        return img;
        //    }
        //}
        //function chkImageFileWH(file) {
        //    var img = getWH(File2Base64(file));
        //    console.log('圖片尺寸: '+img.width+img.height);
        //    return (img.width >= img.height);
        //}

        function checkFile1(sizeLimit) {
            //sizeLimit單位:byte   
            const fileInput = document.querySelector('input[type="file"]');
            const file = fileInput.files[0];
            const fileType = file.type;
            const fileType2 = file.type.split('/')[1];
            console.log('fileType : ' + fileType);
            console.log('fileType2 : ' + fileType2);
            var CanUseExt = CanUseExtension()
            console.log('CanUseExt : ' + CanUseExt);
            if (!CheckExtension(fileType2)) {
                alert("只允許上傳 " + CanUseExt + " 檔！");
                return false;
            }
            
            //if (fileType !== 'application/pdf' || fileType2 !== 'pdf') {
            //    alert('只允許上傳 PDF 檔！');
            //    return false;
            //}
            const fileSize = Math.round(file.size / 1024 / 1024);
            const fileSizeLimit = Math.round(sizeLimit / 1024 / 1024);
            //console.log('fileSize : ' + fileSize);
            //console.log('fileSizeLimit : ' + fileSizeLimit);
            if (fileSize > fileSizeLimit) {
                alert('您所選擇的檔案大小為 ' + fileSize + 'MB，超過了上傳上限! (檔案大小限制' + fileSizeLimit + 'MB以下)\n不允許上傳！');
                document.getElementById("File1").outerHTML = '<input name="File1" type="file" id="File1" size="66" accept=".jpg,.gif,.bmp,.png">';
                return false;
            }
            //if (!chkImageFileWH(file)) {
            //    alert('上傳照片檔案，須為橫式圖片(建議圖片大小：960x480)\n，圖片長寬有誤，不允許上傳！');
            //    document.getElementById("File1").outerHTML = '<input name="File1" type="file" id="File1" size="66" accept=".jpg,.gif,.bmp,.png">';
            //    return false;
            //}
            return true;
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;遠距課程環境設定</asp:Label>
                </td>
            </tr>
        </table>
        <%--<input type="hidden" id="hidLID"  runat="server" />--%>
        <table id="Table1" class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" width="20%">訓練機構</td>
                <td class="whitecol" colspan="3">
                    <asp:Label ID="labORGNAME" runat="server" Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">環境代號 </td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="RMTNO" runat="server" MaxLength="10" Style="width: 50%;"></asp:TextBox>
                    <font color="red">限英數字10字以內</font></td>
                <td class="bluecol_need" width="20%">環境名稱 </td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="RMTNAME" runat="server" MaxLength="50" Style="width: 70%;"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">教學軟體(可複選)</td>
                <td class="whitecol" colspan="3">
                    <asp:CheckBoxList ID="cbl_TEACHSOFT" runat="server" RepeatDirection="Horizontal"></asp:CheckBoxList>
                    <label for="TEACHSOFT_OTH">其他(請說明)</label>
                    <asp:TextBox ID="TEACHSOFT_OTH" runat="server" MaxLength="100"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">教學設備(可複選)</td>
                <td class="whitecol" colspan="3">
                    <asp:CheckBoxList ID="cbl_TEACHDEVICE" runat="server" RepeatDirection="Horizontal"></asp:CheckBoxList>
                    <label for="TEACHDEVICE_OTH">其他(請說明)</label>
                    <asp:TextBox ID="TEACHDEVICE_OTH" runat="server" MaxLength="100"></asp:TextBox>
                </td>
            </tr>

            <tr>
                <td class="bluecol_need">網路環境(可複選)</td>
                <td class="whitecol" colspan="3">
                    <asp:CheckBox ID="CBX_CABLENETWORK" runat="server" /><label for="CBX_CABLENETWORK">有線網路</label>
                    <br />
                    <label for="CABLEDLRATE">頻寬：下載速率(Download)</label>
                    <asp:TextBox ID="CABLEDLRATE" runat="server" MaxLength="10"></asp:TextBox>Mbps
                    <label for="CABLEUPRATE">上傳速率(Upload)</label>
                    <asp:TextBox ID="CABLEUPRATE" runat="server" MaxLength="10"></asp:TextBox>Mbps
			 <br />
                    <asp:CheckBox ID="CBX_WIFINETWORK" runat="server" /><label for="CBX_WIFINETWORK">無線網路</label><br>
                    <label for="WIFIDLRATE">頻寬：下載速率(Download)</label>
                    <asp:TextBox ID="WIFIDLRATE" runat="server" MaxLength="10"></asp:TextBox>Mbps
                    <label for="WIFIUPRATE">上傳速率(Upload)</label>
                    <asp:TextBox ID="WIFIUPRATE" runat="server" MaxLength="10"></asp:TextBox>Mbps
                    <br />
                    <span id="Labmsg_NETWORK" style="color: Red;">建議上傳及下載至少100Mbps</span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">啟用</td>
                <td class="whitecol" colspan="3">
                    <asp:RadioButtonList ID="rblMODIFYTYPE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="Y" Selected="True">啟用</asp:ListItem>
                        <asp:ListItem Value="N">停用</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">教學錄影設備</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="VIDEODEVICE" runat="server" TextMode="MultiLine" Rows="4" Columns="60" placeholder="(請詳列課堂中所使用之教學錄影設備名稱及規格)"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">教學軟體及設備說明</td>
                <td colspan="3" class="whitecol">軟體：<br />
                    <asp:TextBox ID="SOFTDESC" runat="server" TextMode="MultiLine" Rows="4" Columns="60" placeholder="(請說明教學使用之軟體名稱、版本)"></asp:TextBox><br />
                    設備：<br />
                    <asp:TextBox ID="DEVICEDESC" runat="server" TextMode="MultiLine" Rows="4" Columns="60" placeholder="(請說明教學所使用之設備名稱、規格)"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <font color="red">
                        <asp:Label ID="labMsg1" runat="server"></asp:Label></font>
                    <table class="font" id="DataGrid3Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle />
                                    <Columns>
                                        <asp:BoundColumn DataField="depID" HeaderText="序號">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="RMTPIC" HeaderText="照片種類">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="圖檔名稱">
                                            <HeaderStyle Width="70%" />
                                            <ItemTemplate>
                                                <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                                <input id="HFileName" type="hidden" runat="server" />
                                                <input id="HiddepID" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="Butdel4" runat="server" Text="刪除" CommandName="del4" CssClass="asp_button_M"></asp:Button>
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
                <td class="bluecol">照片上傳</td>
                <td colspan="3" class="whitecol">&nbsp;
                    <asp:DropDownList ID="DDL_RMTPIC" runat="server"></asp:DropDownList>
                    <input id="File1" type="file" size="66" name="File1" runat="server" accept=".jpg,.gif,.bmp,.png" />
                    <asp:Button ID="But1" runat="server" Text="確定上傳" CausesValidation="False" CssClass="asp_button_M"></asp:Button><br />
                    <font color="red"><span id="lab_msg_WH1">上傳照片檔案需為圖片類型，檔案請小於或等於10M (建議圖片大小：960x480)，須為橫式圖片</span></font>
                </td>
            </tr>
            <tr>
                <td class="whitecol"></td>
                <td colspan="3" class="whitecol">
                    <font color="red">備註：<br />
                        1.	應敘明每位師資教學環境所需相關軟、硬體設施或設備，並另提供實體照片或佐證資料。<br />
                        2.	教學軟體以使用Google Meet、U會議、Microsoft Teams及Cisco Webex等為原則。<br />
                        3.	教學設備應包含視訊鏡頭、麥克風及收音喇叭等設備。<br />
                        4.	網路環境應能提供參訓學員較高解析度且流暢之課程應備之設施或設備，並敘明網路速度。<br />
                        5.	應敘明教學錄影及其他因應課程需求所需之軟、硬體設施或設備。
                    </font>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="BtnSAVEDATA1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="BtnGOBACK" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    <br />
                    <font color="red">記得按儲存(確定)鈕</font>
                </td>
            </tr>
        </table>

        <input id="RIDValue" type="hidden" runat="server" name="RIDValue" />
        <asp:HiddenField ID="Hid_COMIDNO" runat="server" />
        <asp:HiddenField ID="Hid_ORGID" runat="server" />
        <asp:HiddenField ID="Hid_RMTID" runat="server" />
        <asp:HiddenField ID="Hid_ORG_REMOTER_GUID1" runat="server" />
    </form>
</body>
</html>
