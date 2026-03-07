<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_10_001.aspx.vb" Inherits="WDAIIP.TC_10_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>審查委員名單</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //檢查zipcode(City欄位名,Zip obj,Zip輸入內容)
        function getZipName(CityID, oZipID, ZipValue) {
            var ifrmae = document.getElementById('ifmChceckZip');
            if (!isBlank(oZipID)) {
                //debugger;
                if (isUnsignedInt(ZipValue) && ZipValue.length == 3) {
                    if (_isIE) {
                        ifrmae.document.form1.hidCityID.value = CityID;
                        ifrmae.document.form1.hidZipID.value = oZipID.id;
                        ifrmae.document.form1.hidValue.value = ZipValue;
                        ifrmae.document.form1.submit();
                    }
                    else {
                        var ifrmaeDoc = (ifrmae.contentWindow || ifrmae.contentDocument);
                        ifrmae.contentDocument.getElementById("hidCityID").value = CityID;
                        ifrmae.contentDocument.getElementById("hidZipID").value = oZipID.id;
                        ifrmae.contentDocument.getElementById("hidValue").value = ZipValue;
                        if (ifrmaeDoc.document) ifrmaeDoc = ifrmaeDoc.document;
                        ifrmaeDoc.getElementById("form1").submit(); // ## error 
                        //ifrmae.contentWindow.formSubmit();
                    }
                } else {
                    //debugger;
                    oZipID.value = '';
                    document.getElementById(CityID).value = '';
                    oZipID.focus();
                    alert('查無' + ZipValue + '郵遞區號!');
                }
            } else {
                //debugger;
                document.getElementById(CityID).value = '';
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var objList = document.getElementById(hidobj);
            if (!objList) { return; }
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (objList.value != getCheckBoxListValue(obj).charAt(0)) {
                objList.value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        objList.value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

        //儲存前檢查
        function chkSaveData1() {
            var msg = '';
            var ddlRECRUIT = document.getElementById('ddlRECRUIT');
            var MBRNAME = document.getElementById('MBRNAME');
            var UNITNAME = document.getElementById('UNITNAME');
            var JOBTITLE = document.getElementById('JOBTITLE');
            var PHONE = document.getElementById('PHONE');
            var SERVUNIT1 = document.getElementById('SERVUNIT1');
            var SERVTIME1 = document.getElementById('SERVTIME1');
            var JOBTITLE1 = document.getElementById('JOBTITLE1');

            //var cbPUSHDISTID = document.getElementById('cbPUSHDISTID');
            //var rblRUNTRAIN = document.getElementById('rblRUNTRAIN');
            //var cbTRAINDISTID = document.getElementById('cbTRAINDISTID');

            if (ddlRECRUIT.value == '') { msg += '請選擇 遴聘類別\n'; }
            if (MBRNAME.value == '') { msg += '請輸入 審查委員姓名\n'; }
            if (UNITNAME.value == '') { msg += '請輸入 現職服務機構\n'; }
            if (JOBTITLE.value == '') { msg += '請輸入 職稱\n'; }
            if (PHONE.value == '') { msg += '請輸入 連絡電話\n'; }
            if (SERVUNIT1.value == '') { msg += '請輸入 服務單位1\n'; }
            if (SERVTIME1.value == '') { msg += '請輸入 服務時間1\n'; }
            if (JOBTITLE1.value == '') { msg += '請輸入 職稱1\n'; }
            var v_cbPUSHDISTID = getCheckBoxListValue('cbPUSHDISTID');
            if (parseInt(v_cbPUSHDISTID, 10) == 0) { msg += '請選擇 推薦分署(至少一筆)\n'; }
            //'是否辦訓/'辦訓轄區
            var v_rblRUNTRAIN = getRadioValue(document.form1.rblRUNTRAIN); //取得 RadioButtonList 值 
            var v_cbTRAINDISTID = getCheckBoxListValue('cbTRAINDISTID'); //CheckBoxList控制項 
            if (v_rblRUNTRAIN == 'Y' && parseInt(v_cbTRAINDISTID, 10) == 0) { msg += '是否辦訓 若選擇「是」，請選擇 辦訓轄區(至少一筆)\n'; }
            if (msg != '') {
                msg += '!!!\n';
                alert(msg);
                return false;
            }
            return true;
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查委員管理&gt;&gt;審查委員名單</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="panelEdit" runat="server">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="head_navy" colspan="4">審查委員基本資料</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">遴聘類別</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="ddlRECRUIT" runat="server">
                            <asp:ListItem Value="">==請選擇==</asp:ListItem>
                            <asp:ListItem Value="A">A-產業界</asp:ListItem>
                            <asp:ListItem Value="B">B-學術界</asp:ListItem>
                            <asp:ListItem Value="C">C-勞工團體代表</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">審查委員姓名</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="MBRNAME" runat="server" MaxLength="50" Columns="30"></asp:TextBox>
                        <%--&nbsp;&nbsp;<asp:TextBox ID="MBRNMSEQ" runat="server" MaxLength="3" Columns="3"></asp:TextBox>(同名序號,同名可補序號)--%>
                        &nbsp;&nbsp;<asp:TextBox ID="MBRNMSEQ" runat="server" MaxLength="3" Columns="3"></asp:TextBox>(同名序號,可補序號區分,預設為空,可輸入大於0整數)
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">現職服務機構</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="UNITNAME" runat="server" MaxLength="50" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need">職稱</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="JOBTITLE" runat="server" MaxLength="100" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">學歷</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="DEGREE" runat="server" MaxLength="100" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">證照</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="CERTIFICAT" runat="server" MaxLength="300" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need">專業背景</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="SPECIALTY" runat="server" MaxLength="300" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need">連絡電話</td>
                    <td class="whitecol">
                        <asp:TextBox ID="PHONE" runat="server" MaxLength="20" Columns="22"></asp:TextBox></td>
                    <td class="bluecol">連絡電話2</td>
                    <td class="whitecol">
                        <asp:TextBox ID="PHONE2" runat="server" MaxLength="20" Columns="22"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">手機</td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="CELLPHONE" runat="server" MaxLength="20" Columns="22"></asp:TextBox></td>
                    <td class="bluecol" width="20%">手機2</td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="CELLPHONE2" runat="server" MaxLength="20" Columns="22"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">傳真</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="CONFAX" runat="server" MaxLength="20" Columns="22"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">電子郵件</td>
                    <td class="whitecol">
                        <asp:TextBox ID="EMAIL" runat="server" MaxLength="70" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">電子郵件2</td>
                    <td class="whitecol">
                        <asp:TextBox ID="EMAIL2" runat="server" MaxLength="70" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">地址 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="MADDRESS" runat="server" Columns="60" MaxLength="120"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">地址2</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="MADDRESS2" runat="server" Columns="60" MaxLength="120"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">備註 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:TextBox ID="RMNOTE1" runat="server" TextMode="MultiLine" Columns="50" Rows="8" MaxLength="300"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="center" colspan="4" class="head_navy">經歷</td>
                </tr>
                <tr>
                    <td class="bluecol_need">服務單位1</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="SERVUNIT1" runat="server" MaxLength="100" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need">服務時間1</td>
                    <td class="whitecol">
                        <asp:TextBox ID="SERVTIME1" runat="server" MaxLength="100" Columns="33"></asp:TextBox></td>
                    <td class="bluecol_need">職稱1</td>
                    <td class="whitecol">
                        <asp:TextBox ID="JOBTITLE1" runat="server" MaxLength="100" Columns="33"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="4">
                        <br style="line-height: 10px;" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">服務單位2</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="SERVUNIT2" runat="server" MaxLength="100" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">服務時間2</td>
                    <td class="whitecol">
                        <asp:TextBox ID="SERVTIME2" runat="server" MaxLength="100" Columns="33"></asp:TextBox></td>
                    <td class="bluecol">職稱2</td>
                    <td class="whitecol">
                        <asp:TextBox ID="JOBTITLE2" runat="server" MaxLength="100" Columns="33"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="4">
                        <br style="line-height: 10px;" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">服務單位3</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="SERVUNIT3" runat="server" MaxLength="100" Columns="55"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">服務時間3</td>
                    <td class="whitecol">
                        <asp:TextBox ID="SERVTIME3" runat="server" MaxLength="100" Columns="33"></asp:TextBox></td>
                    <td class="bluecol">職稱3</td>
                    <td class="whitecol">
                        <asp:TextBox ID="JOBTITLE3" runat="server" Columns="33" MaxLength="100"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="4" class="head_navy">其他</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">推薦分署 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:CheckBoxList ID="cbPUSHDISTID" runat="server" CssClass="font" RepeatColumns="3" RepeatDirection="Horizontal">
                        </asp:CheckBoxList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">推薦理由 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:TextBox ID="PUSHREASON" runat="server" TextMode="MultiLine" Columns="50" Rows="8" MaxLength="1000"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">是否辦訓 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:RadioButtonList ID="rblRUNTRAIN" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                            <asp:ListItem Value="Y">是</asp:ListItem>
                        </asp:RadioButtonList></td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">辦訓轄區 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:CheckBoxList ID="cbTRAINDISTID" runat="server" CssClass="font" RepeatColumns="3" RepeatDirection="Horizontal">
                        </asp:CheckBoxList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">審查課程職類</td>
                    <td colspan="3" class="whitecol">
                        <input id="cblGOVCODE3_Hidden" type="hidden" value="0" name="cblGOVCODE3_Hidden" runat="server" />
                        <asp:CheckBoxList ID="cblGOVCODE3" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" Width="100%"></asp:CheckBoxList></td>
                </tr>
                <tr>
                    <td class="bluecol">新增年度</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlADDYEARS" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">啟用 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:RadioButtonList ID="rblSTOPUSE" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="" Selected="True">啟用</asp:ListItem>
                            <asp:ListItem Value="Y">停用</asp:ListItem>
                        </asp:RadioButtonList>

                    </td>
                </tr>
                <%--<tr><td class="whitecol" colspan="4" width="100%"><center><font color="red" style="font-weight: bold;">本人同意勞動部勞動力發展署暨所屬機關，為本人提供職業訓練及就業服務時使用</font></center></td></tr>--%>                
                <tr>
                    <td colspan="4" class="whitecol" align="center" width="100%">
                        <%--Button1_Click Button1--%>
                        <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="panelSch" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <tr>
                    <td class="bluecol" width="20%">遴聘類別</td>
                    <td class="whitecol" width="30%">
                        <asp:DropDownList ID="SCH_RECRUIT" runat="server">
                            <asp:ListItem Value="">==請選擇==</asp:ListItem>
                            <asp:ListItem Value="A">A-產業界</asp:ListItem>
                            <asp:ListItem Value="B">B-學術界</asp:ListItem>
                            <asp:ListItem Value="C">C-勞工團體代表</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td class="bluecol" width="20%">現職服務機構</td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="SCH_UNITNAME" runat="server" MaxLength="55"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">審查委員姓名</td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCH_MBRNAME" runat="server" MaxLength="55"></asp:TextBox>
                    </td>

                    <td class="bluecol">職稱</td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCH_JOBTITLE" runat="server" MaxLength="100"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">推薦分署</td>
                    <td class="whitecol" colspan="3">
                        <asp:CheckBoxList ID="SCH_PUSHDISTID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3"></asp:CheckBoxList>
                        <input id="SCH_PUSHDISTID_List" type="hidden" value="0" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">審查課程職類</td>
                    <td colspan="3" class="whitecol">
                        <input id="SCH_cblGOVCODE3Hidden" runat="server" type="hidden" value="0" name="SCH_cblGOVCODE3Hidden" />
                        <asp:CheckBoxList ID="SCH_cblGOVCODE3" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" Width="100%"></asp:CheckBoxList></td>
                </tr>
                <tr id="trImport1" runat="server">
                    <td class="bluecol">匯入審查委員名冊 </td>
                    <td class="whitecol" colspan="3">
                        <input id="File1" type="file" size="50" name="File1" runat="server" accept=".xls,.ods" />
                        <asp:Button ID="btnIMPORT1" runat="server" Text="匯入名冊" CssClass="asp_Export_M"></asp:Button>(必須為ods或xls格式)
                        <br />
                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/IMP_EXAMINER_v2.zip" ForeColor="#8080FF" CssClass="font">下載整批上載格式檔</asp:HyperLink>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">含已停用資料</td>
                    <td class="whitecol" colspan="3">
                        <asp:CheckBox ID="CHECK_STOPUSE" runat="server" Text="含已停用資料" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="4">
                        <asp:Button ID="BtnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnAddnew" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnExport" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <table id="tbDataGrid1" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <%--序號、遴聘類別、姓名、現職服務機構、職稱、推薦分署--%>
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="RECRUIT" HeaderText="遴聘類別"></asp:BoundColumn>--%>
                                <asp:TemplateColumn HeaderText="遴聘類別">
                                    <ItemTemplate>
                                        <asp:Label ID="labRECRUIT_N" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="MBRNAME" HeaderText="姓名"></asp:BoundColumn>
                                <asp:BoundColumn DataField="UNITNAME" HeaderText="現職服務機構"></asp:BoundColumn>
                                <asp:BoundColumn DataField="JOBTITLE" HeaderText="職稱"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="PUSHDISTID" HeaderText="推薦分署"></asp:BoundColumn>--%>
                                <asp:TemplateColumn HeaderText="推薦分署">
                                    <ItemTemplate>
                                        <asp:Label ID="labPUSHDISTID_N" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                    <HeaderStyle HorizontalAlign="center"></HeaderStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="BTNEDIT1" runat="server" Text="修改" CommandName="EDIT1" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNSTOP1" runat="server" Text="停用" CommandName="STOP1" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNDEL1" runat="server" Text="刪除" CommandName="DEL1" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNVIEW1" runat="server" Text="查看" CommandName="VIEW1" CssClass="asp_button_M" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </td>
                </tr>
            </table>
        </asp:Panel>

        <asp:HiddenField ID="Hid_EMSEQ" runat="server" />
        <asp:HiddenField ID="Hid_STOPUSE" runat="server" />
    </form>
    <iframe id="ifmChceckZip" height="0%" src="../../Common/CheckZip.aspx" width="0%" style="display: none" />
</body>
</html>