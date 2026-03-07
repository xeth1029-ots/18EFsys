<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_001_add.aspx.vb" Inherits="WDAIIP.SYS_01_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>帳號設定新增</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GetID() {
            //Button3
            //var jnameid = document.getElementById('nameid');
            var jnameid = $("#nameid");
            //alert(nameid.length);
            var cst_errorMsg4a = "輸入新增帳號字串太短，請超過5字以上(含)!";
            var cst_errorMsg4b = "輸入新增帳號字串太長，請勿超過15字!";
            if (jnameid.val().length < 5) { alert(cst_errorMsg4a); return false; }
            if (jnameid.val().length > 15) { alert(cst_errorMsg4b); return false; }
            //wopen('../../Common/CheckID.aspx?id=' + jnameid.val(), 'CheckID', 160, 50, 0);
            __doPostBack('CustomPostBack', 'myCheckID');  //(用來取代原先的彈跳視窗，by:20180926)
        }

        function checkIDNO(source, arguments) {
            //common.js
            var flag1 = checkId(arguments.Value);
            var flag2 = checkId2(arguments.Value);
            arguments.IsValid = ((flag1 || flag2) ? true : false);
        }

        /*
        function but_chg() {
            var jRole = $("#Role");
            var jbut_LevPlan = $("#but_LevPlan");
            jbut_LevPlan.removeAttr('disabled');
            if (jRole.val() == "1") {
                jbut_LevPlan.prop('disabled', true);
                jbut_LevPlan.attr('disabled', 'disabled');
            }
        }
        */

        function CheckData(source, args) {
            args.IsValid = true;
            source.errormessage = "";
            var vOrgName = $("#orgname").val();
            //var visBlack = $("#isBlack").val();
            var msg = '';
            //var orgname = document.form1.orgname.value;
            var msg1 = '';
            var msg2 = '';
            if ($("#isBlack").val() == 'Y') {
                msg1 = vOrgName + "，已列入處分名單，是否確定繼續？"
                msg2 = vOrgName + "，已列入處分名單!!"
                if (!confirm(msg1)) {
                    args.IsValid = false;
                    source.errormessage += msg2;
                }
            }
        }
    </script>
    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 49px; }
        .auto-style2 { color: #333333; padding: 4px; height: 49px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBplan" runat="server" onfocus="this.blur()" MaxLength="100" Width="60%"></asp:TextBox>
                                <%--onclick="javascript: wopen('../../Common/LevPlan.aspx?winreload=1&amp;OrgField=orgname&amp;fisBlack=isBlack&amp;SAH=Y', '計畫階段', 850, 400, 1); document.form1.winreload.value = 1;" --%>
                                <input id="but_LevPlan" type="button" value="選擇" runat="server" class="asp_button_M" />
                                <asp:RequiredFieldValidator ID="mstplan" runat="server" ErrorMessage="請選擇計畫階層-訓練機構" Display="None" ControlToValidate="TBplan"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="width: 20%">帳號 </td>
                            <td class="whitecol" style="width: 80%">
                                <asp:TextBox ID="nameid" runat="server" MaxLength="15" Width="30%"></asp:TextBox>
                                <input id="Button3" type="button" value="檢查帳號" name="Button3" runat="server" class="asp_Export_M" /><%--GetID--%>
                                <asp:Button ID="btnChkAccount" runat="server" Text="檢查帳號(x)" Visible="false" class="asp_Export_M" />
                                <asp:RequiredFieldValidator ID="mstid" runat="server" ErrorMessage="請輸入帳號" Display="None" ControlToValidate="nameid"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="mstnum" runat="server" ErrorMessage="帳號請輸入數字或英文字" Display="None" ControlToValidate="nameid" ValidationExpression="[a-zA-Z0-9]{0,15}"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <%--
                        AA9EC75250CE2E87851E04606970A05375A96F9AA608280E59C41D41CA892CAD37B031447C2F421A3FFAF59BB9599A507E2FEC20174AC2194CC5A43C5DC3
                            E92817D9E7186D7025C425A8A7918F17A294C23725775B927BE73CCE368038A0541817490B5FC9CA6B7C0B6FEC6A5F231831E903D3F24F4447B141E8
                            E143B902F6724DF36391C7B962BB388AB8FB375680431B20B8F890D7AE0FB77221A4161F86A1368F7AA728E6765B7896640241834BBEFD13E2C78693
                            9822ECA079E5F1A7006C33EB9B71E3C5894B7EEC7ADFE14B3E72C08BB2BE1A68D793A2A612B1FE2BD76936115A15A91A09D4E4086C05618E89C8027E
                            B31533F94003AC7345DAB1972155E4954851281FBEB32BD5FD2B7ED3306C3B59DFC949E1F863FFCAC429A3AF4DC83A06CDB90157F61A04A2A1086029
                            CAFC25D53C22DA3EF532CF38D8DF1288125454A4C2DF8096E63740A1A6BBD185B6DFB832DE14959F61325EBE30FF2D208E0D106965A985DE41A87FA0
                            351F383EC46EE190753F350A9696B708DB92D54FB41DFBA783AD5DA54FA59CF805B9B9A675524B1C4A709C18FABF34F9454C1C63DAFE25C62F92CC8F
                            E1C185E79224CB2D0A4B83BAEED23DD683B45BF12912F27F83CA63513248A8AAB2E8DB9C4533A6AF4BE2BD9C0A789704758F936AD399104038DAAC1B
                            92964BFE992ED33836CB8CE86B8B738337AB940DF8F79FF3FE331109085CCA688169E76A2A45E9674A61FBB4D704510E3168F49EF8271F338AC1B63B
                            CDBF86C1B235E1C0A45329324FC4CB151D5AD109A9465A965027D37EE1FABEB8A0F8F05A2A26FCA3029D3085A0F5C44FFCB8D5977328576E341321DA
                            01DEF88CBC8D144F631CDD460280D9FF93E854F7974685353D7C9A22771C491A6271761283CBD0119BE8C065B4DFE797579FFAFC1509858B63B21AF2
                            4570617D2E6CB7195DC0578CE3C97230F045E03F0124A8C4466ADDA7E006FBF1AC1AB9027D43847B3426F0E2478B9A8749E4767F17726509E062A158
                            26A61A2FF1637F126A509F2B891D97FCD81225E00365837F4E837C8B4152788E9E2273ABE10D3A00704414E2EBA8B380259744A826D5E841213FC6B9
                            FADB152E04E968919B4FD49796D56ED328CE4B44608E7DB6741812161DE13EDA33D2F28C57E60228B4A82C1A7E76F761FB5053A5608F283E8694FAF8
                            95B7DADB430D03BF8EB1FD7620B5D7C653D62932383073689BC59E970B0190B39FAEDEDE2BA596E4B4CE64D1420D350456B46E2B1B1F9273A3616927
                            B81E784233FA4D65A938B39873D14587D38D86FCFBB197F88653704130B6F877CB89923CACFFAF32C663A743CC8CE92ED0FEDCCD531798067A63203A
                            0B80E601C95818DE804AD67B79B3E1B27A7B939763D01B4CA8EB5AAC4752ADCB52100DB115802A2BA4DFE616ED8DA62491BA17B41FB9EDCEF7186F73
                            D66FFB29B9C112184DA8F3E0DE7FBD083974D1E61A192097D406E5F90153A764CB47CE2E3E292267B101004E4288F0AF933597D5D2FC3582EA3C7D4D
                            1D29470314F532BC5E448AE3EC666BF30317D73251BE3929DD85876998F786E730D232A0AF27E954F3093C0F652402964CBFA68FC39C839CA893067D
                            84855A14BA14E0508023A4647764243D87502601FCE18B186066222344995A4092C3A79016F685A030D17C99BAF0E512024BD9812EF08C6F8B79198E
                            E0B1C9201F1014AD3CB8E5FC748D96A2176F4E638148836566218C63F7ACD0E816AEA26A2C0E59A43EBF17E4A8A47D50BE0932B9D48C54148AE8C303
                            10DDC0000F23FCCFA5E33B8436549F570D45EB708D9ED960841DEDA11511733F0C04303BDAB81C921FAF71EF0E9F9B20CDE012034AB14D8CB9FE54D4
                            B4CCE8729D551387214725F9E2454FD40B1CF775FA5492A79E6F9DA4FA0F5E3899F899609DF97047CB1D728E392301E855E7A93F79B5074816EE9014
                            3B213ED9A3A1E5F0F44289EC010034E435241B92513FF36D2B14F5EB153A8DF501C178632D2BA31F15BE8A1A021C57E31B587D9924B5252AA300CD2E
                            D1EBAEFADD659398A7C96704648511223D7776E2B3C7381B3222B89AA40A281900031E81ABCBDEEEE4B068CF45EC8F91C5B413B177F4D5A2E9BCEB89
                            399E4CC3E5E947A68ADF6EF5499C7220B3755BA5A5C21DF2A5C61D834AB05D9D5AAC6149AC8B3CB80640FC28CA6E3A3F8FB73E1C690B14701E25D06B
                            D7F2CEE096335004DD80762C9514118496CC09CF0EBB9C9954EE9E0307C9BF7085FC474E12F66F21DCA3ABEB2222941F16B2F7622895789E82E1C3DA
                            3574AF6D1757BDD07A02894241E7610AE1582C54E58A50C973376ADF42A8AD50E9166092B5A4DB658DD4098234225D6CD201E491C3BE5F1757BAACB9
                            82ED7699B33CFA05E29BDDAB5DCE3F08401AA45DC4C29BE5D99697C60398A00B869C748F7FF0E9EC7FB8B79177F8F140E024780FB6CC256FD028EFAD
                            314E914B9169D4113C9F1621D1837A832580CDC639FB1A4F93E8F2865920BDBE63597B2299D3DD75E622F82CB14F99B3B624695082F6928F821AF618
                            8303FE92CAEECD16AC4C4215735CAF98B5E666F5D09F536D13662AA6C2900CB5FA90C65FE8637B8A89A1ED65D10E00856FE947874EA3080766E2C886
                            6CD595E3BDE617AE01A9653EBCD05BBB048AA973A91593904013BF48062E549D78A84A34193B797E3B817BC223C09AC860E33C6AB14E2D7C79B81A21
                            CBCD0803C98B29942317E9D70A2402D7595A587E43380B5B01688905F61A5A591F7F72CA303470B8C49E48279BD3DF756C0CDB777B7E84448AEF81DD
                            B4AB5A4D7F323F10A67B5D5EE401E1FACBF2AC582DE2814996A9EDDC99F5B8F3B1EBED3BC909E536CB16E45B8A5A287F3ECB0C308A98957143813E46
                            7B2FF1FF6001D0E970413E74AA12F90AA471B63C8209523C68AAA02876220731ABC2B3D9C43E1ADAF35EEA3E7F916A1A583999FDC3551BE7323FDC85
                            C722E4E2223685F44369FF2395ED246AF2FD494CA8BCA0EFAA6AD3E1DDA05AACD4C79D22B64843DF8333925135BD541E9AE7E6154482D253C115C2BE
                            2121F26423EFE1E6D1BA486614E3C2DA880DFF576F6F701E816C8F9F35659A857F54C7ED9D153387263BA1098AE5911302245496E2AAFCDA6236A5D9
                            B0E0FBB4B7DBA8EB48B6D295DC5CF1948260EC516028985849968F6EFA492D441EB3BA1F48C165FB50360172A8ACD8687AC387481E67F5F490AE2F13
                            A83CAD1DFB3D4B3506C3739B5F80CE065EBFDCEE74033E8CA6B0913CA02FFFF3C8B7B29806A6A1EEDE0FF19CE27851659279325EC77571A1737678F2
                            E6FC371AD6FA08A4833BFAF8351B48B12C444D303AD3DEA14327390BDD67C2D55B5543D9FC319B6E56E5EEADCD69A303EFC8331F84C2122412D88034
                            0233F52223ACC3D4AAD08A6160241EE6307A332BD81E4CF817CFAC371E713CFF3FE1938783C4EE6566893B8EA50F
                        --%>
                        <tr>
                            <td class="bluecol_need">角色 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DDL_Role" runat="server"></asp:DropDownList>
                                <asp:RequiredFieldValidator ID="msttype" runat="server" ErrorMessage="請選擇角色" Display="None" ControlToValidate="DDL_Role"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtname" runat="server" MaxLength="50" Width="40%"></asp:TextBox><asp:RequiredFieldValidator ID="mstname" runat="server" ErrorMessage="請輸入姓名" Display="None" ControlToValidate="txtname"></asp:RequiredFieldValidator></td>
                        </tr>
                        <tr id="tr_IDNO" runat="server">
                            <td class="bluecol_need">身分證號碼 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" MaxLength="11" Width="20%"></asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="檢查身分證" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                                <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="身分證號碼錯誤!" Display="None" ControlToValidate="IDNO" ClientValidationFunction="checkIDNO"></asp:CustomValidator>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="請輸入身分證號碼" Display="None" ControlToValidate="IDNO"></asp:RequiredFieldValidator>
                                <asp:Label ID="reidno" runat="server" ForeColor="Red" Visible="False"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">電話 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="telphone" runat="server" MaxLength="25" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">E_Mail </td>
                            <td class="whitecol">
                                <asp:TextBox ID="email" runat="server" MaxLength="64" Width="50%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="emailrfv1" runat="server" ErrorMessage="請輸入EMAIL" Display="None" ControlToValidate="email"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="mailchk" runat="server" ErrorMessage="請填寫正確的EMAIL" Display="None" ControlToValidate="email" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                                <asp:Button ID="BtnReEmail1" runat="server" Text="重設E-MAIL" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">跨區支援 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblDistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="3"></asp:CheckBoxList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">帳號停用日 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="StopDate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                <span id="span1" runat="server">
                                    <img id="IMG1" style="cursor: pointer" onclick="openCalendar('StopDate','2000/1/1','2100/1/1',document.getElementById('nowdate').value,'','');" alt="" src="../../images/show-calendar.gif" align="absMiddle" runat="server" width="30" height="30"></span>
                                <%--<input id="nowdate" type="hidden" runat="server" size="1" />--%>
                                <%--<asp:Button ID="Button4" Style="display: none" runat="server" Text="顯示課程" CssClass="asp_Export_M" CausesValidation="False"></asp:Button>--%>
                                <asp:Button ID="bt_clearDate" runat="server" Text="清除" CssClass="asp_Export_M" CausesValidation="False"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">啟用 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="IsUsed" runat="server" Checked="True"></asp:CheckBox>
                                <asp:CustomValidator ID="Customvalidator2" runat="server" ClientValidationFunction="CheckData" ErrorMessage="CustomValidator" Display="None"></asp:CustomValidator>
                                <asp:Button ID="BtnReUsed1" runat="server" Text="重啟帳號" CssClass="asp_Export_M" />
                                <asp:Button ID="BtnSendPXDEMAIL" runat="server" CssClass="asp_Export_M" Text="寄送密碼函-使用者重設密碼" />
                                <asp:Button ID="BtnrResetPXD" runat="server" Text="(管理使用)重設密碼" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">登入方式 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="chkLoginWay1" runat="server" Checked="True" Enabled="False" Text="自然人憑證登入" />
                                &nbsp;&nbsp;<asp:CheckBox ID="chkLoginWay2" runat="server" Text="帳號/密碼登入" Enabled="False" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">最後登入時間</td>
                            <td class="whitecol">
                                <asp:Label ID="lab_LastDATE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">&nbsp;
                                <asp:Button ID="Button2" runat="server" Visible="False" Text="清除身分證與憑證序號" CssClass="asp_Export_M" CausesValidation="False" />&nbsp;
                                <asp:Button ID="update11" runat="server" Visible="False" Text="清除自然人憑證序號" CssClass="asp_Export_M" CausesValidation="False" />&nbsp;
                                <asp:Button ID="btu_save" runat="server" Text="儲存" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                <%--<input id="back" type="button" value="回上一頁" name="back" runat="server" class="button_b_M">--%>
                                <asp:Button ID="btn_back1" runat="server" Text="回上一頁" CssClass="asp_Export_M" CausesValidation="False" />&nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="left" class="whitecol"></td>
                            <td valign="top" align="left" class="whitecol">
                                <asp:Label ID="Label1" runat="server" ForeColor="Red">清除身分證與憑證序號：</asp:Label>
                                <asp:Label ID="Label2" runat="server" ForeColor="Blue">因系統只接受一組身分證號用於一單位帳號，若有變更單位帳號等情況，則可使用此功能(將連自然人憑證一起清除，且狀態會設為不啟用)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="left" class="whitecol"></td>
                            <td valign="top" align="left" class="whitecol">
                                <asp:Label ID="Label3" runat="server" ForeColor="Red">清除自然人憑證序號：</asp:Label>
                                <asp:Label ID="Label4" runat="server" ForeColor="Blue">若因第一次選錯使用帳號或是換自然人憑證序號等，可使用此功能(身分證號並不清除)</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:ValidationSummary ID="totalmsg" runat="server" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>
        <%--<input id="Hidden1" type="hidden" value="samename" runat="server" />--%>
        <input id="RIDValue" type="hidden" runat="server" />
        <input id="PlanIDValue" type="hidden" runat="server" />
        <input id="winreload" type="hidden" runat="server" />
        <input id="orgname" type="hidden" runat="server" />
        <input id="isBlack" type="hidden" runat="server" />
        <asp:HiddenField ID="hidOrgID" runat="server" />
        <asp:HiddenField ID="HidsAccount" runat="server" />
    </form>
</body>
</html>
