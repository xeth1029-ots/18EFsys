<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LevOrg.aspx.vb" Inherits="WDAIIP.LevOrg" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>請選擇機構</title>
    <meta content="javascript" name="vs_defaultclientscript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
    <link href="../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function Search_click() {
            if (document.getElementById('txtSearch1')) {
                if (event.keyCode == 13) {
                    document.getElementById('btnSearch').disabled = true;
                    document.getElementById('hbtnSearch').click();
                }
            }
        }

        function Search_click2() {
            if (document.getElementById('btnSearch'))
                document.getElementById('btnSearch').disabled = true;
            if (document.getElementById('hbtnSearch'))
                document.getElementById('hbtnSearch').click();
        }

        function returnValue(rid, orgname, comidno, EMail, orgid, submit, CTName, ZipCode, Address, OrgLevel, isBlack, DistID, PlanID, btnbtn) {
            var Hidnorus = document.getElementById('Hidnorus');
            if (Hidnorus.value == "1") {
                window.opener = null; window.open('', '_self'); window.close();
                return "";
            }
            Hidnorus.value = "1";
            if (getParamValue('fisBlack') != '') {
                if (opener.document.getElementById(getParamValue('fisBlack')))
                    opener.document.getElementById(getParamValue('fisBlack')).value = isBlack;
            }
            if (getParamValue('OrgField') != '') {
                if (opener.document.getElementById(getParamValue('OrgField')))
                    opener.document.getElementById(getParamValue('OrgField')).value = orgname;
            }
            if (getParamValue('PlanID') == '') {
                if (document.form1.OrgName)
                    document.form1.OrgName.value = orgname;
                if (document.form1.OrgRID)
                    document.form1.OrgRID.value = rid;
                if (getParamValue('RIDField') == '') {
                    if (opener.document.form1.RIDValue)
                        opener.document.form1.RIDValue.value = rid;
                    //20180709
                    if (opener.document.form1.RIDValueX1)
                        opener.document.form1.RIDValueX1.value = rid;
                }
                else {
                    if (opener.document.getElementById(getParamValue('RIDField')))
                        opener.document.getElementById(getParamValue('RIDField')).value = rid;
                }
                if (getParamValue('OrgField') == '') {
                    if (opener.document.form1.center)
                        opener.document.form1.center.value = orgname;
                    //20180709
                    if (opener.document.form1.centerX1)
                        opener.document.form1.centerX1.value = orgname;
                }
                else {
                    if (opener.document.getElementById(getParamValue('OrgField')))
                        opener.document.getElementById(getParamValue('OrgField')).value = orgname;
                }
                if (opener.document.getElementById('ComidValue'))
                    opener.document.getElementById('ComidValue').value = comidno;
                //20180709
                if (opener.document.getElementById('txt_ComIDNO'))
                    opener.document.getElementById('txt_ComIDNO').value = comidno;
                //2005/6/13--新增回傳orgid-Melody
                if (opener.document.getElementById('orgid_value'))
                    opener.document.getElementById('orgid_value').value = orgid;
                //20180709
                if (opener.document.getElementById('orgid_valueX1'))
                    opener.document.getElementById('orgid_valueX1').value = orgid;
                if (opener.document.getElementById('orgid_Level'))
                    opener.document.getElementById('orgid_Level').value = OrgLevel;
                if (opener.document.getElementById('EMail'))
                    opener.document.getElementById('EMail').value = EMail;
                //增加回傳地址
                if (opener.document.getElementById('CTName')) opener.document.getElementById('CTName').value = CTName;
                if (opener.document.getElementById('TaddressZip')) opener.document.getElementById('TaddressZip').value = ZipCode;
                if (opener.document.getElementById('TAddress')) opener.document.getElementById('TAddress').value = Address;
                //母頁按鈕
                if (opener.document.getElementById(btnbtn) != null) opener.document.getElementById(btnbtn).click();
                //if (opener.document.getElementById(btnbtn) != null) { alert(btnbtn + ' is not null'); } else { alert(btnbtn + ' is null'); }
                if (submit == 'true') {
                    if (opener.document.getElementById('but_search')) {
                        opener.document.form1.but_search.click();
                    }
                    else {
                        alert('查無指定按鈕功能，請重新整理網頁\n若持續出現此問題，請連絡系統管理人員!!謝謝');
                        window.opener = null; window.open('', '_self'); window.close();
                    }
                }
            }
            else {
                if (opener.document.form1.OrgRID)
                    opener.document.form1.OrgRID.value = rid;
                if (opener.document.form1.OrgName)
                    opener.document.form1.OrgName.value = orgname;
                if (opener.document.getElementById('ComidValue'))
                    opener.document.getElementById('ComidValue').value = comidno;
                //20180709
                if (opener.document.getElementById('txt_ComIDNO'))
                    opener.document.getElementById('txt_ComIDNO').value = comidno;
                if (opener.document.getElementById('EMail')) {
                    opener.document.getElementById('EMail').value = EMail;
                }
            }
            //add 回傳所屬轄區代碼
            var hidDistID = opener.document.getElementById("hidDistID");
            if (hidDistID) {
                hidDistID.value = DistID;
            }
            //add 回傳計畫代碼
            var hidPlanID = opener.document.getElementById("hidPlanID");
            if (hidPlanID) {
                hidPlanID.value = PlanID;
            }
            //回傳值(限定本頁查詢功能按鈕Button1)
            document.getElementById('Button1').click();
            //window.close();
        }

        function GetRID() {
            if (document.form1.DistID) {
                switch (document.form1.DistID.selectedIndex) {
                    case 1:
                        document.form1.RID.value = 'B';
                        break;
                    case 2:
                        document.form1.RID.value = 'C';
                        break;
                    case 3:
                        document.form1.RID.value = 'D';
                        break;
                    case 4:
                        document.form1.RID.value = 'E';
                        break;
                    case 5:
                        document.form1.RID.value = 'F';
                        break;
                    case 6:
                        document.form1.RID.value = 'G';
                        break;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="table_sch" id="supershow" cellspacing="1" cellpadding="1" border="0" runat="server" width="100%">
            <tr>
                <td class="bluecol" style="text-align: left" width="25%">年度：</td>
                <td class="whitecol" width="75%">
                    <asp:DropDownList ID="Downyear" runat="server" AutoPostBack="true"></asp:DropDownList></td>
            </tr>
            <tr>
                <td style="text-align: left;" class="bluecol" width="25%">轄區：</td>
                <td class="whitecol" width="75%">
                    <asp:DropDownList ID="DistID" AutoPostBack="true" runat="server"></asp:DropDownList>
                    <input id="RID" type="hidden" name="RID" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="text-align: left" width="25%">計畫代碼：</td>
                <td class="whitecol" width="75%">
                    <asp:DropDownList ID="planlist" AutoPostBack="true" runat="server" Width="80%"></asp:DropDownList></td>
            </tr>
        </table>
        <table class="table_nw" id="TableSearch" cellspacing="1" cellpadding="1" border="0" runat="server">
            <tr>
                <td class="bluecol" width="25%">關鍵字： </td>
                <td class="whitecol">
                    <asp:TextBox ID="txtSearch1" runat="server" Width="50%" MaxLength="1000"></asp:TextBox>
                    <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <div style="overflow-y: auto; height: 410px;">
            <table class="font">
                <tr>
                    <td>
                        <asp:TreeView ID="TreeView1" runat="server">
                            <NodeStyle ForeColor="#0066FF" />
                            <ParentNodeStyle ForeColor="#0066FF" />
                            <RootNodeStyle ForeColor="#0066FF" />
                        </asp:TreeView>
                    </td>
                </tr>
            </table>
            <asp:Button ID="Button1" Style="display: none" runat="server" Text="存入暫存" CssClass="asp_button_M"></asp:Button>
            <input id="OrgName" type="hidden" name="OrgName" runat="server" />
            <input id="OrgRID" type="hidden" name="OrgRID" runat="server" />
            <input id="hidname" type="hidden" name="hidname" runat="server" />
            <input id="hidbtnName" type="hidden" name="hidbtnName" runat="server" />
            <input type="button" value="搜尋" id="hbtnSearch" runat="server" style="display: none" />
            <input id="Hidnorus" type="hidden" runat="server" />
        </div>
        <asp:HiddenField ID="hid_ORGBLACKLIST" runat="server" />
    </form>
</body>
</html>
