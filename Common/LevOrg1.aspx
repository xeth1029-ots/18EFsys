<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LevOrg1.aspx.vb" Inherits="WDAIIP.LevOrg1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇機構</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <%--<link href="../css/css.css" rel="stylesheet" type="text/css" />--%>
    <link href="../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function returnValue(rid, orgname, btn, orgid, isBlack) {
            //檢查是否是由帳號設定開啟，是 則回傳不同值
            //debugger;
            if (getParamValue('fisBlack') != '') {
                if (opener.document.getElementById(getParamValue('fisBlack')))
                    opener.document.getElementById(getParamValue('fisBlack')).value = isBlack;
            }
            if (getParamValue('OrgField') != '') {
                if (opener.document.getElementById(getParamValue('OrgField')))
                    opener.document.getElementById(getParamValue('OrgField')).value = orgname;
            }
            if (getParamValue('IDSetUp') == 1) {
                if (opener.document.form1.TBplan)
                    opener.document.form1.TBplan.value = orgname;
            }
            else {
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
            }
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
            if (opener.document.getElementById('orgid_value'))
                opener.document.getElementById('orgid_value').value = orgid;
            //20180709
            if (opener.document.getElementById('orgid_valueX1'))
                opener.document.getElementById('orgid_valueX1').value = orgid;
            if (document.getElementById('OrgName'))
                document.getElementById('OrgName').value = orgname;
            if (document.getElementById('OrgRID'))
                document.getElementById('OrgRID').value = rid;
            if (opener.document.getElementById(btn))
                opener.document.getElementById(btn).click();
            if (document.getElementById('Button1'))
                document.getElementById('Button1').click();
        }

        function Search_click() {
            if (document.getElementById('txtSearch1')) {
                if (event.keyCode == 13) {
                    document.getElementById('btnSearch').disabled = true;
                    document.getElementById('hbtnSearch').click();
                }
            }
        }

        function Search_click2() {
            document.getElementById('btnSearch').disabled = true;
            document.getElementById('hbtnSearch').click();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="table_nw" id="TableSearch" cellspacing="1" cellpadding="1" border="0" runat="server">
            <tr>
                <td class="bluecol" width="25%">關鍵字： </td>
                <td class="whitecol">
                    <asp:TextBox ID="txtSearch1" runat="server" MaxLength="1000"></asp:TextBox>
                    <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <br />
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
            <asp:Button ID="Button1" Style="display: none" runat="server" Text="存入暫存"></asp:Button>
            <input id="OrgName" type="hidden" name="OrgName" runat="server">
            <input id="OrgRID" type="hidden" name="OrgRID" runat="server">
            <asp:TextBox ID="TMID1" Style="display: none" runat="server" onfocus="this.blur()"></asp:TextBox>
            <asp:TextBox ID="OCID1" Style="display: none" runat="server" onfocus="this.blur()"></asp:TextBox><br>
            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
            <input id="hbtnSearch" style="display: none" type="button" value="搜尋" name="hbtnSearch" runat="server">
            <input id="hidbtnName" type="hidden" name="hidbtnName" runat="server">
            <%--<asp:HiddenField ID="Hid_OrgName2" runat="server" />--%>
        </div>
    </form>
</body>
</html>