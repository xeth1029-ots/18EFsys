<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MainOrg.aspx.vb" Inherits="WDAIIP.MainOrg" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>MainOrg</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //No Use Button2
        function MainOrgGetValue(OrgName, RID) {
            var MainOrgOrgN = getParamValue('OrgName');
            var MainOrgbtn = getParamValue('BtnName');
            var OrgNVAL = MainOrgOrgN != '' ? MainOrgOrgN : 'center';
            //var center_name = document.getElementById('Hid_org_name');
            if (opener.document.getElementById(OrgNVAL)) { opener.document.getElementById(OrgNVAL).value = OrgName; }

            if (opener.document.getElementById('RIDValue')) opener.document.getElementById('RIDValue').value = RID;

            if (opener.document.getElementById('PlanID')) { opener.document.getElementById('PlanID').value = document.getElementById('PlanIDValue').value; }

            if (MainOrgbtn != '' && opener.document.getElementById(MainOrgbtn)) { opener.document.getElementById(MainOrgbtn).click(); }
            else {
                if (opener.document.getElementById('Button2')) { opener.document.getElementById('Button2').click(); }
            }
            window.opener = null; window.open('', '_self'); window.close();
            //window.close();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div>
            <table id="Table1" cellspacing="1" cellpadding="1" border="0" class="font" width="100%">
                <tr>
                    <td colspan="2">
                        <table class="font" id="PlanTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td class="bluecol" width="20%">計畫：</td>
                                <td class="whitecol" width="80%">
                                    <asp:DropDownList ID="PlanID" runat="server" AutoPostBack="True" Width="80%"></asp:DropDownList></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <div style="max-height: 400px; overflow-y: auto;">
                            <p>
                                <asp:TreeView ID="TreeView1" runat="server"></asp:TreeView>
                                <%--<iewc:TreeView id="TreeView1" runat="server" ExpandLevel="1"></iewc:TreeView>--%>
                            </p>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <div title="請選擇機構">
            <input id="PlanIDValue" type="hidden" runat="server" /></div>
        <%--<asp:HiddenField ID="Hid_org_name" runat="server" />--%>
    </form>
</body>
</html>
