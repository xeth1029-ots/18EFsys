<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LevPlan.aspx.vb" Inherits="WDAIIP.LevPlan" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>LevPlan</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <script type="text/javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function returnValue(rid, planid, orgname, orgid, isBlack) {
            document.form1.OrgName.value = orgname;
            document.form1.OrgRID.value = rid;
            if (opener.document.getElementById('RIDValue')) {
                opener.document.form1.RIDValue.value = rid;
            }
            if (opener.document.getElementById('PlanIDValue')) {
                opener.document.form1.PlanIDValue.value = planid;
            }
            if (opener.document.getElementById('TBplan')) {
                opener.document.form1.TBplan.value = orgname;
            }
            //2005/6/13--新增回傳orgid-Melody //alert(orgid);
            if (opener.document.getElementById('orgid_value')) {
                opener.document.getElementById('orgid_value').value = orgid;
                //alert(opener.document.getElementById('orgid_value').value);
            }
            if (getParamValue('fRID') != '')
                opener.document.getElementById(getParamValue('fRID')).value = rid;
            if (getParamValue('fPlanID') != '')
                opener.document.getElementById(getParamValue('fPlanID')).value = planid;
            if (getParamValue('fTBplan') != '')
                opener.document.getElementById(getParamValue('fTBplan')).value = orgname;
            if (getParamValue('fOrgID') != '')
                opener.document.getElementById(getParamValue('fOrgID')).value = orgid;
            //if (getParamValue('fOrgID') != '')
            //    opener.document.getElementById(getParamValue('fOrgID')).value = orgid;
            if (getParamValue('fisBlack') != '')
                opener.document.getElementById(getParamValue('fisBlack')).value = isBlack;
            //if (getParamValue('fisBlack') != '')
            //    opener.document.getElementById(getParamValue('fisBlack')).value = isBlack;
            if (orgname.indexOf("_") > -1) {
                if (getParamValue('OrgField') != '') {
                    if (opener.document.getElementById(getParamValue('OrgField')))
                        opener.document.getElementById(getParamValue('OrgField')).value = orgname.split("_")[1];
                }
            }
            //2004/1/5------增加母頁reload功能(小圭)
            if (getParamValue('winreload') == "1") { opener.document.form1.submit(); }

            document.getElementById('Button1').click();
            window.close();
        }
    </script>
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
    <link href="../css/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="org_level" type="hidden" name="org_level" value="0" runat="server" size="1" />
        <input id="hiddenlevel" type="hidden" name="hiddenlevel" runat="server" size="1" />
        <input id="OrgName" type="hidden" name="OrgName" runat="server" />
        <input id="OrgRID" type="hidden" name="OrgRID" runat="server" />
        <table class="table_nw" id="Table1" cellspacing="1" cellpadding="1" border="0" runat="server" width="100%">
            <tr>
                <td class="bluecol" width="20%">年度：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True"></asp:DropDownList><asp:CheckBox ID="cbplanlistAll" runat="server" AutoPostBack="True" Text="不排除其它計畫"></asp:CheckBox></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">計畫代碼：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="planlist" runat="server" AutoPostBack="True"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">關鍵字：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtSearch1" runat="server" Width="60%" MaxLength="300"></asp:TextBox>
                    <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>

        </table>
        <div style="overflow-y: auto; height: 400px;">
            <asp:TreeView ID="TreeView1" runat="server"></asp:TreeView>
        </div>
        <%--<iewc:treeview id="TreeView1" runat="server" expandlevel="1"></iewc:treeview>--%>
        <asp:Button ID="Button1" Style="display: none" runat="server" Text="存入暫存"></asp:Button>
        <asp:HiddenField ID="hid_ORGBLACKLIST" runat="server" />
    </form>
</body>
</html>