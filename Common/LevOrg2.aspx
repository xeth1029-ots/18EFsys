<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LevOrg2.aspx.vb" Inherits="WDAIIP.LevOrg2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇機構</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
    <link href="../css/css.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function GetValue(OrgName, RID) {
            opener.document.getElementById('center').value = OrgName;
            opener.document.getElementById('RIDValue').value = RID;
            var btn = getParamValue('BtnName');
            if (btn != '') {
                opener.document.getElementById(btn).click();
            }
            else {
                if (opener.document.getElementById('Button2'))
                    opener.document.getElementById('Button2').click();
            }
            window.close();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 410px;">
            <table class="font">
                <tr>
                    <td>
                        <asp:TreeView ID="TreeView1" runat="server">
                            <NodeStyle ForeColor="#0066FF" />
                            <ParentNodeStyle ForeColor="#0066FF" />
                            <RootNodeStyle ForeColor="#0066FF" />
                        </asp:TreeView>
                        <%--<iewc:TreeView id="TreeView1" runat="server" ExpandLevel="1"></iewc:TreeView>--%>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>