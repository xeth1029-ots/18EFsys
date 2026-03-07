<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_004_add.aspx.vb" Inherits="WDAIIP.SD_02_004_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>甄試通知單設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function checkData() {
            var msg = "";
            var tmpContent = document.getElementById("Notice").value.replace(/\r\n/g, "");

            //alert('length=['+document.getElementById("eComment").value.length+']\n length=['+tmpContent.length+']');
            if (document.getElementById("Notice").value.length > 2000) {
                msg += '「說明事項」欄位 ' + document.getElementById("Notice").value.length + ' 個字元超過欄位 2000 個字元最大限制!\n';
            }

            if (msg != "") {
                alert(msg);
                return false;
            } else {
                return true;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td align="center">
                        <table class="table_sch" id="table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <%--<tr>
							<td colspan="4">
								首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;<font color="#990000"> 甄試通知單設定</font>
							</td>
						</tr>--%>
                            <tr>
                                <td id="td_dist" colspan="4">轄區：<asp:Label ID="DistID" runat="server" CssClass="font"></asp:Label>
                                </td>
                            </tr>
                            <tr id="TR_PlanYear" runat="server">
                                <td id="Td_PlanYear" colspan="4" runat="server">年度：<asp:Label ID="PlanYear" runat="server" CssClass="font"></asp:Label>
                                </td>
                            </tr>
                            <tr id="TR_CtrlOrg" runat="server">
                                <td id="TD_CtrlOrg" colspan="4" runat="server">管控單位：<asp:Label ID="CtrlOrg" runat="server" CssClass="font"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td id="TD_org" colspan="4" runat="server">訓練機構：<asp:Label ID="OrgName" runat="server" CssClass="font"></asp:Label>
                                    <input id="OrgID" style="width: 56px; height: 22px" type="hidden" size="4" name="OrgID" runat="server" />&nbsp;
								<input id="RIDValue" style="width: 56px; height: 22px" type="hidden" size="4" name="RIDValue" runat="server" />
                                    <input id="OCID" style="width: 56px; height: 22px" type="hidden" size="4" name="OCID" runat="server" />
                                    <input id="ParentRID" style="width: 56px; height: 22px" type="hidden" size="4" name="ParentRID" runat="server" />
                                    <input id="ParentOrgID" style="width: 56px; height: 22px" type="hidden" size="4" name="ParentOrgID" runat="server" />
                                    <input id="NoticeFrom" style="width: 56px; height: 22px" type="hidden" size="4" name="NoticeFrom" runat="server" />
                                    <asp:HiddenField ID="hidDistID" runat="server" />
                                    <asp:HiddenField ID="hidPlanID" runat="server" />
                                </td>
                            </tr>
                            <tr id="TR_Class" runat="server">
                                <td id="TD_Class" style="height: 18px" colspan="4" runat="server">班級名稱：<asp:Label ID="ClassName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" align="center">甄試通知單<br>
                                    說明事項：
                                </td>
                                <td colspan="3">
                                    <textarea id="Notice" style="width: 500px; height: 200px" rows="5" cols="60" runat="server"></textarea>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4" class="whitecol">
                                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
								<asp:Button ID="btnBack" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table>
    </form>
</body>
</html>
