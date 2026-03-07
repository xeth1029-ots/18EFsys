<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ProSkill.aspx.vb" Inherits="WDAIIP.ProSkill" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>ProSkill</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <style type="text/css">
        A { font-size: 10pt; font-family: "新細明體"; }
            A:visited { color: #0000ff; }
            A:link { color: #0000ff; }
            A:hover { color: #9933cc; }
            A:active { color: #9933cc; }
    </style>
    <script type="text/javascript" language="javascript">
        function getParamValue(name) {
            var querystring;
            var values;
            var result = "";
            if (location.search.length > 1) {
                querystring = unescape(location.search + "&");
                var re = new RegExp("[\?|\&]" + name + "=(.+)\&");
                values = querystring.match(re);
                if (values != null) {
                    result = values[1];
                    if (result.indexOf("&") != -1) {
                        result = result.substring(0, result.indexOf("&"));
                    }
                }
            }
            return result;
        }

        function showDetailTable(KPID, ProID, ProName) {
            document.getElementById("MainID").value = KPID;
            document.getElementById('HidMainNum').value = ProID;
            document.getElementById('HidMainName').value = ProName;
            document.form1.submit();
        }

        function return_value(KPID, ProID, ProName) {
            var myfield;
            myfield = getParamValue("Skill_Field");
            if (myfield != "") {
                opener.document.getElementById(myfield).value = KPID;
            }
            myfield = getParamValue("Skill_Name_Field");
            if (myfield != "") {
                if (ProID == '')
                    opener.document.getElementById(myfield).value = '';
                else
                    opener.document.getElementById(myfield).value = '【' + ProID + '】' + ProName;
            }
            window.close();
        }

        function showThridTable(KPID, ProID, ProName) {
            document.getElementById('SecID').value = KPID;
            document.getElementById('HidSecNum').value = ProID;
            document.getElementById('HidSecName').value = ProName;
            document.form1.submit();
        }
    </script>
</head>
<body bgcolor="#e6efff">
    <form id="form1" method="post" runat="server">
        <table id="MainBlock" bordercolor="#0000cc" width="90%" align="center" border="2" runat="server">
            <tr>
                <td>
                    <table width="100%" align="center">
                        <tr>
                            <td align="center"><strong><font face="標楷體" color="#000066" size="4">專 業 技 能 分 類</font></strong></td>
                        </tr>
                        <tr>
                            <td align="center"><font color="#ff0000" size="2">※請選擇專業技能大類，再細分小類 ※</font></td>
                        </tr>
                    </table>
                    <table id="MainList" cellspacing="2" cellpadding="6" width="100%" align="center" border="0" runat="server">
                    </table>
                    <div align="center">【<a title="關閉視窗" href="javascript:window.close();">關閉</a>】&nbsp; 【<a title="清除工作職業分類" href="javascript:return_value('','','');">清除</a>】</div>
                </td>
            </tr>
        </table>
        <table id="SubBlock" bordercolor="#0000cc" width="98%" align="center" border="2" runat="server">
            <tr>
                <td>
                    <table width="100%" align="center">
                        <tr>
                            <td align="center"><strong><font face="標楷體" color="#000066" size="4"><strong><font face="標楷體" color="#000066" size="4"><strong><font face="標楷體" color="#000066" size="4">【<asp:Label ID="MainNum" runat="server"></asp:Label>】</font></strong></font></strong><asp:Label ID="MainName" runat="server"></asp:Label></font></strong></td>
                        </tr>
                    </table>
                    <table id="SubList" cellspacing="2" cellpadding="6" width="100%" align="center" border="0" runat="server">
                    </table>
                    <div align="center">【<a title="關閉視窗" href="javascript:window.close();">關閉</a>】&nbsp; 【<a title="清除縣市鄉鎮" href="javascript:return_value('','','');">清除</a>】&nbsp;【<a title="返回上一層選單" href="javascript:history.go(-1);">回上頁</a>】</div>
                </td>
            </tr>
        </table>
        <table id="ThirdBlock" bordercolor="#0000cc" width="98%" align="center" border="2" runat="server">
            <tr>
                <td>
                    <table width="100%" align="center">
                        <tr>
                            <td align="center"><strong><font face="標楷體" color="#000066" size="4"><strong><font face="標楷體" color="#000066" size="4"><strong><font face="標楷體" color="#000066" size="4">【<asp:Label ID="SecNum" runat="server"></asp:Label>】</font></strong></font></strong><asp:Label ID="SecName" runat="server"></asp:Label></font></strong></td>
                        </tr>
                    </table>
                    <table id="ThirdList" cellspacing="2" cellpadding="6" width="100%" align="center" border="0" runat="server">
                    </table>
                    <div align="center">【<a title="關閉視窗" href="javascript:window.close();">關閉</a>】&nbsp; 【<a title="清除縣市鄉鎮" href="javascript:return_value('','','');">清除</a>】&nbsp;【<a title="返回上一層選單" href="javascript:history.go(-1);">回上頁</a>】</div>
                </td>
            </tr>
        </table>
        <input id="MainID" type="hidden" runat="server"><input id="SecID" type="hidden" runat="server"><input id="HidMainNum" type="hidden" runat="server"><input id="HidSecName" type="hidden" runat="server"><input id="HidSecNum" type="hidden" runat="server"><input id="HidMainName" type="hidden" runat="server">
    </form>
</body>
</html>