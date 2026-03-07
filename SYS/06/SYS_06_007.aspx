<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_007.aspx.vb" Inherits="WDAIIP.SYS_06_007" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    擐?&gt;&gt;蝟餌絞蝞∠?&gt;&gt;<font color="#990000">蝟餌絞瘚?閮剖?</font>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="table2" cellspacing="1" cellpadding="1" width="740">
            <tr>
                <td class="bluecol" width="100">
                    閮毀閮
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="planlist" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    蝞∠?摮頂蝯?
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="dopParentNode" runat="server" AutoPostBack="true">
                        <asp:ListItem Value="">隢??/asp:ListItem>
                        <asp:ListItem Value="TC">閮毀璈?蝞∠?</asp:ListItem>
                        <asp:ListItem Value="SD">摮詨??蝞∠?</asp:ListItem>
                        <asp:ListItem Value="CP">?交/蝮暹?蝞∠?</asp:ListItem>
                        <asp:ListItem Value="TR">閮毀?瘙恣??/asp:ListItem>
                        <asp:ListItem Value="CM">閮毀蝬祥?抒恣</asp:ListItem>
                        <asp:ListItem Value="SYS">蝟餌絞蝞∠?</asp:ListItem>
                        <asp:ListItem Value="FAQ">????/asp:ListItem>
                        <asp:ListItem Value="OB">憪?閮毀蝞∠?</asp:ListItem>
                        <asp:ListItem Value="SE">??賣炎摰恣??/asp:ListItem>
                        <asp:ListItem Value="EXAM">?岫蝞∠?</asp:ListItem>
                        <asp:ListItem Value="SV">?蝞∠?</asp:ListItem>
                        <asp:ListItem Value="OO">?嗡?蝟餌絞</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;
                    <asp:DropDownList ID="dopNextNode" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
        <asp:Panel ID="panEdit" runat="server" Width="740" HorizontalAlign="Center">
            <table width="740" class="table_nw" cellpadding="1" cellspacing="1">
                <tr>
                    <td width="100" class="bluecol">
                        ????
                    </td>
                    <td class="whitecol">
                        <div id="drag">
                            <asp:Table ID="tbDataTable" runat="server">
                            </asp:Table>
                        </div>
                        &nbsp;
                        <input id="txtFunID" runat="server" style="display: none;" />
                    </td>
                </tr>
            </table>
            <asp:Button ID="btnSave" runat="server" Text="?脣?" CssClass="asp_button_S" OnClientClick="return checkSubmit();" />
            <br />
            <br />
            <table width="740" class="font">
                <tr>
                    <td colspan="2" class="bluecol">
                        ?瑁??恍?汗
                    </td>
                </tr>
                <tr>
                    <td width="100">
                    </td>
                    <td>
                        <asp:TreeView ID="TreeView1" runat="server" Width="250px">
                            <NodeStyle ForeColor="#0066FF" />
                            <ParentNodeStyle ForeColor="#0066FF" />
                            <RootNodeStyle ForeColor="#0066FF" />
                        </asp:TreeView>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </div>
    </form>
    <script language="JavaScript" type="text/javascript">
    <!--
        function checkSubmit() {
            var planlist = document.getElementById("planlist");
            var dopParentNode = document.getElementById("dopParentNode");
            var txtFunID = document.getElementById("txtFunID");
            var tbDataTable = document.getElementById("tbDataTable");
            var msg = '';
            var tmp = '';
            //alert(tbDataTable.rows.length);
            if (planlist.value == '') {
                msg = '隢????;
            }
            if (dopParentNode.value == '') {
                msg += '隢??蝟餌絞';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                for (var i = 0; i < tbDataTable.rows.length; i++) {

                    tmp += tbDataTable.rows[i].cells[1].firstChild.value + ',';
                }
                txtFunID.value = tmp;
                //alert(txtFunID.value);
                //return false;
            }
        }

        function cleanWhitespace(element) {

            for (var i = 0; i < element.childNodes.length; i++) {
                var node = element.childNodes[i];

                if (node.nodeType == 3 && !/S/.test(node.nodeValue))
                    node.parentNode.removeChild(node);
            }
        }

        var _table = document.getElementById("table1");
        cleanWhitespace(_table);

        function moveUp(_a) {

            var _row = _a.parentNode.parentNode;

            if (_row.previousSibling) swapNode(_row, _row.previousSibling);
        }

        function moveDown(_a) {

            var _row = _a.parentNode.parentNode;

            if (_row.nextSibling) swapNode(_row, _row.nextSibling);
        }

        function swapNode(node1, node2) {

            var _parent = node1.parentNode;

            var _t1 = node1.nextSibling;
            var _t2 = node2.nextSibling;

            if (_t1) _parent.insertBefore(node2, _t1);
            else _parent.appendChild(node2);

            if (_t2) _parent.insertBefore(node1, _t2);
            else _parent.appendChild(node1);
        }
    //-->
    </script>
</body>
</html>
