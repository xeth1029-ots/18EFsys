
/* CSS defined in Style/jquery.dataTables.css */
var paging_html = "<div id='listDataPaging_info' class='dataTables_info'></div>"
    + "<div id='listDataPaging_paginate' class='dataTables_paginate paging_full_numbers'>"
    + "<a id='listDataPaging_first' class='first paginate_button paginate_button_disabled' tabindex='0'>首頁</a>"
    + "<a id='listDataPaging_previous' class='previous paginate_button' tabindex='0'>上一頁</a>"
    + "<span></span>"
    + "<a id='listDataPaging_next' class='next paginate_button' tabindex='0'>下一頁</a>"
    + "<a id='listDataPaging_last' class='last paginate_button' tabindex='0'>末頁</a>"
    + "</div>";

/**
* 顯示分頁連結
* 傳入:
*   container: 用來顯示分頁連結的 DIV 元件 DOM object
*   pagination: 分頁資訊物件, 需有下列屬性
*     Total: 總筆數
*     Start: 顯示筆數起
*     End: 顯示筆數迄
*     PageSize: 每頁筆數
*     PageIdx: 當前頁次
*     TotalPages: 總頁數
*   loadDataFunc: 用來切換分頁的 callback function name, 這個 function 被呼叫時會收到一個 "頁次" 的參數 
*/
function showPaging(container, pagination, loadDataFunc) {
    if (!container || !pagination) {
        return;
    }

    $(container).html(paging_html);

    if (pagination.Total == "0") {
        $("#listDataPaging_info").html("");
    } else {
        var info = "共 " + pagination.Total + " 筆, 顯示 " + (1 + pagination.Start) + " 到 " + (1 + pagination.End) + " 筆";
        $("#listDataPaging_info").html(info);
    }

    // ==== 處理分頁連結 ====

    if (pagination.Total <= pagination.PageSize) {
        $("#listDataPaging_paginate").hide();
    }
    else {
        // 第一頁
        if (pagination.PageIdx > 5) {
            $("#listDataPaging_first").attr("href", "javascript:" + loadDataFunc + "(1)");
            $("#listDataPaging_first").attr("class", "first paginate_button");
        } else {
            $("#listDataPaging_first").attr("href", "javascript:void(0)");
            $("#listDataPaging_first").attr("class", "first paginate_button paginate_disabled_previous");
        }

        // 上一頁
        if (pagination.PageIdx > 1) {
            $("#listDataPaging_previous").attr("href", "javascript:" + loadDataFunc + "(" + (pagination.PageIdx - 1) + ")");
            $("#listDataPaging_previous").attr("class", "previous paginate_button");
        } else {
            $("#listDataPaging_previous").attr("href", "javascript:void(0)");
            $("#listDataPaging_previous").attr("class", "previous paginate_button paginate_disabled_previous");
        }

        // 下一頁
        if (pagination.PageIdx < pagination.TotalPages) {
            $("#listDataPaging_next").attr("href", "javascript:" + loadDataFunc + "(" + (1 + pagination.PageIdx) + ")");
            $("#listDataPaging_next").attr("class", "next paginate_button");
        } else {
            $("#listDataPaging_next").attr("href", "javascript:void(0)");
            $("#listDataPaging_next").attr("class", "next paginate_button paginate_disable_button");
        }

        // 最後一頁
        if (pagination.PageIdx < pagination.TotalPages) {
            $("#listDataPaging_last").attr("href", "javascript:" + loadDataFunc + "(" + (pagination.TotalPages) + ")");
            $("#listDataPaging_last").attr("class", "last paginate_button");
        } else {
            $("#listDataPaging_last").attr("href", "javascript:void(0)");
            $("#listDataPaging_last").attr("class", "last paginate_button paginate_disable_button");
        }

        // 頁次連結
        var pBegin = Math.floor((pagination.PageIdx - 1) / 5) * 5 + 1;
        var pageStr = "";
        for (i = pBegin; i < pBegin + 5; i++) {
            if (i > pagination.TotalPages) {
                break;
            }
            if (i == pagination.PageIdx) {
                pageStr += "<a class='paginate_active' tabindex='0'>" + i + "</a>";
            } else {
                pageStr += "<a class='paginate_button' tabindex='0' href='javascript:" + loadDataFunc + "(" + i + ")' >" + i + "</a>";
            }
        }

        $("#listDataPaging_paginate span").html(pageStr);
        $("#listDataPaging_paginate").show();
    }

}


