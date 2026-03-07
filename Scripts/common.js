
/* 
* 由傳入的 data object 中出每個欄位(properties) 並填入指定的 form 中 
* form 中的每一個欄位(input)必需有 'name' 屬性
*/
function formFill(formId, data) {
    var form = $('#' + formId);
    //alert(form);

    // clear all field
    form.find('input[type=text]').val("");
    form.find('input[type=hidden]').val("");
    form.find('textarea').val("");
    form.find('select').val("");

    $.each(data,
        function (idx, el) {
            var fel = form.find('*[name="' + idx + '"]');
            var type = "", tag = "";

            if (fel.length > 0) {
                //alert(idx + ", fel.length=" + fel.length);

                tag = fel[0].tagName.toLowerCase();

                if (tag == "select" || tag == "textarea") { //...
                    $(fel).val(el);
                }
                else if (tag == "input") {
                    type = $(fel[0]).attr("type");
                    if (type == undefined) {
                        type = "text";
                    }

                    if (type == "text" || type == "password" || type == "hidden") {
                        //alert(el);
                        fel.val(el);
                    }
                    else if (type == "checkbox") {
                        if (el)
                            fel.attr("checked", "checked");
                    }
                    else if (type == "radio") {
                        fel.filter('[value="' + el + '"]').attr("checked", "checked");
                    }
                }
            }
        })

}