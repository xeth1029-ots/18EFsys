
function debugTrace(msg) {
    console.log(msg);
}

/* 覆寫原本的 javascript alert() */
window.alert = function (msg) {
    debugTrace("(override)alert: " + msg);
    if (parent) {
        parent.blockMessage(msg);
    }
    else {
        debugTrace("Iframe parent 不存在!!");
    }
}

if (parent) {
    /* 重設主頁(index.aspx)的 Timer 計數*/
    if (parent.resetTimer) {
        parent.resetTimer();
    }
}

function blockAlert(msg, title, unblockCallback) {
    if (parent) {
        parent.blockAlert(msg, title, unblockCallback);
    }
    else {
        debugTrace("parent.blockAlert() NOT exists.");
    }
}

function blockMessage(msg, title, unblockCallback) {
    if (parent) {
        parent.blockMessage(msg, title, unblockCallback);
    }
    else {
        debugTrace("parent.blockMessage() NOT exists.");
    }
}

function blockConfirm(msg, title, confirmCallback, cancelCallback) {
    if (parent) {
        parent.blockConfirm(msg, title, confirmCallback, cancelCallback);
    }
    else {
        debugTrace("parent.blockConfirm() NOT exists.");
    }
}