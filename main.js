function getUrlPath() {
    let e = document.location.toString();
    return -1 != (e = decodeURI(e)).indexOf("/") && (e = e.substring(0, e.lastIndexOf("/"))), e;
}

function onAction(control) {
    switch (control.Id) {
        case "default-button":
            // window.Application.ShowDialog(getUrlPath() + "/ui/dialog.html", "这是一个对话框网页", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false);
            alert("Hello, World!");
            break;
        default:
            return true;
    }
}

function getImage(control) {
    switch (control.Id) {

        default:
            return "images/error-picture.svg";
    }
}