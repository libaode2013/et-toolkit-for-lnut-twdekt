function getUrlPath() {
    let e = document.location.toString();
    return -1 != (e = decodeURI(e)).indexOf("/") && (e = e.substring(0, e.lastIndexOf("/"))), e;
}

function onAction(control) {
    const eleId = control.Id;
    switch (eleId) {
        case "dialog/credit-summary/calculate-progress/version-20260622A":
            window.Application.ShowDialog(`${getUrlPath()}/${eleId}/index.html`, "学分汇总（计算完成进度）", 320 * window.devicePixelRatio, 640 * window.devicePixelRatio);
            break;
        case "dialog/credit-summary/calculate-valid-credits/version-20260116A":
        case "dialog/credit-summary/calculate-valid-credits/version-20260122A":
        case "dialog/credit-summary/calculate-valid-credits/version-20260125A":
        case "dialog/credit-summary/calculate-valid-credits/version-20260622A":
            window.Application.ShowDialog(`${getUrlPath()}/${eleId}/index.html`, "学分汇总（计算有效学分）", 320 * window.devicePixelRatio, 640 * window.devicePixelRatio);
            break;
        case "dialog/detail-and-summary/check-graduation-qualification/version-20260126A":
        case "dialog/detail-and-summary/check-graduation-qualification/version-20260619A":
        case "dialog/detail-and-summary/check-graduation-qualification/version-20260620A":
            window.Application.ShowDialog(`${getUrlPath()}/${eleId}/index.html`, "汇总+明细（检查毕业资格）", 320 * window.devicePixelRatio, 640 * window.devicePixelRatio);
            break;
        default:
            return true;
    }
}

function getImage(control) {
    const eleId = control.Id;
    switch (eleId) {

        default:
            return "images/error-picture.svg";
    }
}