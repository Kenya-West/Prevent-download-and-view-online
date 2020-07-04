class officeExtensions {
    static #word = ["doc", "docx", "docm"];
    static #excel = ["xls", "xlsx", "xlsm", "xlsb", "csv"];
    static #powerpoint = ["ppt", "pptx", "pptm"];
    static #visio = ["vsd", "vsdx", "vsdm", "vssx", "vssm", "vstx", "vstm"];

    static getAll() {
        const officeExtensionsArray = this.#word.concat(this.#excel, this.#powerpoint, this.#visio);
        return officeExtensionsArray;
    }
}

class OfficeOnline {
    static #strOfficeHost = "view.officeapps.live.com";
    static #strViewOfficeUrl = "/op/view.aspx?src=";

    static buildUrl() {
        return "https://" + this.#strOfficeHost + "/op/view.aspx?src=";
    }
}

chrome.downloads.onCreated.addListener(downloadItem => {
    const url = downloadItem.finalUrl;
    if (officeExtensions.getAll().includes(url.split(".").pop())) {
        chrome.downloads.cancel(downloadItem.id, () => {
            chrome.tabs.create({ url: OfficeOnline.buildUrl() + url });
        })
    }
 });