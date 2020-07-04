class State {
    fileUrl = "";
    notFound = false;
    tabId = "";

    constructor() {
        this.fileUrl = "";
        this.notFound = false;
        this.tabId = "";
    }
}

class officeExtensions {
    static #word = ["doc", "docx", "docm"];
    static #excel = ["xls", "xlsx", "xlsm", "xlsb", "csv"];
    static #powerpoint = ["ppt", "pptx", "pptm"];
    static #visio = ["vsd", "vsdx", "vsdm", "vssx", "vssm", "vstx", "vstm"];

    static getAll() {
        return this.#word.concat(this.#excel, this.#powerpoint, this.#visio);
    }
}

class OfficeOnline {
    static #strOfficeHost = "view.officeapps.live.com";
    static #strViewOfficeUrl = "/op/view.aspx?src=";
    static #strNotFound = "/op/filenotfound.htm"

    static getUrl() {
        return "https://" + this.#strOfficeHost + this.#strViewOfficeUrl;
    }

    static getHost() {
        return "https://" + this.#strOfficeHost;
    }

    static getOpenedUrlPart() {
        return this.#strViewOfficeUrl;
    }

    static getNotFoundUrl() {
        return "https://" + this.#strOfficeHost + this.#strNotFound;
    }

    static saveDocumentUrl(url) {
        return null
    }

    static createUrl(url) {
        return new URL(location.href);
    }
}

const state = new State();

chrome.downloads.onCreated.addListener(downloadItem => {
    const url = downloadItem.finalUrl;
    console.debug(`Start downloading a file by ${url}. Info: !url?.includes(OfficeOnline.getHost(): ${!url?.includes(OfficeOnline.getHost())}`);
    if (
        officeExtensions.getAll().includes(url?.split(".").pop()) &&
        !url?.includes(OfficeOnline.getHost()) &&
        !state?.notFound
    ) {
        console.debug(`Cancel download`);
        chrome.downloads.cancel(downloadItem.id, () => {
            chrome.tabs.create({ url: OfficeOnline.getUrl() + url }, (tab) => {
                console.debug(`Create a tab with id: ${tab.id}`);
                state.tabId = tab.id;
                state.fileUrl = url;
                state.notFound = false;
            });
        })
    }
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (tabId === state.tabId) {
        console.debug(`Found an updated tab with id: ${tabId}`);
        if (changeInfo?.url?.includes(OfficeOnline.getUrl())) {
            try {
                const url = new URL(changeInfo?.url);
                if (url?.searchParams?.get("src")) {
                    state.fileUrl = url.searchParams.get("src");
                    state.notFound = false;
                    console.debug(`Found a link for file displaying: ${state.fileUrl}`);
                } else {
                    console.debug(`changeInfo?.url is: ${changeInfo?.url}`);
                }
            } catch (error) {
                console.error(`changeInfo.url is not valid! ${changeInfo?.url}`)
            }
        } else if (changeInfo?.url?.includes(OfficeOnline.getNotFoundUrl())) {
            state.notFound = true;
            if (state.fileUrl) {
                console.debug(`The file on URL ${state.fileUrl} is not available, attempting to download`);
                chrome.downloads.download({
                    url: state.fileUrl
                }, (downloadId) => {
                    state.notFound = false;
                })
            } else {
                console.error(`No link to download!`);
            }
        }
    }

});