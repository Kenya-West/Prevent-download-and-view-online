class State {
    fileUrl: string;
    notFound: boolean;
    tabId: number;

    constructor() {
        this.fileUrl = "";
        this.notFound = false;
        this.tabId = 0;
    }
}

class OfficeExtensions {
    private static word = ["doc", "docx", "docm"];
    private static excel = ["xls", "xlsx", "xlsm", "xlsb", "csv"];
    private static powerpoint = ["ppt", "pptx", "pptm"];
    private static visio = ["vsd", "vsdx", "vsdm", "vssx", "vssm", "vstx", "vstm"];

    public static getAll() {
        return this.word.concat(this.excel, this.powerpoint, this.visio);
    }
}

class OfficeOnline {
    private static strOfficeHost = "view.officeapps.live.com";
    private static strViewOfficeUrl = "/op/view.aspx?src=";
    private static strNotFound = "/op/filenotfound.htm";

    public static getUrl(): string {
        return "https://" + this.strOfficeHost + this.strViewOfficeUrl;
    }

    public static getHost(): string {
        return "https://" + this.strOfficeHost;
    }

    public static getOpenedUrlPart(): string {
        return this.strViewOfficeUrl;
    }

    static getNotFoundUrl(): string {
        return "https://" + this.strOfficeHost + this.strNotFound;
    }

    static createUrl(url): URL {
        return new URL(location.href);
    }
}

enum broswerNativeFormats {
    pdf = "pdf"
}

const state = new State();

chrome.downloads.onCreated.addListener(downloadItem => {
    const url = downloadItem.finalUrl;
    console.debug(`Start downloading a file by ${url}`);

    // if they are Office files
    if (
        OfficeExtensions.getAll().includes(url?.split(".").pop()) &&
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
        });
    }

    // if this is a pdf file or a file that can be opened in a tab
    if (url?.split(".").pop() === broswerNativeFormats.pdf) {
        console.debug(`Cancel download`);
        state.notFound = true; // disable pdf opening
        chrome.downloads.cancel(downloadItem.id, () => {
            chrome.tabs.create({ url }, (tab) => {
                state.notFound = false;
            });
        });
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
                console.error(`changeInfo.url is not valid! ${changeInfo?.url}`);
            }
        } else if (changeInfo?.url?.includes(OfficeOnline.getNotFoundUrl())) {
            state.notFound = true;
            if (state.fileUrl) {
                console.debug(`The file on URL ${state.fileUrl} is not available, attempting to download`);
                chrome.downloads.download({
                    url: state.fileUrl
                }, (downloadId) => {
                    state.notFound = false;
                });
            } else {
                console.error(`No link to download!`);
            }
        }

    }

});
