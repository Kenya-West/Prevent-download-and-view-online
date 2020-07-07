class State {
    fileUrl: string;
    notFound: boolean;
    tabId: number;
    nativeOpener: {
        pdf: boolean;
    };

    constructor() {
        this.fileUrl = "";
        this.notFound = false;
        this.tabId = 0;
        this.nativeOpener = {
            pdf: false
        };
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
    // documents
    pdf = "pdf",
    // images
    jpg = "jpg",
    jpeg = "jpeg",
    png = "png",
    gif = "gif",
    tiff = "tiff",
    webp = "webp",
    // sound
    mp3 = "mp3",
    wav = "wav",
    ogg = "ogg",
    // video
    mp4 = "mp4",
    webm = "webm",
    // other
    woff = "woff",
    css = "css",
    htm = "htm",
    html = "html",
    js = "js"
}
enum broswerNativeMIME {
    // documents
    pdf = "application/pdf",
    // images
    jpg = "image/jpg",
    jpeg = "image/jpeg",
    png = "image/png",
    gif = "image/gif",
    tiff = "image/tiff",
    webp = "image/webp",
    // sound
    mp3 = "audio/mp3",
    wav = "audio/wav",
    ogg = "audio/ogg",
    // video
    mp4 = "video/mp4",
    webm = "video/webm",
    // other
    woff = "application/font-woff",
    css = "text/css",
    htm = "text/htm",
    html = "text/html",
    js = "text/javascript"
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

});

chrome.webRequest.onHeadersReceived.addListener((details) => {
    if (details.url.split(".").pop() in broswerNativeFormats) {
        console.debug(`Processing the request at url ${details.url}`);
        const fileExtension = details.url.split(".").pop();
        const response = {
            responseHeaders: null
        };
        const headers = details.responseHeaders;
        headers.map((httpHeader) => {
            if (httpHeader.name === "Content-Disposition" && httpHeader.value.includes("attachment")) {
                console.debug(`Found Content-Disposition. Header is modified`);
                httpHeader.value = "inline";
            }
            if (httpHeader.name === "Content-Type" && httpHeader.value.includes("application/x-forcedownload")) {
                console.debug(`Found Content-Type. Header is modified`);
                httpHeader.value = broswerNativeMIME[fileExtension];
            }
        });
        response.responseHeaders = headers;
        console.debug(`Headers are modified compeletely. The response headers are:`);
        console.debug(response.responseHeaders);
        return response;
    }
}, { urls: ["<all_urls>"], types: ["main_frame"] }, ["blocking", "responseHeaders"]);

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
