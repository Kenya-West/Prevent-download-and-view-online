class State {
    downloader: {
        allowDownload: boolean;
        decidedUrl: URL;
    };
    officeOnline: {
        fileUrl: string;
    };
    nativeOpener: {
        pdf: boolean;
    };

    constructor() {
        this.officeOnline = {
            fileUrl: ""
        };
        this.downloader = {
            allowDownload: false,
            decidedUrl: null
        };
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
}

class ChromeTools {
    public cancelDownloadAndOpenTab(downloadItem: chrome.downloads.DownloadItem, url: URL): void {
        if (!url) {
            state.downloader.decidedUrl =
                UrlTools.addUrl(UrlTools.createUrl(downloadItem.url)) ||
                UrlTools.addUrl(UrlTools.createUrl(downloadItem.finalUrl));
        }

        chrome.downloads.cancel(downloadItem.id, () => {
            chrome.tabs.create({ url: OfficeOnline.getUrl() + state.downloader.decidedUrl.href }, (tab) => {
                console.info("%c%s", "color: #D73B02", `Create a tab for Office file with id: ${tab.id}`);
                state.officeOnline.fileUrl = url.href;
                state.downloader.allowDownload = false;
                state.downloader.decidedUrl = null;
            });
        });
    }

    public cancelDownload(downloadItemId: number): void {
        chrome.downloads.cancel(downloadItemId, () => { });
    }

    public recognizeFileExtension(details: chrome.webRequest.WebResponseHeadersDetails): TFileExtensionResponse {
        // There are 3 possible ways to find file name and extension
        console.info("%c%s", "color: #2279CB", `Recognizing file extension at url:\n"${details.url}"`);

        const resultByDetailsUrl = byDetailsUrl(details.url);
        console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Result by details.url is: ${resultByDetailsUrl}`);
        const resultByContentDisposition = byContentDisposition(details.responseHeaders.find(header => header.name.toLowerCase() === "content-disposition")?.value);
        console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Result by details.responseHeaders.content-disposition is: ${resultByContentDisposition}`);
        const resultByContentType = byContentType(details.responseHeaders.find(header => header.name.toLowerCase() === "content-type")?.value);
        console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Result by details.responseHeaders.content-type is: ${resultByContentType}`);

        const extension = resultByDetailsUrl || resultByContentDisposition || resultByContentType;

        return extension;

        function byDetailsUrl(url: string): TFileExtensionResponse | null {
            return UrlTools.getFileExtensionbyUrl(url);
        }
        function byContentDisposition(contentDisposition: string): TFileExtensionResponse | null {
            // https://stackoverflow.com/a/52738125/4846392
            const regex = /filename[^;=\n]*=(?:(\\?['"])(.*?)\1|(?:[^\s]+'.*?')?([^;\n]*))/i;
            const result = regex.exec(contentDisposition);
            if (result) { return UrlTools.getFileExtensionbyString(result[3]) || UrlTools.getFileExtensionbyString(result[2]); }
            return null;
        }

        function byContentType(contentType: string): TFileExtensionResponse | null {
            const keys = Object.keys(broswerNativeFileMIME).filter((key) => broswerNativeFileMIME[key] === contentType);
            return keys.length > 0 ? keys[0] as TFileExtensionResponse : null;
        }

    }

    public filterHeaders(details: chrome.webRequest.WebResponseHeadersDetails): boolean {
        const headers = details.responseHeaders;

        const contentType = headers.find((httpHeader) => httpHeader.name.toLowerCase() === "Content-Type".toLowerCase());

        if (
            contentType.value.includes("text/html") ||
            contentType.value.includes("text/css") ||
            contentType.value.includes("application/json") ||
            contentType.value.includes("application/javascript") ||
            contentType.value.includes("text/javascript") ||
            contentType.value.includes("text/plain") ||
            contentType.value.includes("text/markdown")
        ) {
            return false;
        }

        return true;
    }

    public modifyHeaders(details: chrome.webRequest.WebResponseHeadersDetails,
        fileExtension: broswerNativeFileExtensions | officeFileExtensions): IOnHeadersReceivedResult {
        console.info("%c%s", "color: #2279CB", `Processing the request at url:\n"${details.url}"`);
        const headers = details.responseHeaders;
        headers.map((httpHeader) => {
            if (httpHeader.name.toLowerCase() === "Content-Disposition".toLowerCase()
                && httpHeader.value.toLowerCase().includes("attachment")) {
                httpHeader.value = httpHeader.value.toLowerCase().replace("attachment", "inline");
                console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Found ${httpHeader.name}. Header is modified with value ${httpHeader.value}`);
            }
            if (httpHeader.name.toLowerCase() === "Content-Type".toLowerCase()) {
                httpHeader.value = httpHeader.value.toLowerCase().replace("application/x-forcedownload", `${broswerNativeFileMIME[fileExtension]}`);
                httpHeader.value = httpHeader.value.toLowerCase().replace("application/forcedownload", `${broswerNativeFileMIME[fileExtension]}`);
                httpHeader.value = httpHeader.value.toLowerCase().replace("application/octet-stream", `${broswerNativeFileMIME[fileExtension]}`);
                console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Found ${httpHeader.name}. Header is modified with value ${httpHeader.value}`);
            }
        });

        const result: IOnHeadersReceivedResult = {
            responseHeaders: headers
        };

        return result;
    }

    constructor() {

        chrome.downloads.onCreated.addListener(downloadItem => {
            if (!state.downloader.allowDownload) {
                this.cancelDownloadAndOpenTab(downloadItem, state.downloader.decidedUrl);
            }
        });

    }

}

class UrlTools {
    public static sanitizeUrl(url: URL, skipLog?: boolean): URL {
        try {
            skipLog === false ?? console.info("%c%s", "padding-left: 2rem;", `Sanitized URL out from:\n"${url.href}"\nand got:\n"${new URL(url.protocol + "//" + url.host + url.pathname).href}"`);
            return new URL(url.protocol + "//" + url.host + url.pathname);
        } catch (error) {
            throw new Error("Couldn't convert URL");
        }
    }

    public static createUrl(url: string, skipLog?: boolean): URL {
        try {
            skipLog === false ?? console.info("%c%s", "padding-left: 2rem;", `Created URL from string:\n"${url}"\nand got URL object with href:\n"${new URL(url).href}"`);
            return new URL(url);
        } catch (error) {
            throw new Error("Couldn't convert URL");
        }
    }

    public static addUrl(url: URL): URL {
        return this.getFileExtensionbyUrl(url.href) ? url : null;
    }

    public static getFileExtensionbyUrl(url: string): TFileExtensionResponse | null {
        const filePath = UrlTools.sanitizeUrl(UrlTools.createUrl(url)).href.split(".").pop();
        const fileExtension = broswerNativeFileExtensions[filePath] || officeFileExtensions[filePath];
        return fileExtension;
    }

    public static getFileExtensionbyString(fileName: string): TFileExtensionResponse | null {
        if (fileName) {
            const fileExtension: TFileExtensionResponse = Object.values(broswerNativeFileExtensions).find((extension) => {
                return fileName.includes(extension);
            });
            return fileExtension;
        }
        return null;
    }
}

interface IHTttpHeader {
    key: string;
    value: string;
    keyReplace?: string;
    valueReplace?: string;
    relpaceWith(): void;
}

interface IOnHeadersReceivedResult {
    responseHeaders: chrome.webRequest.HttpHeader[];
}

type TFileExtensionResponse = broswerNativeFileExtensions | officeFileExtensions;

class HttpHeader implements IHTttpHeader {
    key: string;
    value: string;
    keyReplace?: string;
    valueReplace?: string;

    relpaceWith(): void {
        if (this.keyReplace) {
            this.key = this.keyReplace;
        }
        if (this.valueReplace) {
            this.value = this.valueReplace;
        }
    }

    constructor(init?: Partial<HttpHeader>) {
        Object.assign(this, init);
    }
}

enum broswerNativeFileExtensions {
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
enum broswerNativeFileMIME {
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
enum officeFileExtensions {
    doc = "doc",
    docx = "docx",
    xls = "xls",
    xlsx = "xlsx",
    ppt = "ppt",
    pptx = "pptx"
}
enum officeFileMIME {
    doc = "application/msword",
    docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xls = "application/vnd.ms-excel",
    xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ppt = "application/vnd.ms-powerpoint",
    pptx = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
}

const state = new State();
const chromeTools = new ChromeTools();
const httpHeadersToFind: HttpHeader[] = [
    new HttpHeader(
        {
            key: "Content-Type",
            value: "application/x-forcedownload",
            valueReplace: "application/pdf"
        }),
    new HttpHeader(
        {
            key: "Content-Disposition",
            value: "attachment",
            valueReplace: "inline"
        })
];

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (tab.url.includes(OfficeOnline.getHost())) {
        console.info("%c%s", "color: #D73B02", `Found an updated tab with id: ${tabId} and URL:\n"${tab.url}"`);

        if (changeInfo?.url?.includes(OfficeOnline.getUrl())) {
            try {
                const url = UrlTools.createUrl(changeInfo?.url);
                if (url?.searchParams?.get("src")) {
                    state.officeOnline.fileUrl = url.searchParams.get("src");
                    state.downloader.allowDownload = false;
                    console.info("%c%s", "color: #D73B02", `Found a link for file displaying:\n"${state.officeOnline.fileUrl}"`);
                } else {
                    console.info("%c%s", "color: #D73B02", `changeInfo?.url is:\n"${changeInfo?.url}"`);
                }
            } catch (error) {
                console.warn(`changeInfo.url is not valid at path:\n"${changeInfo?.url}"`);
            }
        } else if (changeInfo?.url?.includes(OfficeOnline.getNotFoundUrl())) {
            state.downloader.allowDownload = true;
            if (state.officeOnline.fileUrl) {
                console.info("%c%s", "color: #D73B02", `The file on URL:\n"${state.officeOnline.fileUrl}"\nis not available, attempting to download`);
                chrome.downloads.download({
                    url: state.officeOnline.fileUrl
                }, (downloadId) => {
                    chrome.tabs.remove(tabId);
                    state.downloader.allowDownload = false;
                });
            } else {
                console.warn(`No link to download!`);
            }
        }

    }

});

chrome.webRequest.onHeadersReceived.addListener((details) => {
    if (!details.url.includes(OfficeOnline.getHost())) {
        if (chromeTools.filterHeaders(details)) {

            state.downloader.decidedUrl = null;

            const fileExtension = chromeTools.recognizeFileExtension(details);
            console.info("%c%s", "color: #2279CB", `Searched for file extension. Got: "${fileExtension}"`);
            if (fileExtension in broswerNativeFileExtensions) {

                state.downloader.decidedUrl = UrlTools.addUrl(UrlTools.createUrl(details.url));

                return chromeTools.modifyHeaders(details, fileExtension as broswerNativeFileExtensions);
            }
            if (fileExtension in officeFileExtensions) {

                state.downloader.decidedUrl = UrlTools.addUrl(UrlTools.createUrl(details.url));
                state.downloader.allowDownload = true;

                return chromeTools.modifyHeaders(details, fileExtension as officeFileExtensions);
            }

        }
    }
    return null;

}, { urls: ["<all_urls>"], types: ["main_frame"] }, ["blocking", "responseHeaders"]);
