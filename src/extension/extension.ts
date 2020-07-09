class State {
    fileUrl: string;
    doNotDownload: boolean;
    tabId: number;
    nativeOpener: {
        pdf: boolean;
    };

    constructor() {
        this.fileUrl = "";
        this.doNotDownload = false;
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
}

class ChromeTools {
    public cancelDownloadAndOpenTab(downloadItemId: number, url: URL, media?: boolean): void {
        if (!media) {
            chrome.downloads.cancel(downloadItemId, () => {
                chrome.tabs.create({ url: OfficeOnline.getUrl() + url.href }, (tab) => {
                    console.info("%c%s", "color: #D73B02", `Create a tab for Office file with id: ${tab.id}`);
                    state.tabId = tab.id;
                    state.fileUrl = url.href;
                    state.doNotDownload = false;
                });
            });
        } else {
            chrome.downloads.cancel(downloadItemId, () => {
                chrome.tabs.create({ url: url.href }, (tab) => {
                    console.info(`Create a tab for native browser media file with id: ${tab.id}`);
                    state.doNotDownload = false;
                });
            });
        }
    }

    public cancelDownload(downloadItemId: number): void {
        chrome.downloads.cancel(downloadItemId, () => { });
    }

    public decideUrl(url1: URL, url2: URL): URL {
        const urlStart = url1;
        const urlFinal = url2;
        console.info(`Detected download of a file. Decide which url is true:\n    url: "${urlStart.href}"\n    or finalUrl: "${urlFinal.href}"`);

        let decidedUrl: URL;
        if (urlStart && OfficeExtensions.getAll().includes(urlStart.href?.split(".").pop())) { decidedUrl = urlStart; }
        if (urlFinal && OfficeExtensions.getAll().includes(urlFinal.href?.split(".").pop())) { decidedUrl = urlFinal; }
        if (urlStart && (Object.values(broswerNativeFormats) as Array<string>).includes(urlStart.href?.split(".").pop())) { decidedUrl = urlStart; }
        if (urlFinal && (Object.values(broswerNativeFormats) as Array<string>).includes(urlFinal.href?.split(".").pop())) { decidedUrl = urlFinal; }

        return decidedUrl;
    }

    public findHeaders(details?: chrome.webRequest.WebResponseHeadersDetails): HttpHeader[] {
        if (details.url.split(".").pop() in broswerNativeFormats) {
            console.info("%c%s", "color: #2279CB", `Finding headers at url ${details.url}`);
            const fileExtension = details.url.split(".").pop();
            const response = {
                responseHeaders: null
            };
            const headers = details.responseHeaders;

            const headersFound: HttpHeader[] = [];
            headers.map((httpHeader) => {
                httpHeadersToFind.map((httpHeaderToFind) => {
                    if (httpHeaderToFind.key === httpHeader.name
                        && httpHeaderToFind.value === httpHeader.value) {
                        headersFound.push();
                    }
                });
            });
            return headersFound;
        }
    }

    public recognizeFileExtension(details: chrome.webRequest.WebResponseHeadersDetails): broswerNativeFormats {
        // There are 3 possible ways to find file name and extension
        console.info("%c%s", "color: #2279CB", `Recognizing file extension at url:\n"${details.url}"`);

        const resultByDetailsUrl = byDetailsUrl(details.url);
        console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Result by details.url is: ${resultByDetailsUrl}`);
        const resultByContentDisposition = byContentDisposition(details.responseHeaders.find(header => header.name === "content-disposition")?.value);
        console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Result by details.responseHeaders.content-disposition is: ${resultByContentDisposition}`);
        const resultByContentType = byContentType(details.responseHeaders.find(header => header.name === "content-type")?.value);
        console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Result by details.responseHeaders.content-type is: ${resultByContentType}`);

        const extension = resultByDetailsUrl || resultByContentDisposition || resultByContentType;

        return broswerNativeFormats[extension];

        function byDetailsUrl(url: string): string | null {
            return byFileName(url);
        }
        function byContentDisposition(contentDisposition: string): string | null {
            // https://stackoverflow.com/a/52738125/4846392
            const regex = /filename[^;=\n]*=(?:(\\?['"])(.*?)\1|(?:[^\s]+'.*?')?([^;\n]*))/i;
            const result = regex.exec(contentDisposition);
            if (result) { return byFileName(result[3]) || byFileName(result[2]); }
            return null;
        }

        function byContentType(contentType: string): string | null {
            const keys = Object.keys(broswerNativeMIME).filter((key) => broswerNativeMIME[key] === contentType);
            return keys.length > 0 ? keys[0] : null;
        }

        function byFileName(url: string): string | null {
            const fileExtension = url?.split(".").pop();
            if (fileExtension in broswerNativeFormats) {
                return fileExtension;
            }
            return null;
        }

    }
    public modifyHeaders(details: chrome.webRequest.WebResponseHeadersDetails,
        fileExtension: broswerNativeFormats): IOnHeadersReceivedResult {
        console.info("%c%s", "color: #2279CB", `Processing the request at url:\n"${details.url}"`);
        const headers = details.responseHeaders;
        headers.map((httpHeader) => {
            if (httpHeader.name.toLowerCase() === "Content-Disposition".toLowerCase()
                && httpHeader.value.toLowerCase().includes("attachment")) {
                httpHeader.value = httpHeader.value.toLowerCase().replace("attachment", "inline");
                console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Found ${httpHeader.name}. Header is modified with value ${httpHeader.value}`);
            }
            if (httpHeader.name.toLowerCase() === "Content-Type".toLowerCase()) {
                httpHeader.value = httpHeader.value.toLowerCase().replace("application/x-forcedownload", `${broswerNativeMIME[fileExtension]}`);
                httpHeader.value = httpHeader.value.toLowerCase().replace("application/octet-stream", `${broswerNativeMIME[fileExtension]}`);
                console.info("%c%s", "padding-left: 2rem; color: #2279CB", `Found ${httpHeader.name}. Header is modified with value ${httpHeader.value}`);
            }
        });

        const result: IOnHeadersReceivedResult = {
            responseHeaders: headers
        };

        return result;
    }

    public fixHeaders(details?: chrome.webRequest.WebResponseHeadersDetails): any {
        if (details.url.split(".").pop() in broswerNativeFormats) {
            console.info("%c%s", "color: #2279CB", `Processing the request at url:\n"${details.url}"`);
            const fileExtension = details.url.split(".").pop();
            const response = {
                responseHeaders: null
            };
            const headers = details.responseHeaders;

            let modified = false;
            headers.map((httpHeader) => { // iterate headers
                if (httpHeader.name === "Content-Disposition" && httpHeader.value.includes("attachment")) {
                    console.info("%c%s", "color: #2279CB", `Found Content-Disposition. Header is modified`);
                    httpHeader.value = "inline";
                    modified = true;
                }
                if (httpHeader.name === "Content-Type" && httpHeader.value.includes("application/x-forcedownload")) {
                    console.info("%c%s", "color: #2279CB", `Found Content-Type. Header is modified`);
                    httpHeader.value = broswerNativeMIME[fileExtension];
                    modified = true;
                }
            });
            if (modified) {
                response.responseHeaders = headers;
                console.info("%c%s", "color: #2279CB", `Headers are modified compeletely. The response headers are:`);
                console.info("%c%s", "color: #2279CB", "response.responseHeaders:", response.responseHeaders);
            }
            return response;
        }
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

// chrome.downloads.onCreated.addListener(downloadItem => {
//     state?.doNotDownload ?? console.info("%c%s", "color: #D73B02", `Since file is not found, then download it`);

//     if (!state?.doNotDownload) {

//         const decidedUrl = chromeTools.decideUrl(
//             UrlTools.sanitizeUrl(UrlTools.createUrl(downloadItem.url)),
//             UrlTools.sanitizeUrl(UrlTools.createUrl(downloadItem.finalUrl))
//         );

//         console.info(`Decided URL is:\n    "${decidedUrl?.href}"`);
//         // if they are Office files
//         if (!decidedUrl?.href?.includes(OfficeOnline.getHost())) {
//             if (OfficeExtensions.getAll().includes(decidedUrl.href.split(".").pop())) {
//                 console.info("%c%s", "color: #D73B02", `Recognized an Office file. Cancel download`);
//                 chromeTools.cancelDownloadAndOpenTab(downloadItem.id, decidedUrl);
//             }
//             if ((Object.values(broswerNativeFormats) as Array<string>).includes(decidedUrl.href.split(".").pop())) {
//                 console.info("%c%s", "color: #D73B02", `Recognized a media file. Cancel download`);
//                 chromeTools.cancelDownload(downloadItem.id);
//             }
//         }

//     }
// });

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (tab.url.includes(OfficeOnline.getHost())) {
        console.info("%c%s", "color: #D73B02", `Found an updated tab with id: ${tabId} and URL:\n"${tab.url}"`);

        if (changeInfo?.url?.includes(OfficeOnline.getUrl())) {
            try {
                const url = UrlTools.createUrl(changeInfo?.url);
                if (url?.searchParams?.get("src")) {
                    state.fileUrl = url.searchParams.get("src");
                    state.doNotDownload = false;
                    console.info("%c%s", "color: #D73B02", `Found a link for file displaying:\n"${state.fileUrl}"`);
                } else {
                    console.info("%c%s", "color: #D73B02", `changeInfo?.url is:\n"${changeInfo?.url}"`);
                }
            } catch (error) {
                console.warn(`changeInfo.url is not valid at path:\n"${changeInfo?.url}"`);
            }
        } else if (changeInfo?.url?.includes(OfficeOnline.getNotFoundUrl())) {
            state.doNotDownload = true;
            if (state.fileUrl) {
                console.info("%c%s", "color: #D73B02", `The file on URL:\n"${state.fileUrl}"\nis not available, attempting to download`);
                chrome.downloads.download({
                    url: state.fileUrl
                }, (downloadId) => {
                    state.doNotDownload = false;
                });
            } else {
                console.warn(`No link to download!`);
            }
        }

    }

});

chrome.webRequest.onHeadersReceived.addListener((details) => {
    const fileExtension = chromeTools.recognizeFileExtension(details);
    console.info("%c%s", "color: #2279CB", `Searched for file extension. Got: "${fileExtension}"`);
    if (fileExtension) {
        return chromeTools.modifyHeaders(details, fileExtension);
    }
    return null;

}, { urls: ["<all_urls>"], types: ["main_frame"] }, ["blocking", "responseHeaders"]);
