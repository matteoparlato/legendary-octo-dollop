import * as MicrosoftGraph from "@microsoft/microsoft-graph-client"
import { DriveItem } from "@microsoft/microsoft-graph-types";
import fetch from 'cross-fetch';
import { ConflictBehaviour } from "./task-inputs";

interface AadAuthToken {
    token_type: "Bearer"
    expires_in: number,
    access_token: string
}

interface UploaderAuthOptions {
    tenantId: string;
    clientId: string;
    clientSecret: string;
}

interface UploaderOptions {
    auth: UploaderAuthOptions;
    conflictBehaviour: ConflictBehaviour;
}

const streamToText = async (blob: any) => {
    const readableStream = await blob.getReader();
    const chunk = await readableStream.read();

    return new TextDecoder('utf-8').decode(chunk.value);
};

class SharePointUploader {
    private static _defaults = {
        aadScope: 'https://graph.microsoft.com/.default'
    };

    private _options: UploaderOptions;
    private _client?: MicrosoftGraph.Client;

    constructor(options: UploaderOptions) {
        this._options = options;
    }

    setConflictBehaviour(conflictBehaviour: ConflictBehaviour): void {
        this._options.conflictBehaviour = conflictBehaviour;
    }

    static getAccessTokenEndpoint(tenantId: string): string {
        return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    }

    static async getAccessTokenAsync(authOptions: UploaderAuthOptions): Promise<AadAuthToken> {
        const endpointUrl = this.getAccessTokenEndpoint(authOptions.tenantId);

        /*
        var data = "client_id=9947a6cd-864d-4edd-a8ca-6dcaffc420fe&client_secret=wtq8Q~ewWwJNXcY5hDitIo~3Uu6LIf5f.nZ4mcUu&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&grant_type=client_credentials";

        var xhr = new XMLHttpRequest();
        xhr.withCredentials = false;
        
        xhr.open("POST", "https://login.microsoftonline.com/b4deee66-87bf-414d-9362-8ad13c5693ae/oauth2/v2.0/token", false);
        xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
        xhr.setRequestHeader("Access-Control-Allow-Origin", "*");
        //xhr.setRequestHeader("Host", "login.microsoftonline.com");
        //xhr.setRequestHeader("Accept-Encoding", "gzip, deflate, br");
        //xhr.setRequestHeader("Connection", "keep-alive");
        //xhr.setRequestHeader("Referrer-Policy", "no-referrer");
        
        xhr.send(data);
        if (xhr.status === 200) {
            console.log(xhr.responseText);
        }
        */

        var urlencoded = new URLSearchParams();
        urlencoded.append("client_id", "9947a6cd-864d-4edd-a8ca-6dcaffc420fe");
        urlencoded.append("client_secret", "wtq8Q~ewWwJNXcY5hDitIo~3Uu6LIf5f.nZ4mcUu");
        urlencoded.append("scope", "https://graph.microsoft.com/.default");
        urlencoded.append("grant_type", "client_credentials");

        var response = await fetch(endpointUrl, {
            body: urlencoded,
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Access-Control-Allow-Origin': '*',
                'Host': 'login.microsoftonline.com',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'Referrer-Policy': 'no-referrer'
            }
        });

        var result = await response.json() as AadAuthToken;
        console.log(result.access_token);

        return result;
    }

    static createClient(aadToken: AadAuthToken): MicrosoftGraph.Client {
        const client = MicrosoftGraph.Client.initWithMiddleware({
            authProvider: {
                getAccessToken: () => Promise.resolve(aadToken.access_token)
            }
        });

        return client;
    }

    async getClientAsync(): Promise<MicrosoftGraph.Client> {
        if (!!this._client) {
            return this._client;
        }

        const aadToken = await SharePointUploader.getAccessTokenAsync(this._options.auth);
        const client = SharePointUploader.createClient(aadToken);

        this._client = client;
        return client;
    }

    static constructFileUrl(driveId: string, fileName: string, path: string = '/'): string {
        let pathComponents = path.split("/")
            .concat(fileName.split("/"));

        let encodedPath = pathComponents
            .map((p) => encodeURIComponent(p))
            .filter((p) => p !== '')
            .join("/");

        return `/drives/${driveId}/root:/${encodedPath}`;
    }

    async uploadFileAsync(localFile: File, driveId: string, remotePath: string, remoteFileName: string): Promise<MicrosoftGraph.UploadResult | null> {
        const client = await this.getClientAsync();

        const requestUrl = SharePointUploader.constructFileUrl(driveId, remoteFileName, remotePath)
            + ':/createUploadSession';

        const fileSize = localFile.size;


        var Readable = require('stream').Readable
        var s = new Readable()
        s.push(await streamToText(localFile.stream))    // the string you want
        s.push(null)      // indicates end-of-file basically - the end of the stream

        if (fileSize === 0) {
            const byte0Msg = `SharePoint Online does not support 0-Byte files: '${localFile}'.`;
            console.log(byte0Msg);

            return null;
        }

        const payload = {
            "item": {
                "@microsoft.graph.conflictBehavior": this._options.conflictBehaviour,
                "name": remoteFileName,
            },
            "deferCommit": fileSize === 0
        };

        console.log(`Local file path: '${localFile}'.`);
        console.log(`Remote folder path: '${remotePath}'.`);
        console.log(`Remote file name: '${remoteFileName}'.`);
        console.log(`Request url: '${requestUrl}'.`);

        //const readStream = localFile.stream;
        const fileObject = new MicrosoftGraph.StreamUpload(s, remoteFileName, fileSize);

        const uploadSession = await MicrosoftGraph.LargeFileUploadTask.createUploadSession(client, requestUrl, payload);

        const uploadTask = new MicrosoftGraph.LargeFileUploadTask(client, fileObject, uploadSession);

        const uploadedFile = await uploadTask.upload();
        return uploadedFile;
    }

    async cleanFolderAsync(driveId: string, remotePath: string): Promise<void> {
        console.log(`Cleaning target folder '${remotePath}'.`);

        const client = await this.getClientAsync();

        const folderUrl = SharePointUploader.constructFileUrl(driveId, '', remotePath);

        let result: any;
        try {
            result = await client.api(`${folderUrl}:/children?$select=name,id,folder`)
                .get();
        } catch (error: unknown) {
            if ((error as any).statusCode !== 404) {
                throw error;
            }
        }

        if(!result || !result.value)
            return;

        const items = result.value as DriveItem[];
        for (let i = 0; i < items.length; i++) {
            const item = items[i];

            const itemUrl = `/drives/${driveId}/items/${item.id}`;
            await client.api(itemUrl).delete();
        }
    }
}

export default SharePointUploader;