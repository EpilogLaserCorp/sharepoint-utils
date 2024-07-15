const axios = require('axios');
const fs = require('fs').promises;
const path = require('path');

function convertSize(sizeBytes) {
    if (sizeBytes === 0) return "0B";
    const sizeName = ["B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"];
    const i = Math.floor(Math.log(sizeBytes) / Math.log(1024));
    const p = Math.pow(1024, i);
    const s = (sizeBytes / p).toFixed(2);
    return `${s} ${sizeName[i]}`;
}

class SharePointClient {
    constructor(tenantId, clientId, clientSecret, resourceUrl, siteUrl) {
        this.tenantId = tenantId;
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.resourceUrl = resourceUrl;
        this.baseUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        this.headers = { "Content-Type": "application/x-www-form-urlencoded" };
        this.accessToken = null;
        this.siteId = null;
        this.siteUrl = siteUrl;
    }

    async initialize() {
        this.accessToken = await this.getAccessToken();
        this.siteId = await this.getSiteId(this.siteUrl);
    }

    async getAccessToken() {
        const body = new URLSearchParams({
            grant_type: "client_credentials",
            client_id: this.clientId,
            client_secret: this.clientSecret,
            scope: `${this.resourceUrl}.default`
        });

        const response = await axios.post(this.baseUrl, body, { headers: this.headers });
        return response.data.access_token;
    }

    async getSiteId(siteUrl) {
        const fullUrl = `https://graph.microsoft.com/v1.0/sites/${siteUrl}`;
        const response = await axios.get(fullUrl, {
            headers: { "Authorization": `Bearer ${this.accessToken}` }
        });
        return response.data.id;
    }

    async getDriveId() {
        const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives`;
        const response = await axios.get(drivesUrl, {
            headers: { "Authorization": `Bearer ${this.accessToken}` }
        });
        const drives = response.data.value || [];
        return drives.map(drive => [drive.id, drive.name]);
    }

    async getFolderContent(driveId) {
        const folderUrl = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${driveId}/root/children`;
        const response = await axios.get(folderUrl, {
            headers: { "Authorization": `Bearer ${this.accessToken}` }
        });
        const itemsData = response.data;
        const rootdir = [];
        if (itemsData.value) {
            for (const item of itemsData.value) {
                rootdir.push([item.id, item.name]);
            }
        }
        return rootdir;
    }

    async uploadSmallFile(uploadFilePath, data) {
        const folderUrl = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root:${uploadFilePath}:/content`;
        const response = await axios.put(folderUrl, data, {
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
                "Content-Type": "text/plain"
            }
        });
        return response;
    }

    async uploadSession(filePath) {
        const folderUrl = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root:${filePath}:/createUploadSession`;
        const response = await axios.post(folderUrl, {
            "@microsoft.graph.conflictBehavior": "replace",
            "name": "my_file.txt"
        }, {
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
                "Content-Type": "application/json"
            }
        });
        return response;
    }

    async uploadFile(uploadFilePath, filePath) {
        const fileSize = (await fs.stat(filePath)).size;
        const MAX_SMALL_FILE_SIZE = 1 << 27; // ~137MB

        const fileContent = await fs.readFile(filePath);

        if (fileSize < MAX_SMALL_FILE_SIZE) {
            return await this.uploadSmallFile(uploadFilePath, fileContent);
        } else {
            const response = await this.uploadSession(uploadFilePath);
            const responseJson = response.data;

            if (responseJson.uploadUrl) {
                const totalSize = fileSize;
                const chunkSize = 327680 * 100; // ~30MB
                let offset = 0;
                let ok = true;

                while (ok && offset < totalSize) {
                    const thisChunkSize = Math.min(chunkSize, totalSize - offset);
                    const data = fileContent.slice(offset, offset + thisChunkSize);

                    const headers = {
                        "Authorization": `Bearer ${this.accessToken}`,
                        "Content-Length": `${thisChunkSize}`,
                        "Content-Range": `bytes ${offset}-${offset + thisChunkSize - 1}/${totalSize}`
                    };

                    try {
                        await axios.put(responseJson.uploadUrl, data, { headers });
                        offset += thisChunkSize;
                        console.log(`Uploaded ${convertSize(offset)}/${convertSize(totalSize)}`);
                    } catch (error) {
                        ok = false;
                        console.error("Error doing session upload:", error.response?.data);
                    }
                }
            }
            return response;
        }
    }
}

module.exports = { SharePointClient };