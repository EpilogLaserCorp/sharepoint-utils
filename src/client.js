import axios from 'axios';
import { promises as fs, createWriteStream } from 'fs';
import path from 'path';
import { minimatch } from 'minimatch';

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
            let max_retries = 5;
            let lastError = null;
            while (max_retries > 0)
            {
                try {
                    return await this.uploadSmallFile(uploadFilePath, fileContent);
                } catch (error) {
                    lastError = error;
                    console.error("Error uploading small file:", error.response?.data);
                    max_retries--;
                    const retryAfterSeconds = error.response?.data?.error?.retryAfterSeconds || 5;
                    console.log(`Retry ${max_retries} after ${retryAfterSeconds} seconds...`);
                    await new Promise(resolve => setTimeout(resolve, retryAfterSeconds * 1000));
                }
            }
            return lastError;
        } else {
            const response = await this.uploadSession(uploadFilePath);
            const responseJson = response.data;

            if (responseJson.uploadUrl) {
                const totalSize = fileSize;
                const chunkSize = 327680 * 100; // ~30MB
                let offset = 0;
                let max_retries = 5;
                while (offset < totalSize) {
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
                        // Reset the number of retries if successful, a given chunk should have a chance to be uploaded
                        max_retries = 5;
                        console.log(`Uploaded ${convertSize(offset)}/${convertSize(totalSize)}`);
                    } catch (error) {
                        console.error("Error doing session upload:", error.response?.data);
                        if (max_retries > 0) {
                            max_retries--;
                            const retryAfterSeconds = error.response?.data?.error?.retryAfterSeconds || 5;
                            console.log(`Retry ${max_retries} after ${retryAfterSeconds} seconds...`);
                            await new Promise(resolve => setTimeout(resolve, retryAfterSeconds * 1000));
                            continue;
                        }
                        throw error;
                    }
                }
            }
            return response;
        }
    }

    async downloadFile(sharepointFilePath, localFilePath) {
        const fileUrl = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root:${sharepointFilePath}:/content`;
        let max_retries = 5;
        let lastError = null;

        while (max_retries > 0) {
            try {
                const response = await axios.get(fileUrl, {
                    headers: { "Authorization": `Bearer ${this.accessToken}` },
                    responseType: 'stream'
                });

                const totalSize = parseInt(response.headers['content-length'], 10);
                let downloadedSize = 0;
                
                const writer = createWriteStream(localFilePath);
                
                // Track progress if file size is available
                if (totalSize && totalSize > 1024 * 1024) { // Only show progress for files > 1MB
                    response.data.on('data', (chunk) => {
                        downloadedSize += chunk.length;
                        const percentage = ((downloadedSize / totalSize) * 100).toFixed(1);
                        process.stdout.write(`\rDownloading ${path.basename(sharepointFilePath)}: ${percentage}% (${convertSize(downloadedSize)}/${convertSize(totalSize)})`);
                    });
                }
                
                response.data.pipe(writer);

                return new Promise((resolve, reject) => {
                    writer.on('finish', () => {
                        if (totalSize && totalSize > 1024 * 1024) {
                            process.stdout.write('\n'); // New line after progress
                        }
                        console.log(`Downloaded file to ${localFilePath}`);
                        resolve(response);
                    });
                    writer.on('error', reject);
                });
            } catch (error) {
                lastError = error;
                console.error("Error downloading file:", error.response?.data);
                max_retries--;
                if (max_retries > 0) {
                    const retryAfterSeconds = error.response?.data?.error?.retryAfterSeconds || 5;
                    console.log(`Retry ${max_retries} after ${retryAfterSeconds} seconds...`);
                    await new Promise(resolve => setTimeout(resolve, retryAfterSeconds * 1000));
                } else {
                    throw lastError;
                }
            }
        }
        throw lastError;
    }

    async listFolderRecursively(folderPath = '') {
        const allFiles = [];
        await this._listFolderRecursivelyHelper(folderPath, allFiles);
        return allFiles;
    }

    async _listFolderRecursivelyHelper(folderPath, allFiles) {
        const folderUrl = folderPath 
            ? `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root:${folderPath}:/children`
            : `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root/children`;

        try {
            const response = await axios.get(folderUrl, {
                headers: { "Authorization": `Bearer ${this.accessToken}` }
            });

            const items = response.data.value || [];
            
            for (const item of items) {
                const itemPath = folderPath ? `${folderPath}/${item.name}` : `/${item.name}`;
                
                if (item.folder) {
                    await this._listFolderRecursivelyHelper(itemPath, allFiles);
                } else {
                    allFiles.push({
                        name: item.name,
                        path: itemPath,
                        size: item.size || 0,
                        lastModified: item.lastModifiedDateTime
                    });
                }
            }
        } catch (error) {
            console.error(`Error listing folder ${folderPath}:`, error.response?.data || error.message);
            throw error;
        }
    }

    async listFolder(folderPath = '') {
        const folderUrl = folderPath 
            ? `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root:${folderPath}:/children`
            : `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drive/root/children`;

        try {
            const response = await axios.get(folderUrl, {
                headers: { "Authorization": `Bearer ${this.accessToken}` }
            });

            const items = response.data.value || [];
            return items.map(item => ({
                name: item.name,
                path: folderPath ? `${folderPath}/${item.name}` : `/${item.name}`,
                isFolder: !!item.folder,
                size: item.size || 0,
                lastModified: item.lastModifiedDateTime
            }));
        } catch (error) {
            console.error(`Error listing folder ${folderPath}:`, error.response?.data || error.message);
            throw error;
        }
    }

    matchWildcardPattern(files, pattern) {
        return files.filter(file => {
            const fileName = path.basename(file.path);
            const filePath = file.path;
            
            return minimatch(fileName, pattern) || minimatch(filePath, pattern);
        });
    }

    async downloadMultipleFiles(files, localBasePath) {
        const results = [];
        const totalFiles = files.length;
        
        console.log(`Starting download of ${totalFiles} files...`);
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                // Preserve the relative path structure
                const relativePath = file.path.startsWith('/') ? file.path.substring(1) : file.path;
                const localFilePath = path.join(localBasePath, relativePath);
                
                await fs.mkdir(path.dirname(localFilePath), { recursive: true });
                
                console.log(`[${i + 1}/${totalFiles}] Downloading ${file.path}...`);
                await this.downloadFile(file.path, localFilePath);
                results.push({ success: true, file: file.path, localPath: localFilePath });
                console.log(`Downloaded ${file.path} to ${localFilePath}`);
            } catch (error) {
                results.push({ success: false, file: file.path, error: error.message });
                console.error(`Failed to download ${file.path}:`, error.message);
            }
        }
        
        return results;
    }
}

export { SharePointClient };