import path from 'path';
import * as client_lib from './client.js';

async function uploadFiles(client, uploadPath, filePath) {
    const fileLines = filePath.split('\n');
    for (const line of fileLines) {
        const baseName = path.basename(line);
        let fullUploadPath = uploadPath === "/" ? "/" : uploadPath.replace(/\/$/, "");
        fullUploadPath = `${fullUploadPath}/${baseName}`;
        console.log(`Uploading "${line}" to "${fullUploadPath}"...`);
        try {
            const response = await client.uploadFile(fullUploadPath, line);
            if (response.status === 200 || response.status === 201) {
                console.log(`Uploaded ${line}`);
            } else {
                console.error("Error uploading file. Response:", response.data);
                process.exit(1);
            }
        } catch (error) {
            console.error("Error uploading file:", error.message);
            process.exit(1);
        }
    }
    process.exit(0);
}

async function main() {
    const [action, siteName, sharepointHostName, tenantId, clientId, clientSecret, uploadPath, filePath, maxRetry = "3"] = process.argv.slice(2);
    const resource = "https://graph.microsoft.com/";
    const siteUrl = `${sharepointHostName}:/sites/${siteName}`;

    const client = new client_lib.SharePointClient(tenantId, clientId, clientSecret, resource, siteUrl);
    client.initialize().then(() => {
        switch (action) {
            case "upload_file":
                return uploadFiles(client, uploadPath, filePath);
            case "download_file":
                console.log("To be implemented");
                process.exit(1);
                break;
            default:
                console.error("Unknown action");
                process.exit(1);
        }
    });
}

main().catch(error => {
    console.error("An error occurred:", error);
    process.exit(1);
});