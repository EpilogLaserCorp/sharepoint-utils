const client_lib = require("./client")
const core = require('@actions/core');
const github = require('@actions/github');
const path = require('path');

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

try {
    const action = core.getInput('action')
    const siteName = core.getInput('site_name')
    const sharepointHostName = core.getInput('host_name')
    const tenantId = core.getInput('tenant_id')
    const clientId = core.getInput('client_id')
    const clientSecret = core.getInput('client_secret')
    const uploadPath = core.getInput('upload_path')
    const filePath = core.getInput('file_path')
    const maxRetry = core.getInput('max_retries') || 3
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
} catch (error) {
    core.setFailed(error.message);
}