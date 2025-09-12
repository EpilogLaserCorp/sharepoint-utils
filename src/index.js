const client_lib = require("./client")
const core = require('@actions/core');
const glob = require('glob');
const path = require('path');

async function uploadFiles(client, uploadPath, filePath) {
    const fileLines = filePath.split('\n');
    for (const file_line of fileLines) {
        const file_lines = glob.globSync(file_line)
        for (const line of file_lines) {
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
    }
    process.exit(0);
}

async function downloadFiles(client, downloadPath, localPath) {
    const fileLines = downloadPath.split('\n');
    for (const file_line of fileLines) {
        const trimmedLine = file_line.trim();
        if (!trimmedLine) continue;

        console.log(`Processing download path: "${trimmedLine}"`);

        if (trimmedLine.includes('*') || trimmedLine.includes('?')) {
            console.log(`Detected wildcard pattern: ${trimmedLine}`);
            await downloadWildcardPattern(client, trimmedLine, localPath);
        } else {
            const isFolder = await checkIfFolder(client, trimmedLine);
            if (isFolder) {
                console.log(`Detected folder: ${trimmedLine}`);
                await downloadFolder(client, trimmedLine, localPath);
            } else {
                // Preserve the relative path structure
                const relativePath = trimmedLine.startsWith('/') ? trimmedLine.substring(1) : trimmedLine;
                let fullLocalPath = localPath === "." ? relativePath : path.join(localPath, relativePath);
                
                console.log(`Downloading file "${trimmedLine}" to "${fullLocalPath}"...`);
                try {
                    // Ensure the directory structure exists
                    const fs = require('fs').promises;
                    await fs.mkdir(path.dirname(fullLocalPath), { recursive: true });
                    
                    await client.downloadFile(trimmedLine, fullLocalPath);
                    console.log(`Downloaded ${trimmedLine}`);
                } catch (error) {
                    console.error("Error downloading file:", error.message);
                    process.exit(1);
                }
            }
        }
    }
    process.exit(0);
}

async function checkIfFolder(client, remotePath) {
    try {
        const parentPath = path.dirname(remotePath) === '.' ? '' : path.dirname(remotePath);
        const items = await client.listFolder(parentPath);
        const item = items.find(item => item.path === remotePath);
        return item && item.isFolder;
    } catch (error) {
        return false;
    }
}

async function downloadFolder(client, folderPath, localBasePath) {
    try {
        console.log(`Listing files in folder: ${folderPath}`);
        const files = await client.listFolderRecursively(folderPath);
        
        if (files.length === 0) {
            console.log(`No files found in folder: ${folderPath}`);
            return;
        }
        
        console.log(`Found ${files.length} files in folder: ${folderPath}`);
        const results = await client.downloadMultipleFiles(files, localBasePath);
        
        const successCount = results.filter(r => r.success).length;
        const failCount = results.filter(r => !r.success).length;
        
        console.log(`Download completed: ${successCount} successful, ${failCount} failed`);
        
        if (failCount > 0) {
            console.error("Some files failed to download:");
            results.filter(r => !r.success).forEach(r => {
                console.error(`  ${r.file}: ${r.error}`);
            });
        }
    } catch (error) {
        console.error(`Error downloading folder ${folderPath}:`, error.message);
        process.exit(1);
    }
}

async function downloadWildcardPattern(client, pattern, localBasePath) {
    try {
        const patternDir = path.dirname(pattern);
        const searchDir = patternDir === '.' ? '' : patternDir;
        
        console.log(`Searching for files matching pattern: ${pattern} in directory: ${searchDir || '/'}`);
        
        const files = await client.listFolderRecursively(searchDir);
        const matchingFiles = client.matchWildcardPattern(files, pattern);
        
        if (matchingFiles.length === 0) {
            console.log(`No files found matching pattern: ${pattern}`);
            return;
        }
        
        console.log(`Found ${matchingFiles.length} files matching pattern: ${pattern}`);
        const results = await client.downloadMultipleFiles(matchingFiles, localBasePath);
        
        const successCount = results.filter(r => r.success).length;
        const failCount = results.filter(r => !r.success).length;
        
        console.log(`Download completed: ${successCount} successful, ${failCount} failed`);
        
        if (failCount > 0) {
            console.error("Some files failed to download:");
            results.filter(r => !r.success).forEach(r => {
                console.error(`  ${r.file}: ${r.error}`);
            });
        }
    } catch (error) {
        console.error(`Error downloading files with pattern ${pattern}:`, error.message);
        process.exit(1);
    }
}

function getInputValue(key) {
    // Try command-line arguments first, then GitHub Actions inputs
    if (process.argv.length > 2) {
        switch (key) {
            case 'action': return process.argv[2];
            case 'site_name': return process.argv[3];
            case 'host_name': return process.argv[4];
            case 'tenant_id': return process.argv[5];
            case 'client_id': return process.argv[6];
            case 'client_secret': return process.argv[7];
            case 'upload_path': return process.argv[8];
            case 'file_path': return process.argv[9];
            case 'download_path': return process.argv[8];
            case 'local_path': return process.argv[9] || '.';
            default: // Do nothing and fall through to GitHub Actions input
        }
    }
    return core.getInput(key);
}

try {
    const action = getInputValue('action')
    const siteName = getInputValue('site_name')
    const sharepointHostName = getInputValue('host_name')
    const tenantId = getInputValue('tenant_id')
    const clientId = getInputValue('client_id')
    const clientSecret = getInputValue('client_secret')
    const uploadPath = getInputValue('upload_path')
    const filePath = getInputValue('file_path')
    const downloadPath = getInputValue('download_path')
    const localPath = getInputValue('local_path') || '.'
    const maxRetry = core.getInput('max_retries') || 3
    const resource = "https://graph.microsoft.com/";
    const siteUrl = `${sharepointHostName}:/sites/${siteName}`;

    const client = new client_lib.SharePointClient(tenantId, clientId, clientSecret, resource, siteUrl);
    client.initialize().then(() => {
        switch (action) {
            case "upload_file":
                return uploadFiles(client, uploadPath, filePath);
            case "download_file":
                return downloadFiles(client, downloadPath, localPath);
            default:
                console.error("Unknown action");
                process.exit(1);
        }
    });
} catch (error) {
    core.setFailed(error.message);
}