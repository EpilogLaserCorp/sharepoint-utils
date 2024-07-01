import requests
import os
import sys
import math

def convert_size(size_bytes):
   if size_bytes == 0:
       return "0B"
   size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
   i = int(math.floor(math.log(size_bytes, 1024)))
   p = math.pow(1024, i)
   s = round(size_bytes / p, 2)
   return "%s %s" % (s, size_name[i])

class SharePointClient:
    def __init__(self, tenant_id, client_id, client_secret, resource_url, site_url):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.resource_url = resource_url
        self.base_url = (
            f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        )
        self.headers = {"Content-Type": "application/x-www-form-urlencoded"}
        self.access_token = (
            self.get_access_token()
        )  # Initialize and store the access token upon instantiation
        self.site_id = self.get_site_id(site_url)

    def get_access_token(self):
        # Body for the access token request
        body = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": self.resource_url + ".default",
        }
        response = requests.post(self.base_url, headers=self.headers, data=body)
        return response.json().get(
            "access_token"
        )  # Extract access token from the response

    def get_site_id(self, site_url):
        # Build URL to request site ID
        full_url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"
        response = requests.get(
            full_url, headers={"Authorization": f"Bearer {self.access_token}"}
        )
        return response.json().get("id")  # Return the site ID

    def get_drive_id(self):
        # Retrieve drive IDs and names associated with a site
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_self.id}/drives"
        response = requests.get(
            drives_url, headers={"Authorization": f"Bearer {self.access_token}"}
        )
        drives = response.json().get("value", [])
        return [(drive["id"], drive["name"]) for drive in drives]

    def get_folder_content(self, drive_id):
        # Get the contents of a folder
        folder_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/root/children"
        response = requests.get(
            folder_url, headers={"Authorization": f"Bearer {self.access_token}"}
        )
        items_data = response.json()
        rootdir = []
        if "value" in items_data:
            for item in items_data["value"]:
                rootdir.append((item["id"], item["name"]))
        return rootdir

    def get_folder_content2(self):
        # Get the contents of a folder
        folder_url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root/children"
        )
        response = requests.get(
            folder_url, headers={"Authorization": f"Bearer {self.access_token}"}
        )
        items_data = response.json()
        rootdir = []
        if "value" in items_data:
            for item in items_data["value"]:
                rootdir.append((item["id"], item["name"]))
        return rootdir

    # Recursive function to browse folders
    def list_folder_contents(self, drive_id, folder_id, level=0):
        # Get the contents of a specific folder
        folder_contents_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/items/{folder_id}/children"
        contents_headers = {"Authorization": f"Bearer {self.access_token}"}
        contents_response = requests.get(folder_contents_url, headers=contents_headers)
        folder_contents = contents_response.json()

        items_list = []  # List to store information

        if "value" in folder_contents:
            for item in folder_contents["value"]:
                if "folder" in item:
                    # Add folder to list
                    items_list.append(
                        {
                            "name": item["name"],
                            "type": "Folder",
                            "mimeType": None,
                            "webUrl": item["webUrl"],
                        }
                    )
                    # Recursive call for subfolders
                    items_list.extend(
                        self.list_folder_contents(
                            site_id, drive_id, item["id"], level + 1
                        )
                    )
                elif "file" in item:
                    # Add file to the list with its mimeType
                    items_list.append(
                        {
                            "name": item["name"],
                            "type": "File",
                            "mimeType": item["file"]["mimeType"],
                            "webUrl": item["webUrl"],
                        }
                    )

        return items_list

    def download_file(self, download_url, local_path, file_name):
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(download_url, headers=headers)
        if response.status_code == 200:
            full_path = os.path.join(local_path, file_name)
            with open(full_path, "wb") as file:
                file.write(response.content)
            print(f"File downloaded: {full_path}")
        else:
            print(
                f"Failed to download {file_name}: {response.status_code} - {response.reason}"
            )

    def download_folder_contents(self, drive_id, folder_id, local_folder_path, level=0):
        # Recursively download all contents from a folder
        folder_contents_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/items/{folder_id}/children"
        contents_headers = {"Authorization": f"Bearer {self.access_token}"}
        contents_response = requests.get(folder_contents_url, headers=contents_headers)
        folder_contents = contents_response.json()

        if "value" in folder_contents:
            for item in folder_contents["value"]:
                if "folder" in item:
                    new_path = os.path.join(local_folder_path, item["name"])
                    if not os.path.exists(new_path):
                        os.makedirs(new_path)
                    self.download_folder_contents(
                        self.site_id, drive_id, item["id"], new_path, level + 1
                    )  # Recursive call for subfolders
                elif "file" in item:
                    file_name = item["name"]
                    file_download_url = f"{resource}/v1.0/sites/{self.site_id}/drives/{drive_id}/items/{item['id']}/content"
                    self.download_file(file_download_url, local_folder_path, file_name)

    def upload_small_file(self, upload_file_path, data):
        folder_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:{upload_file_path}:/content"
        response = requests.put(
            folder_url,
            data=data,
            headers={
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "text/plain",
            },
        )
        return response

    def upload_session(self, file_path):
        folder_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:{file_path}:/createUploadSession"
        response = requests.post(
            folder_url,
            json={
                "@microsoft.graph.conflictBehavior": "replace",
                "name": "my_file.txt",
            },
            headers={
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            },
        )
        return response

    def upload_to_session(self, url, data):
        total_size = len(data)
        # Note: If your app splits a file into multiple byte ranges, the size of each byte range MUST be a multiple of 320 KiB (327,680 bytes).
        # From: https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
        chunk_size = 327680 * 10  # ~30MB
        offset = 0
        ok = True
        while ok and (offset < total_size):
            this_chunk_size = min(chunk_size, total_size - offset)
            data_slice = data[offset : offset + this_chunk_size]
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Length": f"{this_chunk_size}",
                "Content-Range": f"bytes {offset}-{offset + this_chunk_size - 1}/{total_size}",
            }
            response = requests.put(url, headers=headers, data=data_slice)
            offset += this_chunk_size
            if not response.ok:
                ok = False
        return response

    def upload_file(self, upload_file_path, file_path):
        file_size = os.path.getsize(file_path)
        # Max file size is 250MB for the "small" upload path, otherwise upload session needed.
        # https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http
        MAX_SMALL_FILE_SIZE = 1 << 27 # ~137MB
        with open(file_path, "rb") as f:
            if file_size < MAX_SMALL_FILE_SIZE:
                contents = f.read()
                return self.upload_small_file(upload_file_path, contents)
            else:
                response = client.upload_session(upload_file_path)
                response_json = response.json()
                if "uploadUrl" in response_json:
                    total_size = file_size
                    # Note: If your app splits a file into multiple byte ranges, the size of each byte range MUST be a multiple of 320 KiB (327,680 bytes).
                    # Max chunk size is 60MB.
                    # From: https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
                    chunk_size = 327680 * 100  # ~30MB
                    offset = 0
                    ok = True
                    while ok and (offset < total_size):
                        this_chunk_size = min(chunk_size, total_size - offset)
                        data = f.read(chunk_size)
                        headers = {
                            "Authorization": f"Bearer {self.access_token}",
                            "Content-Length": f"{this_chunk_size}",
                            "Content-Range": f"bytes {offset}-{offset + this_chunk_size - 1}/{total_size}",
                        }
                        response = requests.put(
                            response_json["uploadUrl"], headers=headers, data=data
                        )
                        offset += this_chunk_size
                        if not response.ok:
                            ok = False
                            print("Error doing session upload:")
                            print(response.json())
                        print(f"Uploaded {convert_size(offset)}/{convert_size(total_size)}")
                return response


action = sys.argv[1]
site_name = sys.argv[2]
sharepoint_host_name = sys.argv[3]
tenant_id = sys.argv[4]
client_id = sys.argv[5]
client_secret = sys.argv[6]
upload_path = sys.argv[7]
file_path = sys.argv[8]
max_retry = int(sys.argv[9]) if len(sys.argv) > 9 else 3
resource = "https://graph.microsoft.com/"
site_url = f"{sharepoint_host_name}:/sites/{site_name}"

client = SharePointClient(tenant_id, client_id, client_secret, resource, site_url)

match action:
    case "upload_file":
        for line in file_path.splitlines():
            base_name = os.path.basename(line)
            full_upload_path = upload_path if upload_path == "/" else upload_path.rstrip("/")
            full_upload_path = full_upload_path + "/" + base_name
            print(f"Uploading \"{line}\" to \"{full_upload_path}\"...")
            response = client.upload_file(full_upload_path, line)
            if response.ok:
                print(f"Uploaded {line}")
            else:
                print("Error uploading file. Response:")
                print(response.json())
                exit(1)
        exit(0)
    case "download_file":
        print("To be implemented")
        exit(1)
