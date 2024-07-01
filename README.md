# Sharepoint Utilities

Perform various Sharepoint site operations using client ID.
**NOTE: Will not run on Windows workers**

## Variables
The following environment variables & secrets must be defined.

If your full Sharepoint upload path is `https://example.sharepoint.com/sites/example_site/Shared%20Documents/my/files`, the following would be defined:

* `host_name`
  * `'example.sharepoint.com'`
* `site_name`
  * `'example_site'`
* `upload_path` (always starts with a `/`)
  * `'/my/files'`


The following will be provided to you by your Sharepoint administrator when you ask for a client ID. A reminder: _put secrets in **Settings/Security/Secrets and variables/Actions**_

* `tenant_id`
* `client_id`
* `client_secret`

You will also need to provide the file or files being sent:

* `file_path`
  * A glob; something like `file.txt` or `*.md`
  * Files can only be in the repository directory and cannot be absolute.

## Example action.yml

```yml
name: example-file-upload
on: workflow_dispatch
jobs:
  get_report:
    runs-on: ubuntu-latest
    steps:
      - name: Create Test File to Upload
        run: touch /tmp/foo.txt
      - name: Upload to Sharepoint
        uses: EpilogLaserCorp/sharepoint-utils@main
        with:
          file_path: "/tmp/foo.txt"
          host_name: 'your.sharepoint.com'
          site_name: 'your_site'
          upload_path: '/my-path-that-starts-with-slash/many_files/big_path/'
          tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
          client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
          client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```
