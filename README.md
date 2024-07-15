# Sharepoint Utilities

Perform various Sharepoint site operations using client ID.

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
      - name: Create Test File
        run: |-
          echo "Hello world!" > foo.txt
          echo "Hello world1!" > foo1.txt
          echo "Hello world2!" > foo2.txt
          echo "Hello world3!" > foo3.txt
          echo "Hello world4!" > foo4.txt
      - name: Check Test File
        run: cat foo.txt
      - name: Send to Sharepoint
        uses: epiloglasercorp/sharepoint-utils@main
        with:
          action: "upload_file"
          file_path: 'foo.txt'
          host_name: 'my_share.sharepoint.com'
          site_name: 'my_site'
          upload_path: '/tmp/'
          tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
          client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
          client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
      - name: Send Multi to Sharepoint
        uses: epiloglasercorp/sharepoint-utils@main
        with:
          action: "upload_file"
          file_path: |-
            foo.txt
            foo1.txt
            foo2.txt
            foo3.txt
            foo4.txt
          host_name: 'my_share.sharepoint.com'
          site_name: 'my_site'
          upload_path: '/tmp/'
          tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
          client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
          client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```
Inspired by: https://github.com/cringdahl/sharepoint-file-upload-action