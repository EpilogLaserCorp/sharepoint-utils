#!/bin/bash

# These should be set in the environment or retrieved from a secure location
if [ -z "$SHAREPOINT_SITE_NAME" ]; then
    echo ""
    read -rp "Sharepoint Site Name: " SHAREPOINT_SITE_NAME
    echo ""
fi

if [ -z "$SHAREPOINT_HOST_NAME" ]; then
    echo ""
    read -rp "Sharepoint Host Name: " SHAREPOINT_HOST_NAME
    echo ""
fi

if [ -z "$TENANT_ID" ]; then
    echo ""
    read -rp "Sharepoint Tenant ID: " TENANT_ID
    echo ""
fi

if [ -z "$CLIENT_ID" ]; then
    echo ""
    read -rp "Sharepoint Client ID: " CLIENT_ID
    echo ""
fi

# If still no secret found, prompt user and store it
if [ -z "$SHAREPOINT_CLIENT_SECRET" ]; then
    echo ""
    read -rsp "Sharepoint Client Secret: " SHAREPOINT_CLIENT_SECRET
    echo ""
fi


dd if=/dev/urandom of=big.file bs=200M count=1
dd if=/dev/urandom of=small.file bs=1M count=1

# Get the md5 sums of the files for later comparison
md5sum big.file > big.md5
md5sum small.file > small.md5

node src/test.js \
    upload_file \
    "$SHAREPOINT_SITE_NAME" \
    "$SHAREPOINT_HOST_NAME" \
    "$TENANT_ID" \
    "$CLIENT_ID" \
    "$SHAREPOINT_CLIENT_SECRET" \
    "/tmp/" \
    "big.file"

node src/test.js \
    upload_file \
    "$SHAREPOINT_SITE_NAME" \
    "$SHAREPOINT_HOST_NAME" \
    "$TENANT_ID" \
    "$CLIENT_ID" \
    "$SHAREPOINT_CLIENT_SECRET" \
    "/tmp/" \
    "small.file"

rm big.file small.file

# Download the files again to verify
node src/index.js \
    download_file \
    "$SHAREPOINT_SITE_NAME" \
    "$SHAREPOINT_HOST_NAME" \
    "$TENANT_ID" \
    "$CLIENT_ID" \
    "$SHAREPOINT_CLIENT_SECRET" \
    "/tmp/big.file" \
    "."

node src/index.js \
    download_file \
    "$SHAREPOINT_SITE_NAME" \
    "$SHAREPOINT_HOST_NAME" \
    "$TENANT_ID" \
    "$CLIENT_ID" \
    "$SHAREPOINT_CLIENT_SECRET" \
    "/tmp/small.file" \
    "."

mv ./tmp/big.file .
mv ./tmp/small.file .

# Verify the md5 sums match
md5sum -c big.md5
md5sum -c small.md5