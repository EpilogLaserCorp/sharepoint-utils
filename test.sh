#!/bin/bash
SHAREPOINT_SITE_NAME="swengteam"
SHAREPOINT_HOST_NAME="epiloglaser.sharepoint.com"
dd if=/dev/urandom of=big.file bs=200M count=1
dd if=/dev/urandom of=small.file bs=1M count=1
node src/test.js \
    upload_file \
    $SHAREPOINT_SITE_NAME \
    $SHAREPOINT_HOST_NAME \
    $TENANT_ID \
    $CLIENT_ID \
    $CLIENT_SEC \
    "/tmp/" \
    "big.file"

node src/test.js \
    upload_file \
    $SHAREPOINT_SITE_NAME \
    $SHAREPOINT_HOST_NAME \
    $TENANT_ID \
    $CLIENT_ID \
    $CLIENT_SEC \
    "/tmp/" \
    "small.file"
