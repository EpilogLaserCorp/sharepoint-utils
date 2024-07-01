#!/bin/bash
docker build -t test .
dd if=/dev/urandom of=big.file bs=200M count=1
dd if=/dev/urandom of=small.file bs=1M count=1
docker run --rm -v ./:/usr/src/app/mnt test \
    upload_file \
    $SHAREPOINT_SITE_NAME \
    $SHAREPOINT_HOST_NAME \
    $TENANT_ID \
    $CLIENT_ID \
    $CLIENT_SEC \
    "/tmp/" \
    "mnt/big.file"

docker run --rm -v ./:/usr/src/app/mnt test \
    upload_file \
    $SHAREPOINT_SITE_NAME \
    $SHAREPOINT_HOST_NAME \
    $TENANT_ID \
    $CLIENT_ID \
    $CLIENT_SEC \
    "/tmp/" \
    "mnt/small.file"
