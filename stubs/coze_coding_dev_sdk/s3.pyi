"""Stub: coze_coding_dev_sdk.s3"""

class S3SyncStorage:
    def __init__(
        self,
        endpoint_url: str = "",
        access_key: str = "",
        secret_key: str = "",
        bucket_name: str = "",
        region: str = "cn-beijing",
    ) -> None: ...

    def upload_file(
        self,
        file_content: bytes,
        file_name: str,
        content_type: str = "application/octet-stream",
    ) -> str:
        """Upload file, return the storage key."""
        ...

    def generate_presigned_url(
        self,
        key: str,
        expire_time: int = 3600,
    ) -> str:
        """Generate a presigned download URL."""
        ...
