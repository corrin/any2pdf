"""Download blobs from a file list to a local directory."""
import argparse
import os
import pathlib
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

load_dotenv()

STORAGE_ACCOUNT_NAME = os.getenv("STORAGE_ACCOUNT_NAME")
CONTAINER_NAME = os.getenv("CONTAINER_NAME")


def main():
    parser = argparse.ArgumentParser(description="Download blobs from a file list")
    parser.add_argument("file_list", type=pathlib.Path, help="Text file with blob paths")
    parser.add_argument("output_dir", type=pathlib.Path, help="Local directory to save files")
    args = parser.parse_args()
    
    # Ensure output directory exists
    args.output_dir.mkdir(parents=True, exist_ok=True)
    
    # Connect to Azure
    credential = DefaultAzureCredential()
    blob_service = BlobServiceClient(
        f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net",
        credential=credential,
    )
    container_client = blob_service.get_container_client(CONTAINER_NAME)
    
    # Read file list
    with open(args.file_list, 'r', encoding='utf-8') as f:
        blob_names = [line.strip() for line in f if line.strip()]
    
    print(f"Downloading {len(blob_names)} files to {args.output_dir}")
    
    for blob_name in blob_names:
        local_name = pathlib.Path(blob_name).name
        local_path = args.output_dir / local_name
        
        try:
            downloader = container_client.download_blob(blob_name)
            with open(local_path, "wb") as f:
                f.write(downloader.readall())
            print(f"OK {local_name}")
        except Exception as e:
            print(f"ERROR {blob_name}: {e}")


if __name__ == "__main__":
    main()
