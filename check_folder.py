"""Check contents of a folder in Azure."""
import os
import sys
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

load_dotenv()

credential = DefaultAzureCredential()
blob_service = BlobServiceClient(
    f"https://{os.getenv('STORAGE_ACCOUNT_NAME')}.blob.core.windows.net",
    credential=credential
)
container_client = blob_service.get_container_client(os.getenv('CONTAINER_NAME'))

input_prefix = os.getenv('INPUT_PREFIX')

# If no argument, list all top-level folders with counts
if len(sys.argv) < 2:
    print(f"Top-level folders in {input_prefix}:\n")
    folder_counts = {}
    for blob in container_client.list_blobs(name_starts_with=input_prefix):
        relative = blob.name[len(input_prefix):]
        if '/' in relative:
            folder = relative.split('/')[0]
            folder_counts[folder] = folder_counts.get(folder, 0) + 1
    for folder in sorted(folder_counts.keys()):
        print(f"  {folder_counts[folder]:>6}  {folder}/")
    print(f"\nUsage: python check_folder.py <folder_name>")
else:
    folder = sys.argv[1]
    prefix = input_prefix + folder + '/'
    print(f"Checking: {prefix}\n")
    count = 0
    for blob in container_client.list_blobs(name_starts_with=prefix):
        ext = os.path.splitext(blob.name)[1]
        print(f"{blob.size:>10} {ext or '(none)':>10}  {blob.name}")
        count += 1
        if count >= 30:
            print("\n... (truncated at 30)")
            break
    if count == 0:
        print("No files found.")
