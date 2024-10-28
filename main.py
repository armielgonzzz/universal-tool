import os
import requests
import shutil
import subprocess
import sys
import zipfile
from io import BytesIO
from dotenv import load_dotenv
from interface.display import main as start_ui

load_dotenv(dotenv_path='misc/.env')

# Replace these with your repository details
GITHUB_USER = os.getenv("USER")
GITHUB_REPO = os.getenv("REPO")
GITHUB_API_URL = f'https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest'

def get_latest_release():
    response = requests.get(GITHUB_API_URL)
    if response.status_code == 200:
        release_data = response.json()
        version = release_data['tag_name']
        main_exe_url = None
        source_zip_url = None
        
        # Locate download URLs for main.exe and source code zip
        for asset in release_data['assets']:
            if asset['name'] == 'main.exe':
                main_exe_url = asset['browser_download_url']
                source_zip_url = release_data['zipball_url']
        
        return version, main_exe_url, source_zip_url
    return None, None, None

def download_new_version(main_exe_url, source_zip_url):
    # Download main.exe
    with requests.get(main_exe_url, stream=True) as r:
        with open("main.exe", 'wb') as f:
            shutil.copyfileobj(r.raw, f)
    print("Downloaded new version of main.exe.")

    # Download and extract Source code (zip)
    response = requests.get(source_zip_url)
    with zipfile.ZipFile(BytesIO(response.content)) as zip_file:
        # Extract files to a temporary directory
        temp_dir = "temp_source"
        os.makedirs(temp_dir, exist_ok=True)
        zip_file.extractall(temp_dir)
        
        # Identify the main folder within the zip (the repo name folder)
        repo_root = os.path.join(temp_dir, os.listdir(temp_dir)[0])

        # Update only the files within this main folder, preserving the structure
        for root, _, files in os.walk(repo_root):
            for file in files:
                source_path = os.path.join(root, file)
                
                # Remove the repo folder prefix to maintain original structure
                rel_path = os.path.relpath(source_path, start=repo_root)
                
                # Destination path in the main directory
                dest_path = os.path.join(".", rel_path)
                
                # Create necessary directories in the main folder structure
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                
                # Move (or overwrite) the file from temp_source to the main directory
                shutil.move(source_path, dest_path)
        
        # Clean up the temporary directory after update
        shutil.rmtree(temp_dir)
    
    print("Source code update successful!\nRestarting the tool...")

def check_for_updates():
    local_version = open('version.txt').read().strip()
    latest_version, main_exe_url, source_zip_url = get_latest_release()

    print(f"Lead Management Tools {local_version}")

    if latest_version and local_version != latest_version:
        print(f"New Tool Version available: {latest_version}\nUpdating...")
        download_new_version(main_exe_url, source_zip_url)

        # Update the version.txt file with the latest version
        with open('version.txt', 'w') as version_file:
            version_file.write(latest_version)

        return True
    else:
        return False

def run_updater():
    subprocess.Popen(["update/run_update.exe", os.path.basename(sys.argv[0])])
    sys.exit()

def main():
    if check_for_updates():
        run_updater()
    else:
        start_ui()

if __name__ == "__main__":
    main()
