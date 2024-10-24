import os
import requests
import shutil
import subprocess
import sys
from interface.display import main as start_ui

# Replace these with your repository details
GITHUB_USER = 'armielgonzzz'
GITHUB_REPO = 'universal-tool'
GITHUB_API_URL = f'https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest'

def get_latest_release():
    response = requests.get(GITHUB_API_URL)
    if response.status_code == 200:
        release_data = response.json()
        version = release_data['tag_name']
        for asset in release_data['assets']:
            if asset['name'] == 'CM Tools.exe':
                download_url = asset['browser_download_url']
                return version, download_url
    return None, None

def download_new_version(download_url):
    with requests.get(download_url, stream=True) as r:
        with open("new_version.exe", 'wb') as f:
            shutil.copyfileobj(r.raw, f)
    print("Downloaded new version.")

def check_for_updates():
    local_version = open('version.txt').read().strip()
    latest_version, download_url = get_latest_release()

    if latest_version and local_version != latest_version:
        print(f"New version available: {latest_version}, updating...")
        download_new_version(download_url)
        return True
    else:
        return False

def run_updater():
    subprocess.Popen(["updater.exe", os.path.basename(sys.argv[0])])
    sys.exit()

def main():
    if check_for_updates():
        run_updater()
    else:
        start_ui()

if __name__ == "__main__":
    main()
