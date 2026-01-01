import json
import os
import time
import hashlib
import threading
import re
import subprocess
import xml.etree.ElementTree as ET

import requests
from PyQt6.QtCore import QObject, pyqtSignal, QCoreApplication

def tr(text: str) -> str:
    return QCoreApplication.translate("VersionManager", text)

class VersionManager(QObject):
    update_available = pyqtSignal(dict)
    update_progress = pyqtSignal(int)
    update_error = pyqtSignal(str)
    update_complete = pyqtSignal()
    update_check_finished = pyqtSignal()
    update_check_started = pyqtSignal()

    def __init__(self, config_path="config/version.json", repo_owner="Haraguse", repo_name="Kazuha"):
        super().__init__()
        self.config_path = config_path
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.current_version_info = self.load_version_info()
        self.latest_release_info = None
        self._is_checking = False

    def load_version_info(self):
        if not os.path.exists(self.config_path):
            return None
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data
        except Exception as e:
            print(f"Error reading version info: {e}")
            return None

    def save_version_info(self, info):
        try:
            temp_path = self.config_path + ".tmp"
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(info, f, indent=4)
                f.flush()
                os.fsync(f.fileno())
            
            if os.path.exists(self.config_path):
                os.remove(self.config_path)
            os.rename(temp_path, self.config_path)
            
        except Exception as e:
            print(f"Error saving version info: {e}")

    def check_for_updates(self):
        if self._is_checking:
            return
        
        # If we already have the latest release info cached, emit it immediately
        # This prevents "empty" logs when opening settings if check was already done
        if self.latest_release_info:
            print("Using cached release info")
            self._handle_release_data(self.latest_release_info)
            return

        self._is_checking = True
        self.update_check_started.emit()
        
        thread = threading.Thread(target=self._check_for_updates_thread, daemon=True)
        thread.start()

    def _check_for_updates_thread(self):
        try:
            # First, check latest release for version comparison
            url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/latest"
            headers = {'Accept': 'application/vnd.github.v3+json'}
            
            latest_release = None
            
            # Try GitHub API
            max_retries = 2
            for attempt in range(max_retries):
                try:
                    response = requests.get(url, headers=headers, timeout=10)
                    if response.status_code == 200:
                        latest_release = response.json()
                        break
                except Exception:
                    pass
                time.sleep(1)

            # Try mirrors if GitHub API fails
            if not latest_release:
                mirrors = ["api.bgithub.xyz", "api.github-api.com"]
                for mirror in mirrors:
                    try:
                        mirror_url = url.replace("api.github.com", mirror)
                        response = requests.get(mirror_url, headers=headers, timeout=10)
                        if response.status_code == 200:
                            latest_release = response.json()
                            break
                    except Exception:
                        continue
            
            # Fallback to Atom Feed if API fails completely
            if not latest_release:
                self._check_updates_via_feed()
                return

            # Now handle the version comparison
            # We want to show the changelog of the *current* version if we are up-to-date
            # OR show the changelog of the *latest* version if there is an update
            
            self._handle_release_data(latest_release)

        finally:
            self._is_checking = False
            try:
                self.update_check_finished.emit()
            except Exception:
                pass

    def _get_release_by_tag(self, tag_name):
        # Helper to fetch specific tag release info
        url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/tags/{tag_name}"
        headers = {'Accept': 'application/vnd.github.v3+json'}
        
        # Try direct GitHub API first
        try:
            print(f"Fetching tag: {url}")
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                return response.json()
        except Exception as e:
            print(f"Direct fetch failed: {e}")

        # Try mirrors
        mirrors = ["api.bgithub.xyz", "api.github-api.com"]
        for mirror in mirrors:
            try:
                mirror_url = url.replace("api.github.com", mirror)
                print(f"Fetching tag via mirror: {mirror_url}")
                response = requests.get(mirror_url, headers=headers, timeout=10)
                if response.status_code == 200:
                    return response.json()
            except Exception as e:
                print(f"Mirror {mirror} fetch failed: {e}")
                continue
                
        return None

    def _handle_release_data(self, release_data):
        try:
            if not release_data:
                print("Release data is empty")
                return

            name = release_data.get('name', '')
            tag_name = release_data.get('tag_name', '')
            body = release_data.get('body', '')
            
            # Ensure body is string
            if not isinstance(body, str):
                body = str(body) if body is not None else ""
            
            print(f"Handling release: {name} ({tag_name})")
            print(f"Body length: {len(body)}")
            
            remote_code = None
            match = re.search(r'（(\d+)）', name) or re.search(r'（(\d+)）', tag_name)
            if match:
                remote_code = int(match.group(1))
            else:
                match = re.search(r'\((\d+)\)', name) or re.search(r'\((\d+)\)', tag_name)
                if match:
                    remote_code = int(match.group(1))
            
            local_version_info = self.current_version_info or {}
            local_code = int(local_version_info.get('versionCode', '0'))
            local_tag = local_version_info.get('versionName', '')
            
            # If we found a valid remote code and it's newer, show that release
            if remote_code is not None and remote_code > local_code:
                self.latest_release_info = release_data
                self.update_available.emit({
                    'version': tag_name,
                    'versionCode': remote_code,
                    'name': name,
                    'body': body,
                    'assets': release_data.get('assets', [])
                })
            else:
                print(f"No new version. Remote: {remote_code}, Local: {local_code}")
                
                # User requested to see the LATEST release log even if up-to-date
                # Simplify logic: Just show the latest release body.
                # Do NOT try to fetch specific tag which might fail.
                
                # IMPORTANT: Pass local_code as versionCode to indicate NO update
                self.update_available.emit({
                    'version': tag_name, # Show latest version tag name (e.g. v1.0.5)
                    'versionCode': local_code, # Use LOCAL code to tell UI "we are on this version"
                    'name': name,
                    'body': body, # Show latest body
                    'assets': []
                })
        except Exception as e:
            print(f"Handle release data error: {e}")

    def _check_updates_via_feed(self):
        try:
            feed_url = f"https://github.com/Haraguse/Kazuha/releases.atom"
            sources = [
                feed_url,
                feed_url.replace("github.com", "bgithub.xyz"),
                feed_url.replace("github.com", "kkgithub.com")
            ]
            
            r = None
            for src in sources:
                try:
                    r = requests.get(src, timeout=10)
                    if r.status_code == 200:
                        break
                except Exception:
                    continue
            
            if not r or r.status_code != 200:
                return

            root = ET.fromstring(r.text)
            ns = {'atom': 'http://www.w3.org/2005/Atom'}
            entry = root.find('atom:entry', ns)
            if entry is None:
                return
                
            title_el = entry.find('atom:title', ns)
            content_el = entry.find('atom:content', ns)
            tag_title = title_el.text.strip() if title_el is not None and title_el.text else ''
            
            body = content_el.text or '' if content_el is not None else ''
            
            body = re.sub(r'<br\s*/?>', '\n', body)
            body = re.sub(r'<[^>]+>', '', body)
            
            self._handle_release_data({
                'tag_name': tag_title.split()[0] if tag_title else 'v0.0.0',
                'name': tag_title,
                'body': body,
                'assets': []
            })
        except Exception:
            pass

    def download_and_install(self, asset_url, sha256_hash=None):
        threading.Thread(target=self._download_and_install_thread, args=(asset_url, sha256_hash), daemon=True).start()

    def _download_and_install_thread(self, asset_url, sha256_hash=None):
        try:
            urls = [asset_url]
            if "github.com" in asset_url:
                urls.append(f"https://ghproxy.net/{asset_url}")
                urls.append(asset_url.replace("github.com", "kkgithub.com"))
            
            response = None
            last_err = None
            for url in urls:
                try:
                    response = requests.get(url, stream=True, timeout=20)
                    if response.status_code == 200:
                        break
                except Exception as e:
                    last_err = e
                    continue
            
            if not response or response.status_code != 200:
                self.update_error.emit(tr("无法连接到下载服务器: {0}").format(str(last_err or "Unknown")))
                return

            total_size = int(response.headers.get('content-length', 0))
            downloaded_size = 0
            
            temp_file = os.path.join(os.environ.get('TEMP', '.'), "update_installer.exe")
            
            with open(temp_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        if total_size > 0:
                            progress = int((downloaded_size / total_size) * 100)
                            self.update_progress.emit(progress)
            
            if sha256_hash:
                if not self.verify_hash(temp_file, sha256_hash):
                    self.update_error.emit(tr("文件哈希校验失败"))
                    return
                else:
                    print("Hash verification passed.")

            self.install_update(temp_file)
            return True
            
        except Exception as e:
            self.update_error.emit(tr("下载或安装更新时出错: {e}").format(e=str(e)))
            return False

    def verify_hash(self, file_path, expected_hash):
        if not expected_hash:
            return True
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest().lower() == expected_hash.lower()

    def install_update(self, installer_path):
        try:
            subprocess.Popen([installer_path, "/S"], shell=True)
            self.update_complete.emit()
        except Exception as e:
            self.update_error.emit(tr("安装失败: {e}").format(e=e))
