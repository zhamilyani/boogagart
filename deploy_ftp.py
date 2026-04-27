"""FTP deploy script for boogagart.com on ps.kz — uploads site to httpdocs/."""
import os
import sys
from ftplib import FTP_TLS, error_perm

HOST = "srv-plesk65.ps.kz"
USER = os.environ.get("FTP_USER", "")
PASS = os.environ.get("FTP_PASS", "")
REMOTE_BASE = "httpdocs"

ROOT_FILES = [
    "index.html",
    "logo.png",
    "team.png",
    "catering-logo.png",
    "catering-logo.jpg",
    "catering-logo-white.png",
    "catering-logo-preview.png",
]

PHOTOS_DIR = "photos"


def ensure_remote_dir(ftp, path):
    parts = path.split("/")
    cur = ""
    for p in parts:
        if not p:
            continue
        cur = cur + "/" + p if cur else p
        try:
            ftp.mkd(cur)
        except error_perm:
            pass


def upload_file(ftp, local_path, remote_path):
    remote_dir = os.path.dirname(remote_path)
    if remote_dir:
        ensure_remote_dir(ftp, remote_dir)
    with open(local_path, "rb") as f:
        try:
            ftp.storbinary(f"STOR {remote_path}", f)
            size = os.path.getsize(local_path)
            print(f"  OK  {remote_path}  ({size:,} bytes)")
            return True
        except Exception as e:
            print(f"  ERR {remote_path}: {e}")
            return False


def main():
    if not USER or not PASS:
        sys.exit("FTP_USER and FTP_PASS env vars must be set")

    print(f"Connecting to {HOST} as {USER}...")
    try:
        ftp = FTP_TLS(HOST, timeout=60)
        ftp.login(USER, PASS)
        ftp.prot_p()
        print("  TLS OK")
    except Exception as e:
        print(f"  FTPS failed: {e}, trying plain FTP...")
        from ftplib import FTP
        ftp = FTP(HOST, timeout=60)
        ftp.login(USER, PASS)
        print("  plain FTP OK")

    ftp.cwd(REMOTE_BASE)
    print(f"cwd: {REMOTE_BASE}")

    uploaded = 0
    failed = 0

    for rel in ROOT_FILES:
        if not os.path.exists(rel):
            print(f"  SKIP {rel} (not found)")
            continue
        ok = upload_file(ftp, rel, rel)
        uploaded += 1 if ok else 0
        failed += 0 if ok else 1

    if os.path.isdir(PHOTOS_DIR):
        for dirpath, _, filenames in os.walk(PHOTOS_DIR):
            for fn in filenames:
                local = os.path.join(dirpath, fn).replace("\\", "/")
                remote = local
                ok = upload_file(ftp, local, remote)
                uploaded += 1 if ok else 0
                failed += 0 if ok else 1

    ftp.quit()
    print(f"\nUploaded {uploaded}, failed {failed}")


if __name__ == "__main__":
    main()
