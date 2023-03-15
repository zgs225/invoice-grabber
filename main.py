import json
from imaplib import IMAP4_SSL
import datetime
import email
import os
from selenium.webdriver.common.by import By
import socks
import re
import requests

from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
import time


def load_config():
    # Load config from json file
    with open("config.json") as f:
        return json.load(f)


def check_and_rename_file(dirpath, filepath):
    filename = os.path.basename(filepath)
    base, ext = os.path.splitext(filename)
    i = 1
    while os.path.exists(os.path.join(dirpath, filename)):
        filename = f"{base}_{i}{ext}"
        i += 1
    return os.path.join(dirpath, filename)


def try_download_file_from_quanjia(cfg, msg):
    # Get the email subject
    subject = msg["Subject"]
    if subject is None:
        print("No subject for message")
        return

    if subject.startswith("=?"):
        subject = email.header.decode_header(subject)[0][0]
        try:
            subject = subject.decode("utf-8")
        except:
            print("Cannot decode subject")
            return False

    if "顶全便利店" not in subject:
        return False

    print(f"Subject: {subject} (from 全家便利店)")

    body = msg.get_payload(decode=True)
    if body is None:
        print("No body for message")
        return True
    body = body.decode("utf-8")
    if 'href="' not in body:
        print("No href in body")
        return True

    # Get the download link
    result = re.search(r"href=['\"](https?://fpj.datarj.com/[^\s\"']*)['\"]", body)
    if result is None:
        print("No download page link found")
        return True

    download_link = result.group(1)
    print(f"Download page link: {download_link}")

    # request download link, match r"href=['\"](https?://fpj.datarj.com/e-invoice-file/[^'\"]+.pdf)['\"]" from response body
    opts = Options()
    opts.add_argument("--headless")
    browser = Firefox(options=opts)
    browser.implicitly_wait(30)

    try:
        browser.get(download_link)
        time.sleep(1)
    except Exception as e:
        print(f"Error opening download page: {e}")
        return True

    links = browser.find_elements(By.TAG_NAME, "a")

    for link in links:
        href = link.get_dom_attribute("href")

        if href is None:
            continue

        href = href.strip()
        result = re.match(r"https?://fpj.datarj.com/e-invoice-file/[^'\"]+.pdf", href)
        if result is None:
            continue

        print(f"Found download link: {href}")
        print("Downloading...")

        filename = os.path.basename(href)
        download_dir = cfg["download_dir"]
        response = requests.get(href)

        if response.status_code != 200:
            print("Error downloading file")
            continue

        filepath = check_and_rename_file(download_dir, filename)

        os.makedirs(download_dir, exist_ok=True)
        with open(filepath, "wb+") as f:
            f.write(response.content)

        print(f"Downloaded {filename} to {filepath}")

    browser.close()

    return True


def roam(cfg):
    # Set proxy if needed
    if cfg["proxy"]:
        print("Setting proxy...")
        socks.set_default_proxy(socks.SOCKS5, cfg["proxy"])

    # Connect to the server
    print("Connecting to server {}...".format(cfg["server"]))
    M = IMAP4_SSL(cfg["server"], timeout=10)

    print("Logging in...")
    M.login(cfg["username"], cfg["password"])

    # Select the inbox
    M.select("INBOX")

    since = cfg["since"]
    since_date = datetime.datetime.strptime(since, "%Y-%m-%d").date()
    since = since_date.strftime("%d-%b-%Y")
    crit = f'(SINCE "{since}")'

    print(f"Searching for messages since {since}...")
    typ, data = M.search(None, crit)

    if typ != "OK":
        print("Error searching for messages")
        return

    print(f"Found {len(data[0].split())} messages")

    for num in data[0].split():
        typ, data = M.fetch(num, "(RFC822)")
        if typ != "OK":
            print("Error fetching message")
            continue
        if data is None:
            print("No data for message")
            continue
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        if try_download_file_from_quanjia(cfg, msg):
            continue

        for part in msg.walk():
            if part.get_content_type() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue
            filename = part.get_filename()
            if filename is None:
                continue

            # decode the filename, if needed
            if filename.startswith("=?"):
                filename = email.header.decode_header(filename)[0][0]
                filename = filename.decode("utf-8")

            if not filename.endswith(".pdf"):
                print(f"Skipping {filename} (not a PDF)")
                continue

            download_dir = cfg["download_dir"]
            os.makedirs(download_dir, exist_ok=True)
            filepath = check_and_rename_file(download_dir, filename)
            with open(filepath, "wb+") as f:
                f.write(part.get_payload(decode=True))
            print(f"Downloaded {filename} to {filepath}")

    # Close the connection
    M.close()
    M.logout()


def main():
    cfg = load_config()
    roam(cfg)


if __name__ == "__main__":
    main()
