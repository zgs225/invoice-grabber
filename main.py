import json
from imaplib import IMAP4_SSL
import datetime
import email
import os
import socks

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

            # create download dir if needed
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
