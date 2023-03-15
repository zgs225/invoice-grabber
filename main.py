import csv
import datetime
import email
from imaplib import IMAP4_SSL
import json
import os
import re
import shutil
import sys
import time

from aip import AipOcr
import requests
from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
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
    util = datetime.datetime.today().strftime("%d-%b-%Y")
    if 'util' in cfg:
        util_date = datetime.datetime.strptime(cfg["util"], "%Y-%m-%d").date()
        util = util_date.strftime("%d-%b-%Y")
    crit = f'(SINCE "{since}" BEFORE "{util}")'

    print(f"Searching for messages since {since} util {util}...")
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


def get_file_content(filepath):
    with open(filepath, "rb+") as f:
        return f.read()


def recognize_invoices(cfg):
    client = AipOcr(
        cfg["baidu_ocr"]["app_id"],
        cfg["baidu_ocr"]["api_key"],
        cfg["baidu_ocr"]["secret_key"],
    )

    files = os.listdir(cfg["download_dir"])
    entities = {}

    for filename in files:
        if not filename.endswith(".pdf"):
            continue

        print(f"Recognizing {filename}...")
        filepath = os.path.join(cfg["download_dir"], filename)
        pdf_file = get_file_content(filepath)
        res = client.vatInvoicePdf(pdf_file)

        if "error_code" in res:
            print("Error recognizing file: " + res["error_msg"])
            continue

        entities[filename] = res

    output_dir = cfg["output"]["dir"]
    recognized_json = cfg["output"]["recognized_json"]
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, recognized_json)
    with open(filepath, "w+", encoding="utf-8") as f:
        json.dump(entities, f, indent=4, ensure_ascii=False)
        print(f"Saved recognized results to {filepath}")


def load_recognized_invoices(cfg):
    output_dir = cfg["output"]["dir"]
    recognized_json = cfg["output"]["recognized_json"]
    filepath = os.path.join(output_dir, recognized_json)
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def rename_invoices(cfg):
    invoices = load_recognized_invoices(cfg)
    download_dir = cfg["download_dir"]
    output_dir = cfg["output"]["dir"]
    name = cfg["output"]["name"]

    os.makedirs(output_dir, exist_ok=True)

    for filename, invoice in invoices.items():
        total_amount = invoice["words_result"]["AmountInFiguers"]
        filename2 = f"{name} 发票 {total_amount}.pdf"
        file1 = os.path.join(download_dir, filename)
        file2 = check_and_rename_file(output_dir, filename2)
        shutil.copy(file1, file2)
        print(f"Copied {filename} to {file2}")


def generate_excel_records(cfg):
    invoices = load_recognized_invoices(cfg).values()
    output_dir = cfg["output"]["dir"]
    csv_file = cfg["output"]["result_csv"]
    name = cfg["output"]["name"]
    csv_filepath = os.path.join(output_dir, csv_file)
    data = []
    index = 0

    sorted_invoices = sorted(invoices, key=lambda x: x["words_result"]["InvoiceDate"])

    for invoice in sorted_invoices:
        total_amount = float(invoice["words_result"]["AmountInFiguers"])
        index += 1
        invoice_date = invoice["words_result"]["InvoiceDate"]
        seller_name = invoice["words_result"]["SellerName"]

        data.append(
            {
                "序号": index,
                "用餐日期": invoice_date[5:],
                "用餐缘由": "加班",
                "用餐商户名称": seller_name,
                "总人数": "1",
                "人员名单": name,
                "发票金额": total_amount,
                "按标准应报销金额": 50 if total_amount > 50 else total_amount,
                "责任部门": "区块链创新应用部",
                "经办人/报销人": name,
            }
        )

    if len(data) == 0:
        print("No data to generate")
        return

    os.makedirs(output_dir, exist_ok=True)

    with open(csv_filepath, "w+") as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)
        print(f"Saved result to {csv_filepath}")


def print_usage():
    print("Usage: python3 main.py [--no-email] [--no-ocr] [--no-rename] [--no-excel]")


def main():
    if "--help" in sys.argv or "-h" in sys.argv:
        print_usage()
        return

    cfg = load_config()

    if "--no-email" not in sys.argv:
        roam(cfg)

    if "--no-ocr" not in sys.argv:
        recognize_invoices(cfg)

    if "--no-rename" not in sys.argv:
        rename_invoices(cfg)

    if "--no-excel" not in sys.argv:
        generate_excel_records(cfg)


if __name__ == "__main__":
    main()
