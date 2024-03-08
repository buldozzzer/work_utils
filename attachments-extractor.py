#!/usr/bin/env python

__description__ = "Extract attachments from .eml files"
__author__ = "Tariq Alashaikh"
__date__ = "07/12/2023"

import argparse
import email
import email.generator as generator
import email.policy as policy
import glob
import os
import random
import re
import string
import sys


def get_unique_path(path):
    filename, extension = os.path.splitext(path)
    while os.path.exists(path):
        path = (
            f"{filename}_{''.join([random.choice(string.ascii_letters + string.digits) for _ in range(5)])}{extension}"
        )
    return path


def sanitize_filename(filename):
    return re.sub(r"[/\\|\[\]\{\}:<>+=;,?!*\"~#$%&@']", "_", filename)


def save_attachment(location, attachment, keep):
    filename = sanitize_filename(attachment.get_filename())
    path = os.path.join(location, filename)

    if keep:
        path = get_unique_path(path)

    if attachment.get_content_type() == "message/rfc822":
        data = attachment.get_payload(0)
        with open(path, "w") as file:
            gen = generator.Generator(file)
            gen.flatten(data)
    else:
        data = attachment.get_payload(decode=True)
        with open(path, "wb") as file:
            file.write(data)


def get_attachments(msg):
    if msg.is_multipart():
        return [item for item in msg.iter_attachments()]
    elif msg.is_attachment() or msg.get_content_disposition() == "inline":
        return [msg]


def parse_arguments():
    parser = argparse.ArgumentParser(description=__description__, add_help=False)
    parser.add_argument("email", type=str, nargs="+", help="path to .eml file(s)")
    parser.add_argument(
        "-o",
        "--organize",
        action="store_true",
        dest="organize",
        help="organize attachments into subfolders based on .eml filename",
    )
    parser.add_argument(
        "-k", "--keep", action="store_true", dest="keep", help="keep duplicate files by appending a random suffix"
    )
    parser.add_argument("-h", "--help", action="help", help="show this help message and exit")
    return parser.parse_args()


def main():
    args = parse_arguments()
    emls = [eml for items in args.email for eml in glob.glob(items)]

    for eml in emls:
        try:
            with open(eml, "rb") as file:
                msg = email.message_from_binary_file(file, policy=policy.default)
        except Exception:
            print(f"Error: {eml} <failed to read file>", file=sys.stderr)
            continue

        attachments = get_attachments(msg)

        if not attachments:
            print(f"{eml} (0)")
            continue

        try:
            location = (
                "attachments"
                if not args.organize
                else os.path.join("attachments", os.path.basename(os.path.splitext(eml)[0]))
            )

            if args.organize and args.keep:
                location = get_unique_path(location)

            os.makedirs(location, exist_ok=True)

            for attachment in attachments:
                save_attachment(location, attachment, args.keep)

            print(f"{eml} ({len(attachments)})")

        except Exception as err:
            print(f"Error: {os.path.basename(eml)} <{err}>", file=sys.stderr)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)