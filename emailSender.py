#!/usr/bin/env python3

import glob as _glob
from email.mime.image import MIMEImage as _MIMEImage
from email.mime.text import MIMEText as _MIMEText
from email.mime.multipart import MIMEMultipart as _MIMEMultipart
from jinja2 import Environment as _Environment
from jinja2 import FileSystemLoader as _FileSystemLoader
from jinja2 import Template as _Template
import os as _os
from simple_settings import settings as _settings
import smtplib as _smtplib
import sys as _sys


def get_context(dir_path):
    """Get the default context for the template based on the html files of dir_path"""

    assert type(dir_path) is str and dir_path != "", "dir_path was not string or empty"

    _context = {}
    html_list = _glob.glob(dir_path + _os.sep + "*.html")
    for html_file in html_list:
        key = _os.path.basename(html_file)
        key = key[:-5]  # Remove .html extension
        with open(html_file) as file:
            _context[key] = file.read()
    return _context


def add_image_attachment(outer, dir_path):
    """Add the image attachments to outer message
        outer: root message
        dir_path: directory containing the images.
     """

    png_list = _glob.glob(dir_path + _os.sep + "*.png")
    for png_file in png_list:
        with open(png_file, "rb") as file:
            img = _MIMEImage(file.read())
        img.add_header('Content-ID', f'<{_os.path.splitext(_os.path.basename(png_file))[0]}>')
        outer.attach(img)


def get_template(path_to_template):
    assert type(path_to_template) is str and path_to_template != "", "path_to_template was not string or empty"

    return _jinja2_environment.get_template(_os.path.basename(path_to_template))



if __name__ == "__main__":
    assert len(_sys.argv) == 3, "This script must be running the name of the Jinja template:" \
                                "emailSender.py --simple-settings=settings template.jinja2"
    path_to_template = _os.path.abspath(_sys.argv[2])
    _jinja2_environment = _Environment(trim_blocks=True, loader=_FileSystemLoader(_os.path.dirname(path_to_template))
                                       , autoescape=False)
    with open(_settings.CSV_FILE) as csv_file:
        # Smtp login:
        smtp = _smtplib.SMTP(_settings.SMPT_HOST, _settings.SMPT_PORT)
        try:
            smtp.ehlo()
            smtp.starttls()  # enable TLS
            smtp.ehlo()
            smtp.login(_settings.SMTP_USER, _settings.SMTP_PASSWORD)

            # Read message body:
            template = get_template(path_to_template)

            template_dir_name = _os.path.dirname(path_to_template)
            html_context = get_context(template_dir_name)  # context from html files

            csv_keys = []  # First text line of csv contains the keys used in the template

            for index, line in enumerate(csv_file):
                line = line.strip()
                if line:
                    if not csv_keys:
                        csv_keys = line.split(",")
                        continue

                    values = line.split(",")
                    name = None
                    context = dict(html_context)
                    for value_index, csv_key in enumerate(csv_keys):
                        value = values[value_index]
                        context[csv_key] = value
                        if csv_key == "name":
                            name = value
                        elif csv_key == "email":
                            email = value

                    if name is None:
                        name = context["firstName"] + " " + context["lastName"]

                    # Render items in context because they can be template too
                    context = {key: _Template(context[key]).render(context) for key in context}

                    # Create the container (outer) email message.
                    outer = _MIMEMultipart()
                    outer['Subject'] = context["title"]
                    outer['From'] = _settings.EMAIL_FROM
                    outer['To'] = "%s <%s>" % (name, email)

                    msgAlternative = _MIMEMultipart('alternative')
                    outer.attach(msgAlternative)

                    # Add our message:
                    part1 = _MIMEText(template.render(context), 'html')
                    msgAlternative.attach(part1)

                    # Add the immage:
                    add_image_attachment(outer, template_dir_name)

                    # Now send the message
                    smtp.sendmail(_settings.SMTP_USER, email, outer.as_string())

                    print("Sending to: " + email)
        finally:
            smtp.quit()
