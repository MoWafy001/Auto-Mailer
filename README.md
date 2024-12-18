# Auto Mailer
Python Tkinter app used to automate sending emails using SendGrid. It works by extracting emails from a csv/xls/xlsx file then automatically send an email to each user.

<!-- vim-markdown-toc GFM -->

* [Usage](#usage)
* [Download](#download)
* [Install](#install)
* [Running The App](#running-the-app)

<!-- vim-markdown-toc -->

# Usage
1. Select a file to extract the data from. `(People Tab)`
2. Write an email template using _HTML_ or _plain text_. `(Template Tab)`
3. Add your _sender email_, _SendGrid Token_, and the _email title_. `(Settings Tab)`
4. Send the emails. `(Send Emails Tab)`

# Download
[releases](https://github.com/MoWafy001/Auto-Mailer/releases)

# Install
You will need **Python3** installed as well as **Pip3** to install and run the app.

Install the libraries listed in `requirements.txt`
```sh
pip install -r requirements.txt
```

# Running The App
To run the app, run `app.py`
```sh
python app.py
```
