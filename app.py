#!/usr/bin/env python
# -*- coding: utf-8 -*-
import json
from flask import Flask, request, render_template, redirect
from flask_httpauth import HTTPBasicAuth
from config import ALLOWED_EXTENSIONS, USERS
from file_upload import FileUpload, SummaryHandle


app = Flask(__name__)
auth = HTTPBasicAuth()


@auth.get_password
def get_pw(username):
    if username in USERS:
        return USERS.get(username)
    return None


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def xlsx_upload(request):
    plat_name = request.values.get("plat_name", 0)

    if 'xlsx_file' not in request.files:
        return redirect(request.url)

    xlsx_file = request.files['xlsx_file']

    if xlsx_file.filename == '':
        return redirect(request.url)

    if xlsx_file and allowed_file(xlsx_file.filename):
        file_upload = FileUpload(
                xlsx_file, plat_name)
        code, msg = file_upload.upload()
        return msg, plat_name


@app.route('/')
@auth.login_required
def index():
    return render_template('index.html')


@app.route('/plat/upload/', methods=['GET', 'POST'])
@auth.login_required
def xlsx_plat_to_mongo():
    if request.method == 'POST':
        msg, plat_name = xlsx_upload(request)
        return render_template('plat_upload.html', data=msg)
    return render_template("plat_upload.html")


@app.route('/summary/create/')
@auth.login_required
def xlsx_summary_create():
    return render_template("create_or_modify_summary.html")


@app.route('/summary/', methods=['GET', 'POST'])
@auth.login_required
def summary():
    sheets_name = request.values.get("sheets_name", 0)
    col_dict = request.values.get("col_dict", 0)
    plat_name = request.values.get("plat_name", 0)
    ncols = request.values.get("ncols", 0)
    ncols = int(ncols)
    col_dict = json.loads(col_dict)
    sheets_name = sheets_name.split(",")
    sheets_name = [i.replace(' ', '') for i in sheets_name]

    summary = SummaryHandle(plat_name, ncols, sheets_name, col_dict)
    if request.method == 'POST':
        summary = SummaryHandle(plat_name, ncols, sheets_name, col_dict)
        ret = summary.create_or_modify()
        return render_template("create_or_modify_summary.html", data=ret)
    return summary.summary_list()


if __name__ == '__main__':
    app.run(host="0.0.0.0", debug=True)
