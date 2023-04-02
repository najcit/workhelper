import os
from flask import Flask, send_file, request

from appmyapp import AppMyApp

app = Flask(__name__)


@app.route('/')
@app.route('/latest')
def latest_version():
    release = os.path.join('.', 'release')
    versions = sorted(os.listdir(release), reverse=True)
    return versions[0] if len(versions) > 0 else ''


@app.route('/release')
@app.route('/query/versions')
def query_versions():
    release = os.path.join('.', 'release')
    versions = sorted(os.listdir(release), reverse=True)
    return versions


@app.route('/share/app/', methods=['POST'])
def share_app():
    print('name', request.form.get('name'))
    print('icon', request.form.get('icon'))
    print('path', request.form.get('path'))
    print('time', request.form.get('time'))
    print('type', request.form.get('type'))
    print('parent', request.form.get('parent'))
    file = request.files.get('app')
    parent = request.form.get('parent')
    root_path = os.path.join('.', 'share')
    parent_path = os.path.join(root_path, parent)
    print(parent_path)
    if not os.path.exists(parent_path):
        os.makedirs(parent_path)
    if file:
        file.save(os.path.join(parent_path, file.filename))
    return 'ok'


@app.route('/query/apps', methods=['GET'])
def query_apps():
    root_path = os.path.join('.', 'share')
    my_app = AppMyApp()
    apps_info = my_app.get_apps(root_path)
    return apps_info


@app.route('/download', methods=['GET'])
@app.route('/download/<string:version>', methods=['GET'])
def download_version(version=None):
    release = os.path.join('.', 'release')
    if version is None:
        versions = sorted(os.listdir(release), reverse=True)
        version = versions[0] if len(versions) > 0 else ''
    file_name = 'workhelper.exe'
    file_path = os.path.join(release, version, file_name)
    return send_file(file_path, download_name=file_name)


class AppService:
    def __init__(self):
        self.app = app

    def run(self):
        self.app.run()


if __name__ == '__main__':
    AppService().run()
