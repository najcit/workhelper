import os
from flask import Flask, send_file

app = Flask(__name__)


@app.route('/')
@app.route('/latest')
def latest_version():
    release = os.path.join('.', 'release')
    versions = sorted(os.listdir(release), reverse=True)
    return versions[0] if len(versions) > 0 else ''


@app.route('/release')
def release_versions():
    release = os.path.join('.', 'release')
    versions = sorted(os.listdir(release), reverse=True)
    return versions


@app.route('/download/', methods=['GET'])
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
    app_service = AppService()
    app_service.run()
