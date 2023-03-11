import winreg


class AppLocalApp(object):

    def __init__(self):
        self.apps = []

    def get_local_installed_app_list(self):
        self.apps = []
        key_name_list = [r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                         r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall']
        for key_name in key_name_list:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_name, 0, winreg.KEY_READ)
            key_info = winreg.QueryInfoKey(key)[0]
            for i in range(key_info):
                app = {}
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    app['name'] = winreg.QueryValueEx(subkey, "DisplayName")[0]
                    app['path'] = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                    app['key'] = app['path']
                    app['type'] = 'local'
                    if not app['path']:
                        continue
                except EnvironmentError:
                    continue
                try:
                    app['icon'] = winreg.QueryValueEx(subkey, "DisplayIcon")[0].split(',')[0].replace('"', '')
                    if not app['icon']:
                        continue
                    if app['icon'].find('.exe') != -1:
                        app['path'] = app['icon']
                except EnvironmentError:
                    app['icon'] = ''
                try:
                    app['time'] = winreg.QueryValueEx(subkey, "InstallDate")[0]
                except EnvironmentError:
                    app['time'] = ''
                try:
                    app['version'] = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                except EnvironmentError:
                    app['version'] = ''
                try:
                    app['publisher'] = winreg.QueryValueEx(subkey, "Publisher")[0]
                except EnvironmentError:
                    app['publisher'] = ''
                self.apps.append(app)
        return self.apps

    def get_local_installed_browsers(self):
        if not self.apps:
            self.get_local_installed_app_list()
        browsers = set()
        for app in self.apps:
            for name in ['firefox', 'chrome', 'edge']:
                if app['name'].lower().find(name) != -1:
                    browsers.add(name)
        return sorted(browsers)
