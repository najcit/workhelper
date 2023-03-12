import winreg
import ctypes


class Registry(object):

    def __init__(self, key_location, key_path):
        self.reg_key = winreg.OpenKey(key_location, key_path, 0, winreg.KEY_ALL_ACCESS)

    def set_key(self, name, value):
        try:
            _, reg_type = winreg.QueryValueEx(self.reg_key, name)
        except WindowsError:
            # If the value does not exist yet, we (guess) use a string as the reg_type
            reg_type = winreg.REG_SZ
        winreg.SetValueEx(self.reg_key, name, 0, reg_type, value)

    def delete_key(self, name):
        try:
            winreg.DeleteValue(self.reg_key, name)
        except WindowsError:
            # Ignores if the key value doesn't exist
            pass


class EnvironmentVariables(Registry):
    """
    Configures the HTTP_PROXY environment variable
    """

    def __init__(self):
        super(EnvironmentVariables, self).__init__(
            winreg.HKEY_CURRENT_USER,
            'Environment')

    def on(self, key, value):
        self.set_key(key, value)
        self.refresh()

    def off(self, key):
        self.delete_key(key)
        self.refresh()

    @staticmethod
    def refresh(self):
        hwnd_broadcast = 0xFFFF
        wm_settingchange = 0x1A
        smto_abortifhung = 0x0002
        result = ctypes.c_long()
        ctypes.windll.user32.SendMessageTimeoutW(hwnd_broadcast, wm_settingchange, 0, u'Environment',
                                                 smto_abortifhung, 5000, ctypes.byref(result))


if __name__ == '__main__':
    ev = EnvironmentVariables()
    ev.on('My', 'Test')
    ev.off('My')
