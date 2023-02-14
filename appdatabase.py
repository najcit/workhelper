import time
import unqlite


class AppDatabase:
    def __init__(self, db_name):
        self.db_name = db_name
        self.db = unqlite.UnQLite(self.db_name)
        apps = self.db.collection('apps')
        if not apps.exists():
            apps.create()

    def __del__(self):
        self.close()

    def add_app(self, app_name):
        apps = self.db.collection('apps')
        app = apps.filter(lambda obj: obj['name'] == app_name)
        if not app:
            app = [{'name': app_name, 'add_time': time.time()}]
            apps.store(app)
        self.db.commit()

    def start_app(self, app_name):
        apps = self.db.collection('apps')
        app = apps.filter(lambda obj: obj['name'] == app_name)
        if not app:
            app = [{'name': app_name, 'count': 1, 'start_time': time.time()}]
            apps.store(app)
        else:
            if 'count' not in app[0]:
                app[0]['count'] = 1
            else:
                app[0]['count'] += 1
            apps.update(0, app[0])
        self.db.commit()
        print(app)

    def stop_app(self, app_name):
        apps = self.db.collection('apps')
        app = apps.filter(lambda obj: obj['name'] == app_name)
        if app:
            app[0]['stop_time'] = time.time()
            apps.update(0, app[0])
        self.db.commit()
        print(app)

    def close(self):
        self.db.close()
