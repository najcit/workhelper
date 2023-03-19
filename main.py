# -*- coding: utf-8 -*-
"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""
import click
import os.path
import shutil
import time
from app import MyApp


@click.command()
@click.option('--root', default='', type=str, help='the root directory of the application')
@click.option('--theme', default='', type=str, help='the theme of the application')
@click.option('--enable_systray', is_flag=True, default=False, help='enable system tray for the application')
@click.option('--enable_env', is_flag=True, default=True, help='enable environment variable for the application')
@click.option('--auto_update', is_flag=True, default=False, help='enable auto update for the application')
@click.option('--update_path1', default='', type=str, help='copy self to that path and run self')
@click.option('--update_path2', default='', type=str, help='copy self to that path and run self')
def main(root, theme, enable_systray, enable_env, auto_update, update_path1, update_path2):
    if update_path1 and update_path2:
        print(update_path1, update_path2)
        time.sleep(5)
        shutil.copy(update_path1, update_path2)
    else:
        MyApp.show(root=root, theme=theme, enable_systray=enable_systray,
                   enable_env=enable_env, auto_update=auto_update)


if __name__ == '__main__':
    main()
