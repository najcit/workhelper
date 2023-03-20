# -*- coding: utf-8 -*-
"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""
import click
import os
from app import MyApp


@click.command()
@click.option('--root', default='', type=str, help='the root directory of the application')
@click.option('--theme', default='', type=str, help='the theme of the application')
@click.option('--enable_systray', is_flag=True, default=False, help='enable system tray for the application')
@click.option('--enable_env', is_flag=True, default=True, help='enable environment variable for the application')
@click.option('--app_path', default='', type=str, help='the newest application file')
@click.option('--install_path', default='', type=str, help='the install path of the application')
@click.option('--timeout', default=3, type=int, help='the timeout to install the application')
def main(root, theme, enable_systray, enable_env, app_path, install_path, timeout):
    if os.path.isfile(app_path) and os.path.isdir(install_path):
        MyApp.upgrade(app_path, install_path, timeout)
    else:
        MyApp.show(root=root, theme=theme, enable_systray=enable_systray, enable_env=enable_env)


if __name__ == '__main__':
    main()
