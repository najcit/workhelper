# -*- coding: utf-8 -*-
"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""
import click
from app import MyApp


@click.command()
@click.option('--root', default='root', type=str, help='the root directory of the application')
@click.option('--theme', default='DarkGreen7', type=str, help='the theme of the application')
@click.option('--enable_systray', is_flag=True, default=False, help='enable system tray for the application')
@click.option('--enable_env', is_flag=True, default=True, help='enable environment variable for the application')
@click.option('--auto_update', is_flag=True, default=False, help='enable auto update for the application')
def main(root, theme, enable_systray, enable_env, auto_update):
    MyApp.show(root, theme, enable_systray, enable_env, auto_update)


if __name__ == '__main__':
    main()
