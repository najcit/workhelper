import zipfile
import click


@click.command()
@click.option('--app', help='The executable app ')
@click.option('--release', help='The release version')
def main(app, release):
    zip_file = f'{release}.zip'.format(release=release)
    z = zipfile.ZipFile(zip_file, 'w')
    z.write(app)
    z.close()


if __name__ == '__main__':
    main()
