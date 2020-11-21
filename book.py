from os.path import dirname, exists, join as join_path, realpath
from sys import argv, path as sys_path


parent_dir = dirname(realpath(__file__))
config_path = join_path(parent_dir, 'config')
if exists(config_path):
    sys_path.append(config_path)
    from checkmyxl import main, start
    if __name__ == '__main__':
        start(argv)

