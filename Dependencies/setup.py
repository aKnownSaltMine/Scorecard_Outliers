import sys
import subprocess
import os

def setup(*args):

    if len(args) == 0:
        # grabbing the requirements text next to the script
        requirements_txt = os.path.join(os.path.dirname(__file__), 'requirements.txt')

        # open and read packages that are needed.
        with open(requirements_txt, 'r', encoding='UTF-8') as file:
            data = file.read()
            modules = data.split('\n')
        print('-'*20)
        print('These Packages are needed:')
        print(*modules, sep='\n')
    else:
        modules = args

    # list all packages installed through pip in freeze format
    stream = os.popen('pip list --user --format=freeze')
    Package = stream.read()
    pip_list = list(Package.split("\n"))

    # loop through required modules, if they are in the pip_list skip, if not, install
    for package in modules:
        if package not in pip_list:
            package_name = package.split("==")[0]
            package_version = package.split("==")[-1]
            if package_name != package_version:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', f'{package_name}[full]=={package_version}', '--user'])
            else:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', f'{package_name}[full]', '--user'])
            print(f'Installed: {package}')
        else:
            print(f'{package} is already installed')
if __name__ == '__main__':
    setup()
        