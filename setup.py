from setuptools import setup, find_packages

setup(

    name='robotframework-openpyxllib',
    version='0.2',
    description='Robotframework library for excel xlsx file format',
    author='Vallinayagam.K',
    author_email='valli.python@gmail.com',
    packages=find_packages(),
    install_requires=['openpyxl']
    
)