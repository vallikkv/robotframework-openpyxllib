from setuptools import setup, find_packages

with open("README.txt", "r") as fh:
	long_description = fh.read()

setup(

    name='robotframework-openpyxllib',
    version='0.7',
    description='Robotframework library for excel xlsx file format',
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/vallikkv/robotframework-openpyxllib',
    author='Vallinayagam.K',
    author_email='valli.python@gmail.com',
    packages=find_packages(),
    install_requires=['openpyxl']
    
)
