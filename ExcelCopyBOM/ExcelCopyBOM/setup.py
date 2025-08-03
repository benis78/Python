from setuptools import setup, find_packages

setup(
    name="ExcelCopyBOM",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        'pandas',
        'openpyxl',
        'Pillow',
        'pywin32'
    ],
    python_requires='>=3.9'
) 