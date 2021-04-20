from distutils.core import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name='yandex_direct_stats',
    packages=['yandex_direct_stats'],
    version='0.1',
    license='MIT',
    description='Some tools for perform Yandex Direct reports',
    author='Danila Panteleev',
    author_email='pont131995@gmail.com',
    url='https://github.com/danila-panteleev/yandex_direct_stats',
    download_url='https://github.com/danila-panteleev/yandex_direct_stats/archive/refs/tags/0.1.tar.gz',
    keywords=['Yandex Direct', 'Яндекс Директ'],
    install_requires=[
        'tapi-yandex-direct',
        'pandas',
        'openpyxl',
        'gspread',
    ],
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3'
    ],
)
