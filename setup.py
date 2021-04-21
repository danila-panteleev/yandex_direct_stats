from distutils.core import setup

setup(
    name='yandex_direct_stats',
    packages=['yandex_direct_stats'],
    version='0.3',
    license='MIT',
    description='Some tools for perform Yandex Direct reports',
    author='Danila Panteleev',
    author_email='pont131995@gmail.com',
    url='https://github.com/danila-panteleev/yandex_direct_stats',
    download_url='https://github.com/danila-panteleev/yandex_direct_stats/archive/refs/tags/0.3.tar.gz',
    keywords=['Yandex Direct', 'Яндекс Директ'],
    install_requires=[
        'tapi-yandex-direct',
        'pandas',
        'openpyxl',
        'gspread',
        'simplejson',

    ],
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3'
    ],
)
