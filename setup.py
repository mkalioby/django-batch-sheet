#!/usr/bin/env python

from setuptools import find_packages, setup

setup(
    name='django_batch_sheet',
    version='0.90.0',
    description='Create Excel Sheet from Django Model and load them automatically',
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",

    author='Mohamed El-Kalioby, Mahmood Nasr',
    author_email = 'mkalioby@mkalioby.com',
    url = 'https://github.com/mkalioby/django-batch-sheet/',
    download_url='https://github.com/mkalioby/django-batch-sheet/',
    license='MIT',
    packages=['batch_sheet','batch_sheet.management','batch_sheet.management.commands'],
    install_requires=[
        'django >= 2.2',
        'xlsxwriter',
        'xlrd'
      ],
    python_requires=">=3.6",
    include_package_data=True,
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Environment :: Web Environment",
        "Framework :: Django",
        "Framework :: Django :: 2.2",
        "Framework :: Django :: 3.0",
        "Framework :: Django :: 3.1",
        "Framework :: Django :: 3.2",
        "Framework :: Django :: 4.0",
        "Framework :: Django :: 4.1",
        "Framework :: Django :: 4.2",
        "Framework :: Django :: 5.0",
        "Framework :: Django :: 5.1",
        "Framework :: Django :: 5.2",
        "Framework :: Django :: 6.0",
        "Intended Audience :: Developers",
        "Operating System :: OS Independent",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Programming Language :: Python :: 3.14",
        "Topic :: Software Development :: Libraries :: Python Modules",
]
)
