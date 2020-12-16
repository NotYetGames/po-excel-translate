from setuptools import setup, find_packages
import sys
import os

root_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(root_dir, "VERSION")) as version_file:
    version = version_file.read().strip()

setup(
    name="po-excel-translate",
    version=version,
    description="Convert between Excel and PO files",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    classifiers=[
        "Environment :: Console",
        "Intended Audience :: Developers",
        "License :: DFSG approved",
        "License :: OSI Approved :: BSD License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
    ],
    keywords="translation po gettext Babel lingua excel portable object",
    author="Daniel Butum",
    author_email="daniel@notyet.eu",
    url="https://github.com/NotYetGames/po-excel-translate",
    license="BSD",
    # packages=find_packages(),
    py_modules=["po_excel_translate", "po2xls", "xls2po"],
    include_package_data=True,
    zip_safe=True,
    install_requires=["click", "polib", "openpyxl"],
    entry_points={"console_scripts": ["po2xls=po2xls:main", "xls2po=xls2po:main"]},
    python_requires=">=3.2",
)
