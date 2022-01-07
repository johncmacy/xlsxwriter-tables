from setuptools import setup, find_packages

with open("README.md", "r") as readme_file:
    readme = readme_file.read()

requirements = []

setup(
    name="xlsxwriter-tables",
    version="0.0.1",
    author="John Macy",
    author_email="johncmacy@gmail.com",
    description="Easily export nested data to Excel",
    long_description=readme,
    long_description_content_type="text/markdown",
    url="https://github.com/johncmacy/xlsxwriter-tables/",
    packages=find_packages(),
    install_requires=requirements,
    classifiers=[
        "Programming Language :: Python :: 3.9",
        "License :: OSI Approved :: MIT License",
    ],
)