import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pylightxl", # Replace with your own username
    version="1.61",
    author="Viktor Kis",
    author_email="realpydpiper@gmail.com",
    license="MIT",
    description="A light weight excel read/writer for python27 and python3 with no dependencies",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/PydPiper/pylightxl",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=2.7',
)