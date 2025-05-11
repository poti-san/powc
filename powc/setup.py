from setuptools import find_packages, setup

setup(
    name="powc",
    version="0.0.1",
    install_requires=("comtypes"),
    packages=find_packages(where="src"),
    package_dir={"": "src"},
)
