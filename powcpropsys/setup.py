from setuptools import find_packages, setup

setup(
    name="powcpropsys",
    version="0.0.1",
    install_requires=("powc"),
    packages=find_packages(where="src"),
    package_dir={"": "src"},
)
