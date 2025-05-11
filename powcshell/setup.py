from setuptools import find_packages, setup

setup(
    name="powcshell",
    version="0.0.1",
    install_requires=("powcpropsys"),
    packages=find_packages(where="src"),
    package_dir={"": "src"},
)
