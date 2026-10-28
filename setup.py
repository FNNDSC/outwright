from pathlib import Path

from setuptools import find_packages, setup


# Read the contents of requirements.txt
def parse_requirements(file_path: str) -> list[str]:
    """
    Reads the requirements file and returns a list of dependencies.

    Args:
        file_path (str): Path to the requirements.txt file.

    Returns:
        list[str]: List of dependencies.
    """
    return Path(file_path).read_text().splitlines()


def readme():
    readme_path = Path("README.adoc")
    if readme_path.exists():
        return readme_path.read_text()
    return ""


setup(
    name="outlook_email_sender",
    version="2.0.0",
    description="A Python package for authenticating to Outlook Web and sending emails using Playwright.",
    long_description=readme(),
    author="Rudolph Pienaar",
    author_email="rudolph.pienaar@childrens.harvard.edu",
    packages=find_packages(where="src"),  # Point to src directory for modules
    package_dir={"": "src"},  # Tell setuptools where to find the code
    install_requires=parse_requirements("requirements.txt"),
    entry_points={
        "console_scripts": [
            "outwright=outlook_email_sender:outwright_main",  # Bind the entry point to 'outwright' command
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.11",
)
