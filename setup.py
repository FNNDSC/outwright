from setuptools import setup, find_packages
from pathlib import Path


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


setup(
    name="outlook_email_sender",
    version="0.1.0",
    description="A Python package for authenticating to Outlook Web and sending emails using Playwright.",
    author="Rudolph Pienaar",
    author_email="rudolph.pienaar@childrens.harvard.edu",
    packages=find_packages(where="src"),  # Point to src directory for modules
    package_dir={"": "src"},  # Tell setuptools where to find the code
    install_requires=parse_requirements("requirements.txt"),
    entry_points={
        "console_scripts": [
            "outwright=outlook_email_sender:main",  # Bind the main function to 'outwright' command
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.11",
)
