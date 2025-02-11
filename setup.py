"""Setup configuration for the Jira Sprint Analysis package."""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="jira-sprint-analysis",
    version="1.0.0",
    author="Shahil Bhimji",
    description="A tool for analyzing Jira sprint estimates and actuals",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/shahil-bhimji/halfords-jira-sprint-analysis",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pandas>=1.3.0",
        "streamlit>=1.0.0",
        "matplotlib>=3.4.0",
        "openpyxl>=3.0.0",
        "numpy>=1.20.0",
    ],
    entry_points={
        "console_scripts": [
            "jira-analysis=app:main",
        ],
    },
)
