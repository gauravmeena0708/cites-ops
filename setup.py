from setuptools import setup, find_packages

setup(
    name="cites-ops",
    version="1.0.0",
    description="Enterprise Operations Intelligence & Issue Triage Framework",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_type="text/markdown",
    packages=find_packages(),
    python_requires=">=3.9",
    install_requires=[
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
        "python-docx>=1.1.0",
        "python-pptx>=0.6.23",
        "jinja2>=3.1.0",
        "pyyaml>=6.0",
    ],
    entry_points={
        "console_scripts": [
            "cites-ops=cites_ops.cli:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
)
