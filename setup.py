from setuptools import setup, find_packages

setup(
    name="mail-cleaner-pro",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "pywin32>=308",
        "PyQt6>=6.8.1",
        "python-dateutil>=2.8.2",
        "psutil>=5.9.8",
    ],
    entry_points={
        "console_scripts": [
            "mail-cleaner=main:main",
        ],
    },
    author="Your Name",
    author_email="your.email@example.com",
    description="A high-performance Outlook email management tool",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    keywords="outlook, email, management, cleanup",
    url="https://github.com/yourusername/mail-cleaner-pro",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Topic :: Office/Business :: Email",
    ],
    python_requires=">=3.9",
) 