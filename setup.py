from setuptools import setup, find_packages

setup(
    name="IT_Bench_Inventory",
    version="1.0",
    packages=find_packages(),
    install_requires=[
        "selenium>=4.15.2",
        "requests>=2.28.0",  # Add requests module
    ],
    python_requires='>=3.8',
    entry_points={
        'console_scripts': [
            'itbench=main:main',
        ],
    },
)
