from setuptools import setup

setup(
    name="doc2pdf",
    version="0.0.1",
    url="https://github.com/andiwand/doc2pdf",
    license="GNU Lesser General Public License",
    author="Andreas Stefl",
    install_requires=["watchdog", "pywin32"],
    author_email="stefl.andreas@gmail.com",
    description="full automatic ms office to pdf converter",
    long_description="",
    package_dir={"": "src"},
    packages=["doc2pdf"],
    platforms="windows",
    entry_points={
        "console_scripts": ["doc2pdf = doc2pdf.cli:main"],
        "gui_scripts": ["doc2pdfw = doc2pdf.cli:main"]
    },
)
