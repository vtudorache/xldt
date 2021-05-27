import setuptools

xldt_description = None

with open("README.md", "r", encoding="utf-8") as f:
    xldt_description = f.read()

KEYWORDS = ["python", "excel", "date", "time"]

setuptools.setup(name="xldt",
        version="0.1.1",
        author="Vlad Tudorache",
        author_email="tudorache.vlad@gmail.com",
        description="Excel-compatible date and time functions.",
        long_description=xldt_description,
        long_description_content_type="text/markdown",
        url="https://github.com/vtudorache/xldt",
        classifiers=[
            "Programming Language :: Python :: 3",
            "License :: OSI Approved :: MIT License",
            "Operating System :: OS Independent",
        ],
        keywords = " ".join(KEYWORDS),
        package_dir={"": "src"},
        packages=setuptools.find_packages(where="src"),
        ext_modules=[setuptools.Extension("xldt._xldt", ["src/xldt.c"])],
        python_requires=">=3.6"
    )
