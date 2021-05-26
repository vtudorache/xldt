from distutils.core import setup, Extension

def main():
    setup(name="xldt",
        version="0.1.0",
        description="Excel-compatible date and time functions.",
        author="Vlad Tudorache",
        author_email="tudorache.vlad@gmail.com",
        ext_modules=[Extension("xldt", ["src/xldt.c"])]
    )

if __name__ == "__main__":
    main()
