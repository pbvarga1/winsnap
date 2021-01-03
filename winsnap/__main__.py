from pathlib import Path
import site

path = (Path(__file__) / ".." / ".." / ".." / "app_packages").resolve()
if path.exists():
    site.addsitedir(path)
    site.addsitedir(path / "comtypes")


if __name__ == "__main__":
    from winsnap.winsnap import main

    main()
