from invoke import task
import shutil
from pathlib import Path

@task
def lint(c):
    c.run("isort --check winsnap", warn=True)
    c.run("black --check winsnap", warn=True)
    c.run("flake8 winsnap", warn=True)


@task
def format(c):
    c.run("isort winsnap", warn=True)
    c.run("black winsnap", warn=True)


@task
def bootstrap(c):
    import comtypes.client
    comtypes.client.GetModule(r"C:\\Windows\\System32\\UIAutomationCore.dll")

@task(bootstrap)
def package(c):
    windows = (Path(__file__) / ".." / "windows").resolve()
    if windows.exists():
        shutil.rmtree(windows)
    c.run("briefcase create")
    c.run("briefcase build")
    # The first time it runs it fails but the next time it works. This is to bootstrap the uia
    # automation files for comtypes
    c.run("briefcase run", echo=False, warn=True)
    c.run("briefcase package")


@task(package)
def install(c):
    windows = (Path(__file__) / ".." / "windows").resolve()
    msi = list(windows.glob("*.msi"))[0]
    c.run(str(msi))


@task(bootstrap)
def exe(c):
    c.run("pyinstaller winsnap.spec")
