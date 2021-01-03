from invoke import task


@task
def lint(c):
    c.run("isort --check winsnap.py", warn=True)
    c.run("black --check winsnap.py", warn=True)
    c.run("flake8 winsnap.py", warn=True)


@task
def format(c):
    c.run("isort winsnap.py", warn=True)
    c.run("black winsnap.py", warn=True)
