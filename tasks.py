"""
Defines useful development functions, available using the invoke command
Tasks can be run with "invoke *task name" eg "invoke lint"
"""
from pathlib import Path
from invoke import task

# from Main import optimize as _optimize


def _find_packages(path: Path):
    for pkg in path.iterdir():
        if pkg.is_dir() and len(list(pkg.glob("**/*.py"))) >= 1:
            yield pkg


def _find_scripts(path: Path):
    return path.glob("**/*.py")


@task
def lint(c):
    """Uses flake8 to report linting errors in your code"""
    c.run("flake8 .", echo=True, pty=True)


@task
def format(c, fix=False, diff=False):
    """Uses black to report any formatting issues in your code

    Args:
        fix: Flag to automatically fix formatting issues in your code
        diff: Flag to include a diff between your current code and the recommended code
    """
    if fix and diff:
        print("Please use only --diff OR --fix, but not both, when calling format.")
    else:
        if fix:
            arg = ""
        elif diff:
            arg = "--diff"
        else:
            arg = "--check"

        c.run(f"black {arg} --line-length=99 --skip-magic-trailing-comma .", echo=True, pty=True)


# @task
# def optimize(c, save_optimziation=True):
#     """
#
#     Determines optimal core count for parallel executions
#     :param save_optimziation: saves optimizations rather than just printing
#     """
#     _optimize(save_optimziation)
