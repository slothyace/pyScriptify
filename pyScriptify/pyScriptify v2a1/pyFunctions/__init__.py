import os
import importlib

packageDir = os.path.dirname(__file__)

for moduleName in os.listdir(packageDir):
    if moduleName.endswith('.py') and moduleName != "__init__.py":
        moduleName = moduleName.strip(".py")
        importlib.import_module(f".{moduleName}", package=__name__)