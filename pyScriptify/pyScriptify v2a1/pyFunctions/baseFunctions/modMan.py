import os

def modMan(library):
    for module in library:
        try:
            print(f"Attemping to import {module}.")
            __import__(module)
            print(f"{module} import success.")

        except ImportError:
            try:
                print(f"{module} missing, attempting to install...")
                os.system(f"pip install {module}")
                print(f"{module} installed.")

            except Exception as e:
                print(f"Error trying to install {module}. {e}")