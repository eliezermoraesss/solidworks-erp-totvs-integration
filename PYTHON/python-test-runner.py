import subprocess

subprocess.Popen(["python", "python-test.py"], creationflags=subprocess.DETACHED_PROCESS, close_fds=True)
