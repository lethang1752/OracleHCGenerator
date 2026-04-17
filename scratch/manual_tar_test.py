import tarfile
from pathlib import Path
import os

print("Starting manual test...")
path = Path("scratch/test_tar.tar.bz2")
if path.exists(): os.unlink(path)

with tarfile.open(path, "w:bz2") as tar:
    f = open("scratch/dummy.txt", "w")
    f.write("hello world")
    f.close()
    tar.add("scratch/dummy.txt", arcname="dummy.txt")

print("Archive created.")

try:
    with tarfile.open(path, "r:bz2") as tar:
        print(f"Open successful. Members: {tar.getnames()}")
        m = tar.getmember("dummy.txt")
        content = tar.extractfile(m).read().decode()
        print(f"Content: {content}")
except Exception as e:
    print(f"FAILED: {e}")

print("Done.")
