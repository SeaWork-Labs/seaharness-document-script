import hashlib
import os
import tempfile


def sha256_file(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_replace(dest: str, data: bytes, expected_sha: str | None = None) -> None:
    if expected_sha is not None and os.path.exists(dest):
        current = sha256_file(dest)
        if current != expected_sha:
            raise SystemExit(f"sha256 conflict: {current} != {expected_sha}")
    directory = os.path.dirname(dest) or "."
    os.makedirs(directory, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=f".{os.path.basename(dest)}.", dir=directory)
    try:
        os.write(fd, data)
        os.close(fd)
        fd = -1
        os.replace(tmp, dest)
    except Exception:
        if fd >= 0:
            os.close(fd)
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise
