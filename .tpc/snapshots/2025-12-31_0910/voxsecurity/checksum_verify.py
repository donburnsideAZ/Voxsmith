import hashlib, os, sys

def _read_checksums(path):
    checks = {}
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split()
                if len(parts) >= 2:
                    h, name = parts[0], parts[-1]
                    checks[name] = h.lower()
    except FileNotFoundError:
        pass
    return checks

def verify_self(checksums_path=None, target_path=None):
    # For dev runs use the running interpreter; for exe runs use the exe
    target_path = target_path or (sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
    base = os.path.dirname(target_path)
    checksums_path = checksums_path or os.path.join(base, "checksums.sha256")

    expected_map = _read_checksums(checksums_path)
    target_name = os.path.basename(target_path)
    expected = expected_map.get(target_name)
    if not expected:
        return False, "Unverified build"

    hasher = hashlib.sha256()
    try:
        with open(target_path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                hasher.update(chunk)
    except OSError:
        return False, "Unverified build"

    ok = hasher.hexdigest().lower() == expected
    return ok, ("Verified build" if ok else "Checksum mismatch")
