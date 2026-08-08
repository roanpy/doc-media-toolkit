#!/usr/bin/env python3

from __future__ import annotations

import os
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ISOLATED_PREFIX = "test_verified_regressions.DesktopLifecycleTest."
sys.path[:0] = [str(ROOT), str(ROOT / "src"), str(ROOT / "tests")]


def test_ids(suite: unittest.TestSuite) -> list[str]:
    ids: list[str] = []
    for test in suite:
        if isinstance(test, unittest.TestSuite):
            ids.extend(test_ids(test))
        else:
            ids.append(test.id())
    return ids


def run_batch(names: list[str]) -> int:
    result = unittest.TextTestRunner(verbosity=2).run(
        unittest.defaultTestLoader.loadTestsFromNames(names)
    )
    exit_code = 0 if result.wasSuccessful() else 1
    if any(name.startswith(ISOLATED_PREFIX) for name in names):
        sys.stdout.flush()
        sys.stderr.flush()
        os._exit(exit_code)
    return exit_code


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--batch":
        return run_batch(sys.argv[2:])

    tests = test_ids(unittest.defaultTestLoader.discover(ROOT / "tests"))
    shared = [test for test in tests if not test.startswith(ISOLATED_PREFIX)]
    batches = [shared[index : index + 20] for index in range(0, len(shared), 20)]
    batches.extend([test] for test in tests if test.startswith(ISOLATED_PREFIX))

    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join(
        filter(
            None,
            (
                str(ROOT / "src"),
                str(ROOT / "tests"),
                env.get("PYTHONPATH"),
            ),
        )
    )
    for batch in batches:
        result = subprocess.run(
            [sys.executable, __file__, "--batch", *batch],
            cwd=ROOT,
            env=env,
            check=False,
            creationflags=(
                subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
            ),
        )
        if result.returncode:
            return result.returncode
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
