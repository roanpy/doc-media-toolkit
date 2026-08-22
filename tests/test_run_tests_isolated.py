from __future__ import annotations

import subprocess
import sys
import unittest
from unittest.mock import patch

from scripts import run_tests_isolated


class IsolatedTestRunnerTest(unittest.TestCase):
    def test_gui_process_exit_is_retried_once(self) -> None:
        test_name = (
            "test_verified_regressions.DesktopLifecycleTest."
            "test_workspace_controls_stay_in_bounds_across_size_matrix"
        )
        results = (
            subprocess.CompletedProcess([], 1),
            subprocess.CompletedProcess([], 0),
        )
        with (
            patch.object(sys, "argv", ["run_tests_isolated.py"]),
            patch.object(run_tests_isolated, "test_ids", return_value=[test_name]),
            patch.object(
                run_tests_isolated.subprocess, "run", side_effect=results
            ) as run,
        ):
            self.assertEqual(run_tests_isolated.main(), 0)
        self.assertEqual(run.call_count, 2)


if __name__ == "__main__":
    unittest.main()
