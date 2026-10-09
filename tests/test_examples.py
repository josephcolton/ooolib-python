#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Runs every example script and checks the documents it creates."""

# Import python modules
import glob
import os
import subprocess
import sys
import unittest

# Import test helpers (also puts code/ on the path)
import helpers
import ooolib

class ExamplesTest(helpers.DocumentTestCase):
    def runExample(self, script, *arguments):
        environment = dict(os.environ)
        environment["PYTHONPATH"] = helpers.CODE_DIR
        result = subprocess.run([sys.executable, script] + list(arguments), cwd=self.tempDir,
                                env=environment, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, "%s failed:\n%s" % (script, result.stderr))
        return result.stdout

    def test_examples(self):
        scripts = sorted(glob.glob(os.path.join(helpers.EXAMPLES_DIR, "*-example*.py")))
        self.assertTrue(scripts)
        for script in scripts:
            if os.path.basename(script).startswith("read-"): continue
            with self.subTest(os.path.basename(script)):
                before = set(os.listdir(self.tempDir))
                self.runExample(script)
                created = set(os.listdir(self.tempDir)) - before
                self.assertTrue(created, "%s did not create a document" % script)
                for filename in created:
                    self.assertValidODF(self.path(filename))
                    ooolib.openDocument(self.path(filename))

    def test_read_example(self):
        self.runExample(os.path.join(helpers.EXAMPLES_DIR, "draw-example01.py"))
        output = self.runExample(os.path.join(helpers.EXAMPLES_DIR, "read-example01.py"))
        self.assertIn("Title:       Draw Example 1", output)
        self.assertIn("Hello from ooolib Draw!", output)


if __name__ == "__main__":
    unittest.main()
