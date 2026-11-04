# ChronoFlow for Apple Silicon

This adapts the existing PySide6 application to macOS ARM64. It has not been run on an actual Mac; the GitHub Actions build is the first macOS test.

Upload app.py, requirements-macos.txt, and .github/workflows/build-macos.yml into the **root** of your existing TimeTracker repository. In GitHub Actions, select Build ChronoFlow for Apple Silicon, Run workflow. Download artifact ChronoFlow-AppleSilicon; unzip the artifact, then unzip ChronoFlow-AppleSilicon.zip to obtain ChronoFlow.app. Move to Applications.

The unsigned app may be blocked by macOS Gatekeeper. Review the source before approving it in System Settings > Privacy & Security. App and window activity detection can require Screen Recording permissions; some window titles are unavailable. The global hotkey currently only works on Windows; in-window shortcut remains.

Local database: ~/Library/Application Support/ChronoFlow/chronoflow.db. For migration from Windows, use the app's backup/restore feature. No Python installation on your Mac is required.
