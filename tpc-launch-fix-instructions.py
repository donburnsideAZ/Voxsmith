# TPC Launch Fix - Smart Python Detection
# ==========================================
# 
# Problem: Launch uses hardcoded "python3" which:
#   1. Doesn't exist on Windows (uses "python")
#   2. Doesn't use project venvs where dependencies are installed
#
# Solution: Detect and use the right Python interpreter

# STEP 1: Add this import at the top of ui/workspace.py
# -----------------------------------------------------
# Find the existing imports and add:

import platform

# (If platform isn't already imported)


# STEP 2: Add this method to the WorkspaceWidget class
# ----------------------------------------------------
# Add it near the other helper methods (like _get_project_key)

def get_python_for_project(self) -> str:
    """
    Get the best Python interpreter for launching this project.
    
    Priority:
    1. TPC's build venv for this project (if Pack was used)
    2. Project-local venv (venv/ or .venv/)
    3. System Python (platform-appropriate)
    """
    if not self.project:
        return "python" if platform.system() == "Windows" else "python3"
    
    # Check for TPC's build venv
    tpc_venv = Path.home() / ".tpc" / "venvs" / self.project.name
    if tpc_venv.exists():
        if platform.system() == "Windows":
            python_path = tpc_venv / "Scripts" / "python.exe"
        else:
            python_path = tpc_venv / "bin" / "python"
        
        if python_path.exists():
            return str(python_path)
    
    # Check for project-local venv
    for venv_name in ["venv", ".venv", "env"]:
        local_venv = self.project.path / venv_name
        if local_venv.exists():
            if platform.system() == "Windows":
                python_path = local_venv / "Scripts" / "python.exe"
            else:
                python_path = local_venv / "bin" / "python"
            
            if python_path.exists():
                return str(python_path)
    
    # Fall back to system Python
    return "python" if platform.system() == "Windows" else "python3"


# STEP 3: Update the on_launch method
# -----------------------------------
# Find on_launch() and replace it with this:

def on_launch(self):
    """Launch the project's main file."""
    if not self.project:
        return
    
    project_key = self._get_project_key()
    if not project_key:
        return
    
    # Clear previous output for this project
    self.output_panel.clear()
    
    # Get the right Python for this project
    python_cmd = self.get_python_for_project()
    
    # Show which Python we're using (helpful for debugging)
    if python_cmd not in ("python", "python3"):
        self.append_output(f"▶ Running {self.project.main_file} (using project venv)...\n\n")
    else:
        self.append_output(f"▶ Running {self.project.main_file}...\n\n")
    
    # Disable launch button while running
    self.btn_launch.setEnabled(False)
    self.btn_launch.setText("⏳ Running...")
    
    # Use QProcess for better Qt integration
    process = QProcess(self)
    process.setWorkingDirectory(str(self.project.path))
    
    # Store which project this process belongs to
    process.setProperty("project_key", project_key)
    
    # Connect signals
    process.readyReadStandardOutput.connect(lambda: self.on_stdout_ready(process))
    process.readyReadStandardError.connect(lambda: self.on_stderr_ready(process))
    process.finished.connect(lambda exit_code, exit_status: self.on_process_finished(process, exit_code, exit_status))
    
    # Track this process for this project
    self.running_processes[project_key] = process
    
    # Start the process with the right Python
    process.start(python_cmd, [self.project.main_file])


# WHAT THIS FIXES:
# ================
# 
# 1. Windows compatibility - uses "python" instead of "python3"
# 
# 2. Uses TPC's build venv if you've run Pack on this project
#    Location: ~/.tpc/venvs/ProjectName/
#    
# 3. Uses project-local venv if one exists (venv/, .venv/, env/)
#    This catches projects with their own environment
#
# 4. Falls back gracefully to system Python
#
# Now when you click Launch:
# - If Pack created a venv with PyQt6 → uses that Python → works!
# - If project has local venv → uses that → works!
# - Otherwise → system Python (may fail if deps not installed)
