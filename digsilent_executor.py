"""
Module untuk eksekusi skrip di DIgSILENT PowerFactory
"""

import sys
import os
import subprocess
import time


class DIgSILENTExecutor:
    """
    Class untuk eksekusi skrip Python di DIgSILENT PowerFactory
    """

    def __init__(self):
        """Initialize executor"""
        self.pf_paths = self._find_powerfactory_paths()

    def _find_powerfactory_paths(self):
        """
        Cari path instalasi PowerFactory (support 2021 dan 2022)
        """
        possible_paths = [
            # PowerFactory 2021
            r"C:\Program Files\DIgSILENT\PowerFactory 2021\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP3\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP3\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP3\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP4\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP4\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP4\Python\3.10",
            # PowerFactory 2022
            r"C:\Program Files\DIgSILENT\PowerFactory 2022\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022\Python\3.11",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.11",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP2\Python\3.8",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP2\Python\3.9",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP2\Python\3.10",
            r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP2\Python\3.11",
        ]

        found_paths = []
        for path in possible_paths:
            if os.path.exists(path):
                found_paths.append(path)

        return found_paths

    def execute_script_direct(self, script_path):
        """
        Eksekusi skrip langsung dengan menambahkan PowerFactory path ke sys.path
        Metode ini untuk skrip yang dipanggil dari Python biasa

        Args:
            script_path: Path ke skrip yang akan dijalankan

        Returns:
            True jika sukses, False jika gagal
        """
        if not os.path.exists(script_path):
            print(f"✗ Script not found: {script_path}")
            return False

        print(f"Executing script: {script_path}")
        print("="*60)

        # Add PowerFactory path
        if self.pf_paths:
            pf_path = self.pf_paths[0]
            if pf_path not in sys.path:
                sys.path.append(pf_path)
            print(f"Added PowerFactory path: {pf_path}")
        else:
            print("Warning: PowerFactory path not found")

        # Execute script
        try:
            with open(script_path, 'r') as f:
                script_code = f.read()

            exec(script_code, {'__name__': '__main__'})
            return True

        except Exception as e:
            print(f"✗ Error executing script: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def execute_script_subprocess(self, script_path, python_executable=None):
        """
        Eksekusi skrip menggunakan subprocess
        Berguna jika ingin menjalankan di Python interpreter terpisah

        Args:
            script_path: Path ke skrip yang akan dijalankan
            python_executable: Path ke Python executable (optional)

        Returns:
            True jika sukses, False jika gagal
        """
        if not os.path.exists(script_path):
            print(f"✗ Script not found: {script_path}")
            return False

        if python_executable is None:
            python_executable = sys.executable

        print(f"Executing script: {script_path}")
        print(f"Using Python: {python_executable}")
        print("="*60)

        # Set environment untuk menambahkan PowerFactory path
        env = os.environ.copy()
        if self.pf_paths:
            pythonpath = env.get('PYTHONPATH', '')
            if pythonpath:
                pythonpath = f"{self.pf_paths[0]};{pythonpath}"
            else:
                pythonpath = self.pf_paths[0]
            env['PYTHONPATH'] = pythonpath

        # Execute
        try:
            result = subprocess.run(
                [python_executable, script_path],
                env=env,
                capture_output=True,
                text=True
            )

            # Print output
            if result.stdout:
                print(result.stdout)

            if result.stderr:
                print("STDERR:")
                print(result.stderr)

            if result.returncode == 0:
                print("="*60)
                print("✓ Script executed successfully")
                return True
            else:
                print("="*60)
                print(f"✗ Script failed with return code: {result.returncode}")
                return False

        except Exception as e:
            print(f"✗ Error executing script: {str(e)}")
            return False

    def execute_in_powerfactory(self, script_path):
        """
        Eksekusi skrip langsung di DIgSILENT PowerFactory
        Menggunakan PowerFactory Python API

        Args:
            script_path: Path ke skrip yang akan dijalankan

        Returns:
            True jika sukses, False jika gagal
        """
        if not os.path.exists(script_path):
            print(f"✗ Script not found: {script_path}")
            return False

        # Add PowerFactory path
        if self.pf_paths:
            pf_path = self.pf_paths[0]
            if pf_path not in sys.path:
                sys.path.append(pf_path)
        else:
            print("✗ PowerFactory path not found")
            return False

        try:
            import powerfactory

            # Get PowerFactory application
            app = powerfactory.GetApplication()
            if app is None:
                print("✗ Cannot connect to PowerFactory. Make sure PowerFactory is running.")
                return False

            print("✓ Connected to PowerFactory")

            # Get ComPython object untuk eksekusi skrip
            # Method 1: Langsung execute script
            print(f"Executing script: {script_path}")
            print("="*60)

            with open(script_path, 'r') as f:
                script_code = f.read()

            # Execute dalam context PowerFactory
            exec(script_code, {
                '__name__': '__main__',
                'powerfactory': powerfactory,
                'pf': powerfactory
            })

            print("="*60)
            print("✓ Script executed in PowerFactory context")
            return True

        except ImportError:
            print("✗ Cannot import powerfactory module")
            return False
        except Exception as e:
            print(f"✗ Error executing script: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def execute_and_wait(self, script_path, method='direct', wait_time=2):
        """
        Eksekusi skrip dan tunggu selesai

        Args:
            script_path: Path ke skrip yang akan dijalankan
            method: Metode eksekusi ('direct', 'subprocess', 'powerfactory')
            wait_time: Waktu tunggu setelah eksekusi (detik)

        Returns:
            True jika sukses, False jika gagal
        """
        start_time = time.time()

        if method == 'direct':
            success = self.execute_script_direct(script_path)
        elif method == 'subprocess':
            success = self.execute_script_subprocess(script_path)
        elif method == 'powerfactory':
            success = self.execute_in_powerfactory(script_path)
        else:
            print(f"✗ Unknown method: {method}")
            return False

        execution_time = time.time() - start_time
        print(f"\nExecution time: {execution_time:.2f} seconds")

        if wait_time > 0:
            print(f"Waiting {wait_time} seconds...")
            time.sleep(wait_time)

        return success
