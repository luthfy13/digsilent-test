"""
Module untuk generate skrip DIgSILENT PowerFactory
"""

import os
from datetime import datetime


class DIgSILENTScriptGenerator:
    """
    Class untuk generate skrip yang akan dijalankan di DIgSILENT
    """

    def __init__(self, output_dir="generated_scripts"):
        """
        Initialize generator

        Args:
            output_dir: Folder untuk menyimpan skrip yang di-generate
        """
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

    def generate_load_flow_script(self, project_name=None, study_case=None):
        """
        Generate skrip untuk menjalankan load flow calculation

        Args:
            project_name: Nama project (optional)
            study_case: Nama study case (optional)

        Returns:
            Path ke file skrip yang di-generate
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_name = f"loadflow_{timestamp}.py"
        script_path = os.path.join(self.output_dir, script_name)

        script_content = '''"""
Auto-generated script untuk Load Flow Calculation
Generated at: {timestamp}
"""

import powerfactory as pf

def run_load_flow():
    # Get PowerFactory application
    app = pf.GetApplication()
    if app is None:
        print("Error: Cannot connect to PowerFactory")
        return False

    print("Connected to PowerFactory")

    # Get active project
    project = app.GetActiveProject()
    if project is None:
        print("Error: No active project")
        return False

    print(f"Active Project: {{project.GetFullName()}}")

    # Activate study case if specified
{study_case_code}

    # Get Load Flow command
    ldf = app.GetFromStudyCase("ComLdf")
    if ldf is None:
        print("Error: Cannot get Load Flow command")
        return False

    # Execute Load Flow
    print("Executing Load Flow...")
    result = ldf.Execute()

    if result == 0:
        print("✓ Load Flow calculation successful")
        return True
    else:
        print(f"✗ Load Flow calculation failed with error code: {{result}}")
        return False

if __name__ == "__main__":
    success = run_load_flow()
    print("\\n" + "="*60)
    if success:
        print("SCRIPT COMPLETED SUCCESSFULLY")
    else:
        print("SCRIPT FAILED")
    print("="*60)
'''.format(
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            study_case_code=self._generate_study_case_code(study_case) if study_case else "    # Using current active study case"
        )

        with open(script_path, 'w') as f:
            f.write(script_content)

        print(f"✓ Generated script: {script_path}")
        return script_path

    def generate_export_results_script(self, export_path="results.csv", elements=None):
        """
        Generate skrip untuk export hasil kalkulasi

        Args:
            export_path: Path untuk file export
            elements: List elemen yang akan di-export (optional)

        Returns:
            Path ke file skrip yang di-generate
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_name = f"export_results_{timestamp}.py"
        script_path = os.path.join(self.output_dir, script_name)

        script_content = '''"""
Auto-generated script untuk Export Results
Generated at: {timestamp}
"""

import powerfactory as pf
import csv

def export_results():
    # Get PowerFactory application
    app = pf.GetApplication()
    if app is None:
        print("Error: Cannot connect to PowerFactory")
        return False

    print("Connected to PowerFactory")

    # Get active project
    project = app.GetActiveProject()
    if project is None:
        print("Error: No active project")
        return False

    print(f"Active Project: {{project.GetFullName()}}")

    # Get all terminals (busbar)
    terminals = app.GetCalcRelevantObjects("*.ElmTerm")

    print(f"Found {{len(terminals)}} terminals")

    # Export to CSV
    export_file = r"{export_path}"
    print(f"Exporting to: {{export_file}}")

    with open(export_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        # Header
        writer.writerow(['Name', 'Voltage (kV)', 'Angle (deg)', 'Type'])

        # Data
        for term in terminals:
            name = term.GetAttribute('loc_name')
            voltage = term.GetAttribute('m:u')  # Voltage in p.u.
            angle = term.GetAttribute('m:phiu')  # Voltage angle in degrees
            term_type = term.GetAttribute('iUsage')

            writer.writerow([name, voltage, angle, term_type])

    print(f"✓ Exported {{len(terminals)}} terminals to {{export_file}}")
    return True

if __name__ == "__main__":
    success = export_results()
    print("\\n" + "="*60)
    if success:
        print("EXPORT COMPLETED SUCCESSFULLY")
    else:
        print("EXPORT FAILED")
    print("="*60)
'''.format(
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            export_path=export_path
        )

        with open(script_path, 'w') as f:
            f.write(script_content)

        print(f"✓ Generated script: {script_path}")
        return script_path

    def generate_custom_script(self, script_name, script_body):
        """
        Generate custom skrip

        Args:
            script_name: Nama file skrip
            script_body: Isi skrip (Python code)

        Returns:
            Path ke file skrip yang di-generate
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if not script_name.endswith('.py'):
            script_name += '.py'

        script_name = f"{script_name.replace('.py', '')}_{timestamp}.py"
        script_path = os.path.join(self.output_dir, script_name)

        script_content = f'''"""
Auto-generated custom script
Generated at: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
"""

import powerfactory as pf

{script_body}
'''

        with open(script_path, 'w') as f:
            f.write(script_content)

        print(f"✓ Generated script: {script_path}")
        return script_path

    def _generate_study_case_code(self, study_case_name):
        """Generate code untuk aktivasi study case"""
        return f'''    # Activate study case
    study_case = project.GetContents("{study_case_name}")[0]
    if study_case:
        study_case.Activate()
        print(f"Activated study case: {study_case_name}")
    else:
        print(f"Warning: Study case '{study_case_name}' not found")
'''
