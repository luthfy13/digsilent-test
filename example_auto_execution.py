"""
Contoh penggunaan: Generate dan eksekusi skrip DIgSILENT secara otomatis

Flow:
1. Aplikasi Python generate skrip
2. Skrip otomatis dieksekusi di DIgSILENT
3. Hasil ditampilkan
"""

from digsilent_script_generator import DIgSILENTScriptGenerator
from digsilent_executor import DIgSILENTExecutor
import os


def example_1_load_flow():
    """
    Contoh 1: Generate dan jalankan load flow calculation
    """
    print("="*60)
    print("CONTOH 1: Load Flow Calculation")
    print("="*60)

    # Step 1: Generate script
    generator = DIgSILENTScriptGenerator()
    script_path = generator.generate_load_flow_script()

    print(f"\n✓ Script generated: {script_path}")

    # Step 2: Execute script
    executor = DIgSILENTExecutor()
    print("\nExecuting script in PowerFactory...")

    success = executor.execute_in_powerfactory(script_path)

    if success:
        print("\n✓ Load Flow calculation completed successfully!")
    else:
        print("\n✗ Load Flow calculation failed!")

    return success


def example_2_export_results():
    """
    Contoh 2: Generate dan jalankan export results
    """
    print("\n\n")
    print("="*60)
    print("CONTOH 2: Export Results")
    print("="*60)

    # Step 1: Generate script
    generator = DIgSILENTScriptGenerator()
    export_path = os.path.join(os.getcwd(), "results", "exported_data.csv")

    # Buat folder results jika belum ada
    os.makedirs(os.path.dirname(export_path), exist_ok=True)

    script_path = generator.generate_export_results_script(export_path=export_path)

    print(f"\n✓ Script generated: {script_path}")

    # Step 2: Execute script
    executor = DIgSILENTExecutor()
    print("\nExecuting script in PowerFactory...")

    success = executor.execute_in_powerfactory(script_path)

    if success:
        print(f"\n✓ Results exported to: {export_path}")
    else:
        print("\n✗ Export failed!")

    return success


def example_3_custom_script():
    """
    Contoh 3: Generate dan jalankan custom script
    """
    print("\n\n")
    print("="*60)
    print("CONTOH 3: Custom Script - Get Network Elements")
    print("="*60)

    # Custom script body
    custom_code = """
def get_network_info():
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

    print(f"\\nActive Project: {project.GetFullName()}")

    # Get network elements
    print("\\n" + "="*60)
    print("Network Elements:")
    print("="*60)

    # Buses
    terminals = app.GetCalcRelevantObjects("*.ElmTerm")
    print(f"\\nBuses/Terminals: {len(terminals)}")
    for i, term in enumerate(terminals[:5]):  # Show first 5
        print(f"  {i+1}. {term.GetAttribute('loc_name')}")
    if len(terminals) > 5:
        print(f"  ... and {len(terminals) - 5} more")

    # Lines
    lines = app.GetCalcRelevantObjects("*.ElmLne")
    print(f"\\nLines: {len(lines)}")
    for i, line in enumerate(lines[:5]):  # Show first 5
        print(f"  {i+1}. {line.GetAttribute('loc_name')}")
    if len(lines) > 5:
        print(f"  ... and {len(lines) - 5} more")

    # Generators
    generators = app.GetCalcRelevantObjects("*.ElmSym")
    print(f"\\nGenerators: {len(generators)}")
    for i, gen in enumerate(generators[:5]):  # Show first 5
        print(f"  {i+1}. {gen.GetAttribute('loc_name')}")
    if len(generators) > 5:
        print(f"  ... and {len(generators) - 5} more")

    # Loads
    loads = app.GetCalcRelevantObjects("*.ElmLod")
    print(f"\\nLoads: {len(loads)}")
    for i, load in enumerate(loads[:5]):  # Show first 5
        print(f"  {i+1}. {load.GetAttribute('loc_name')}")
    if len(loads) > 5:
        print(f"  ... and {len(loads) - 5} more")

    print("\\n" + "="*60)
    return True

if __name__ == "__main__":
    get_network_info()
"""

    # Step 1: Generate script
    generator = DIgSILENTScriptGenerator()
    script_path = generator.generate_custom_script("get_network_info", custom_code)

    print(f"\n✓ Script generated: {script_path}")

    # Step 2: Execute script
    executor = DIgSILENTExecutor()
    print("\nExecuting script in PowerFactory...")

    success = executor.execute_in_powerfactory(script_path)

    if success:
        print("\n✓ Custom script executed successfully!")
    else:
        print("\n✗ Custom script failed!")

    return success


def example_4_sequential_execution():
    """
    Contoh 4: Eksekusi beberapa skrip secara berurutan
    """
    print("\n\n")
    print("="*60)
    print("CONTOH 4: Sequential Execution")
    print("="*60)

    generator = DIgSILENTScriptGenerator()
    executor = DIgSILENTExecutor()

    # Step 1: Run load flow
    print("\nStep 1: Running Load Flow...")
    script1 = generator.generate_load_flow_script()
    success1 = executor.execute_in_powerfactory(script1)

    if not success1:
        print("✗ Load Flow failed, stopping execution")
        return False

    # Step 2: Export results
    print("\n\nStep 2: Exporting Results...")
    export_path = os.path.join(os.getcwd(), "results", "sequential_export.csv")
    os.makedirs(os.path.dirname(export_path), exist_ok=True)

    script2 = generator.generate_export_results_script(export_path=export_path)
    success2 = executor.execute_in_powerfactory(script2)

    if success1 and success2:
        print("\n\n" + "="*60)
        print("✓ All scripts executed successfully!")
        print("="*60)
        return True
    else:
        print("\n\n" + "="*60)
        print("✗ Some scripts failed!")
        print("="*60)
        return False


def main():
    """
    Main function - jalankan semua contoh
    """
    print("="*60)
    print("AUTO EXECUTION EXAMPLES")
    print("DIgSILENT PowerFactory Script Generator & Executor")
    print("="*60)

    print("\nPastikan DIgSILENT PowerFactory sedang berjalan")
    print("dan ada project yang aktif!")

    input("\nTekan Enter untuk melanjutkan...")

    # Pilih contoh mana yang akan dijalankan
    print("\n\nPilih contoh:")
    print("1. Load Flow Calculation")
    print("2. Export Results")
    print("3. Custom Script - Get Network Info")
    print("4. Sequential Execution (Load Flow + Export)")
    print("5. Jalankan semua contoh")

    choice = input("\nPilihan (1-5): ").strip()

    if choice == "1":
        example_1_load_flow()
    elif choice == "2":
        example_2_export_results()
    elif choice == "3":
        example_3_custom_script()
    elif choice == "4":
        example_4_sequential_execution()
    elif choice == "5":
        print("\n\nMenjalankan semua contoh...\n")
        example_1_load_flow()
        example_2_export_results()
        example_3_custom_script()
        example_4_sequential_execution()
    else:
        print("Pilihan tidak valid")


if __name__ == "__main__":
    main()
