"""
Script untuk diagnose instalasi PowerFactory
Membantu menemukan lokasi DLL yang benar
"""

import os
import sys

def diagnose_installation(python_path):
    """
    Diagnose PowerFactory installation structure

    Args:
        python_path: Path ke folder Python PowerFactory
                    Contoh: D:\Digsilent Powerfactory 2021\Digsilent\Python\3.8
    """
    print("="*60)
    print("PowerFactory Installation Diagnostic")
    print("="*60)

    print(f"\nPython Path: {python_path}")

    if not os.path.exists(python_path):
        print(f"✗ Path tidak ditemukan: {python_path}")
        return

    print("✓ Path ditemukan")

    # Check files in Python path
    print("\n" + "-"*60)
    print("Files in Python path:")
    print("-"*60)
    try:
        files = os.listdir(python_path)
        for f in files[:20]:  # Show first 20 files
            full_path = os.path.join(python_path, f)
            if os.path.isfile(full_path):
                print(f"  📄 {f}")
            else:
                print(f"  📁 {f}/")
        if len(files) > 20:
            print(f"  ... and {len(files) - 20} more files")
    except Exception as e:
        print(f"✗ Error listing files: {e}")

    # Go up levels and check structure
    current = python_path
    for level in range(5):
        parent = os.path.dirname(current)
        if parent == current:  # Reached root
            break

        print("\n" + "-"*60)
        print(f"Level {level+1} up: {parent}")
        print("-"*60)

        try:
            items = os.listdir(parent)
            print(f"Contents ({len(items)} items):")
            for item in items[:30]:
                full_path = os.path.join(parent, item)
                if os.path.isfile(full_path):
                    # Check if it's a DLL
                    if item.lower().endswith('.dll'):
                        print(f"  🔵 {item} (DLL)")
                    elif item.lower().endswith('.exe'):
                        print(f"  ⚙️  {item} (EXE)")
                    else:
                        print(f"  📄 {item}")
                else:
                    # Check if it's bin folder
                    if item.lower() == 'bin':
                        print(f"  📁 {item}/ ⭐ (BIN FOLDER)")
                    else:
                        print(f"  📁 {item}/")
            if len(items) > 30:
                print(f"  ... and {len(items) - 30} more items")
        except Exception as e:
            print(f"✗ Error: {e}")

        current = parent

    # Check common DLL locations
    print("\n" + "="*60)
    print("Checking common DLL locations:")
    print("="*60)

    base_candidates = [
        os.path.dirname(os.path.dirname(python_path)),  # Up 2 levels
        os.path.dirname(os.path.dirname(os.path.dirname(python_path))),  # Up 3 levels
    ]

    dll_locations = []

    for base in base_candidates:
        if not os.path.exists(base):
            continue

        print(f"\nChecking: {base}")

        # Check base directory
        try:
            files = os.listdir(base)
            dll_files = [f for f in files if f.lower().endswith('.dll')]
            if dll_files:
                print(f"  ✓ Found {len(dll_files)} DLL files in base directory")
                dll_locations.append(base)
                for dll in dll_files[:5]:
                    print(f"    - {dll}")
                if len(dll_files) > 5:
                    print(f"    ... and {len(dll_files) - 5} more")
        except:
            pass

        # Check bin folder
        bin_folder = os.path.join(base, 'bin')
        if os.path.exists(bin_folder):
            try:
                files = os.listdir(bin_folder)
                dll_files = [f for f in files if f.lower().endswith('.dll')]
                if dll_files:
                    print(f"  ✓ Found {len(dll_files)} DLL files in bin/ folder")
                    dll_locations.append(bin_folder)
                    for dll in dll_files[:5]:
                        print(f"    - {dll}")
                    if len(dll_files) > 5:
                        print(f"    ... and {len(dll_files) - 5} more")
            except:
                pass

    # Summary
    print("\n" + "="*60)
    print("SUMMARY & RECOMMENDATIONS")
    print("="*60)

    if dll_locations:
        print("\n✓ Found DLL locations:")
        for loc in dll_locations:
            print(f"  - {loc}")

        print("\n📋 Add these paths to PATH environment:")
        print("```python")
        for loc in dll_locations:
            print(f'os.environ["PATH"] = r"{loc}" + os.pathsep + os.environ["PATH"]')
        print("```")
    else:
        print("\n✗ No DLL locations found!")
        print("\nTroubleshooting:")
        print("1. Pastikan PowerFactory terinstall dengan benar")
        print("2. Cek apakah ada file .exe PowerFactory di folder instalasi")
        print("3. Coba jalankan PowerFactory untuk memastikan aplikasi bisa jalan")

    print("\n" + "="*60)


if __name__ == "__main__":
    # Default path - ganti sesuai instalasi Anda
    default_path = r"D:\Digsilent Powerfactory 2021\Digsilent\Python\3.8"

    print("PowerFactory Diagnostic Tool")
    print("="*60)
    print(f"\nDefault path: {default_path}")

    user_input = input("\nMasukkan path Python PowerFactory (atau tekan Enter untuk default): ").strip()

    if user_input:
        path_to_check = user_input
    else:
        path_to_check = default_path

    diagnose_installation(path_to_check)
