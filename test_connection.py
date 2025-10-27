"""
Script untuk testing koneksi ke DIgSILENT PowerFactory
"""

import sys
import os

def add_powerfactory_path():
    """
    Menambahkan path PowerFactory ke sys.path dan PATH environment (support 2021 dan 2022)
    """
    # Possible PowerFactory installation paths
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
        # Custom paths (non-standard installation)
        r"D:\Digsilent Powerfactory 2021\Digsilent\Python\3.8",
        r"D:\Digsilent Powerfactory 2021\Digsilent\Python\3.9",
        r"D:\Digsilent Powerfactory 2021\Digsilent\Python\3.10",
    ]

    print("Mencari instalasi PowerFactory...")
    for path in possible_paths:
        if os.path.exists(path):
            print(f"   ✓ Ditemukan: {path}")

            # Add to sys.path
            if path not in sys.path:
                sys.path.append(path)
                print(f"   ✓ Path ditambahkan ke sys.path")

            # Get base PowerFactory directory (go up 2 levels from Python\3.x)
            pf_base = os.path.dirname(os.path.dirname(path))
            print(f"   ✓ PowerFactory base directory: {pf_base}")

            # Add DLL paths to PATH environment
            # PowerFactory DLLs biasanya ada di base directory
            if pf_base not in os.environ['PATH']:
                os.environ['PATH'] = pf_base + os.pathsep + os.environ['PATH']
                print(f"   ✓ DLL path ditambahkan ke PATH environment")

            # Juga cek folder bin jika ada
            bin_path = os.path.join(pf_base, 'bin')
            if os.path.exists(bin_path) and bin_path not in os.environ['PATH']:
                os.environ['PATH'] = bin_path + os.pathsep + os.environ['PATH']
                print(f"   ✓ Bin path ditambahkan: {bin_path}")

            return True

    print("   ✗ Instalasi PowerFactory tidak ditemukan di lokasi default")
    print("\n   Silakan tambahkan path manual:")
    print("   Cari folder instalasi PowerFactory 2021/2022, kemudian tambahkan:")
    print("   sys.path.append(r'C:\\Path\\To\\PowerFactory 20XX\\Python\\3.x')")
    return False

def test_digsilent_connection():
    """
    Test koneksi ke DIgSILENT PowerFactory
    """
    print("="*60)
    print("Testing Koneksi DIgSILENT PowerFactory")
    print("="*60)

    # Step 0: Add PowerFactory path
    print("\n0. Menambahkan PowerFactory path...")
    if not add_powerfactory_path():
        print("\n   Coba cari manual lokasi instalasi PowerFactory 2022")
        return False

    # Step 1: Check if PowerFactory is installed
    print("\n1. Checking DIgSILENT PowerFactory installation...")

    try:
        import powerfactory
        print("   ✓ Module powerfactory ditemukan")
    except ImportError as e:
        print("   ✗ Module powerfactory tidak ditemukan")
        print("   Error:", str(e))
        print("\n   Pastikan DIgSILENT PowerFactory sudah terinstall")
        print("   dan Python API sudah dikonfigurasi dengan benar")
        return False

    # Step 2: Try to connect to PowerFactory
    print("\n2. Mencoba koneksi ke PowerFactory...")

    try:
        app = powerfactory.GetApplication()

        if app is None:
            print("   ✗ Gagal mendapatkan aplikasi PowerFactory")
            print("   Pastikan PowerFactory sedang berjalan")
            return False

        print("   ✓ Berhasil terhubung ke PowerFactory")

        # Step 3: Get PowerFactory version
        print("\n3. Informasi PowerFactory:")
        try:
            version = app.GetVersion()
            print(f"   - Version: {version}")
        except:
            print("   - Version: (tidak dapat diambil)")

        # Step 4: Get current user
        try:
            user = app.GetCurrentUser()
            if user:
                print(f"   - User: {user}")
            else:
                print("   - User: (tidak ada user aktif)")
        except:
            print("   - User: (tidak dapat diambil)")

        # Step 5: Get active project
        print("\n4. Checking active project...")
        try:
            project = app.GetActiveProject()
            if project:
                print(f"   ✓ Active Project: {project.GetFullName()}")
            else:
                print("   ⚠ Tidak ada project yang aktif")
                print("   (Ini normal jika belum ada project yang dibuka)")
        except Exception as e:
            print(f"   ⚠ Error saat mengecek project: {str(e)}")

        # Step 6: Get study case
        print("\n5. Checking active study case...")
        try:
            study_case = app.GetActiveStudyCase()
            if study_case:
                print(f"   ✓ Active Study Case: {study_case.GetFullName()}")
            else:
                print("   ⚠ Tidak ada study case yang aktif")
        except Exception as e:
            print(f"   ⚠ Error saat mengecek study case: {str(e)}")

        print("\n" + "="*60)
        print("✓ Koneksi ke DIgSILENT PowerFactory BERHASIL!")
        print("="*60)
        return True

    except Exception as e:
        print(f"   ✗ Error saat koneksi: {str(e)}")
        print("\n   Troubleshooting:")
        print("   - Pastikan DIgSILENT PowerFactory sedang berjalan")
        print("   - Pastikan Python API sudah dikonfigurasi")
        print("   - Cek apakah path Python API sudah benar di sys.path")
        print("\n" + "="*60)
        print("✗ Koneksi GAGAL")
        print("="*60)
        return False

if __name__ == "__main__":
    # Jalankan test
    success = test_digsilent_connection()

    # Exit dengan status code
    sys.exit(0 if success else 1)
