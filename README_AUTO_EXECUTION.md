# Auto Execution DIgSILENT PowerFactory Scripts

Sistem untuk generate dan eksekusi skrip Python di DIgSILENT PowerFactory secara otomatis.

## Struktur File

```
test-digsilent/
├── digsilent_script_generator.py    # Generator untuk membuat skrip
├── digsilent_executor.py            # Executor untuk menjalankan skrip
├── example_auto_execution.py        # Contoh penggunaan lengkap
├── test_connection.py               # Test koneksi ke DIgSILENT
└── generated_scripts/               # Folder untuk skrip yang di-generate (otomatis dibuat)
```

## Cara Kerja

### Flow Proses:
```
Aplikasi Python → Generate Script → Execute di DIgSILENT → Hasil
```

### 1. Generate Script
`DIgSILENTScriptGenerator` menghasilkan skrip Python yang akan dijalankan di DIgSILENT:

```python
from digsilent_script_generator import DIgSILENTScriptGenerator

generator = DIgSILENTScriptGenerator()

# Generate load flow script
script_path = generator.generate_load_flow_script()

# Generate export script
script_path = generator.generate_export_results_script(export_path="results.csv")

# Generate custom script
custom_code = """
def my_function():
    app = pf.GetApplication()
    # Your code here
"""
script_path = generator.generate_custom_script("my_script", custom_code)
```

### 2. Execute Script
`DIgSILENTExecutor` menjalankan skrip yang sudah di-generate:

```python
from digsilent_executor import DIgSILENTExecutor

executor = DIgSILENTExecutor()

# Method 1: Execute langsung di PowerFactory context (RECOMMENDED)
success = executor.execute_in_powerfactory(script_path)

# Method 2: Execute dengan subprocess
success = executor.execute_script_subprocess(script_path)

# Method 3: Execute langsung
success = executor.execute_script_direct(script_path)
```

## Metode Eksekusi

### 1. `execute_in_powerfactory()` - RECOMMENDED
- Eksekusi langsung di context PowerFactory yang sedang running
- Paling cepat dan reliable
- Memerlukan PowerFactory sedang berjalan
- Akses langsung ke active project dan study case

### 2. `execute_script_subprocess()`
- Menjalankan di Python process terpisah
- Berguna untuk isolasi atau debugging
- Lebih lambat dari method 1

### 3. `execute_script_direct()`
- Eksekusi langsung dengan exec()
- Paling sederhana
- Bisa error jika ada issue dengan environment

## Contoh Penggunaan

### Contoh 1: Load Flow Calculation

```python
from digsilent_script_generator import DIgSILENTScriptGenerator
from digsilent_executor import DIgSILENTExecutor

# Generate script
generator = DIgSILENTScriptGenerator()
script_path = generator.generate_load_flow_script()

# Execute
executor = DIgSILENTExecutor()
success = executor.execute_in_powerfactory(script_path)

if success:
    print("Load flow completed!")
```

### Contoh 2: Export Results

```python
# Generate export script
script_path = generator.generate_export_results_script(
    export_path="d:/results/data.csv"
)

# Execute
success = executor.execute_in_powerfactory(script_path)
```

### Contoh 3: Custom Script

```python
# Define custom logic
custom_code = """
def analyze_network():
    app = pf.GetApplication()
    project = app.GetActiveProject()

    # Your custom analysis here
    terminals = app.GetCalcRelevantObjects("*.ElmTerm")
    print(f"Found {len(terminals)} buses")

    for term in terminals:
        voltage = term.GetAttribute('m:u')
        print(f"{term.GetAttribute('loc_name')}: {voltage} p.u.")

if __name__ == "__main__":
    analyze_network()
"""

# Generate and execute
script_path = generator.generate_custom_script("network_analysis", custom_code)
success = executor.execute_in_powerfactory(script_path)
```

### Contoh 4: Sequential Execution

```python
# Jalankan beberapa skrip berurutan
scripts = [
    generator.generate_load_flow_script(),
    generator.generate_export_results_script("results1.csv"),
    generator.generate_custom_script("analysis", custom_code),
]

executor = DIgSILENTExecutor()

for script in scripts:
    success = executor.execute_in_powerfactory(script)
    if not success:
        print(f"Failed at script: {script}")
        break
```

## Menjalankan Contoh

### Quick Start:

1. Pastikan DIgSILENT PowerFactory 2022 sedang berjalan
2. Buka project di PowerFactory
3. Jalankan contoh:

```bash
python example_auto_execution.py
```

### Test Koneksi Dulu:

```bash
python test_connection.py
```

## API Script Generator

### `DIgSILENTScriptGenerator(output_dir="generated_scripts")`

#### Methods:

**`generate_load_flow_script(project_name=None, study_case=None)`**
- Generate skrip untuk menjalankan load flow
- Returns: path ke skrip yang di-generate

**`generate_export_results_script(export_path="results.csv", elements=None)`**
- Generate skrip untuk export hasil kalkulasi
- Returns: path ke skrip yang di-generate

**`generate_custom_script(script_name, script_body)`**
- Generate custom skrip dengan code yang Anda tentukan
- Returns: path ke skrip yang di-generate

## API Executor

### `DIgSILENTExecutor()`

#### Methods:

**`execute_in_powerfactory(script_path)`**
- Eksekusi di PowerFactory context (RECOMMENDED)
- Returns: True/False

**`execute_script_subprocess(script_path, python_executable=None)`**
- Eksekusi di subprocess terpisah
- Returns: True/False

**`execute_script_direct(script_path)`**
- Eksekusi langsung dengan exec()
- Returns: True/False

**`execute_and_wait(script_path, method='direct', wait_time=2)`**
- Eksekusi dan tunggu sampai selesai
- Returns: True/False

## Troubleshooting

### Error: "No module named 'powerfactory'"
- Pastikan PowerFactory 2022 terinstall
- Check path di `digsilent_executor.py`
- Jalankan `test_connection.py` untuk cek instalasi

### Error: "Cannot connect to PowerFactory"
- Pastikan PowerFactory sedang berjalan
- Buka project di PowerFactory
- Coba restart PowerFactory

### Error: "No active project"
- Buka project di PowerFactory terlebih dahulu
- Atau specify project_name di generator

### Script tidak jalan
- Cek log error di console
- Pastikan skrip yang di-generate valid
- Test dengan contoh sederhana dulu

## Use Cases

### 1. Automated Analysis
Generate dan jalankan analisis secara otomatis dari aplikasi Python

### 2. Batch Processing
Proses multiple scenarios secara berurutan

### 3. Integration dengan System Lain
Integrate PowerFactory dengan system monitoring, database, dll

### 4. Scheduled Tasks
Jalankan analisis terjadwal dengan task scheduler

### 5. Web API Backend
Backend untuk web application yang control PowerFactory

## Tips

1. **Selalu cek koneksi dulu** dengan `test_connection.py`
2. **Gunakan `execute_in_powerfactory()`** untuk performa terbaik
3. **Generate script sekali, execute berkali-kali** jika logic sama
4. **Handle errors** dengan proper try-catch
5. **Log semua eksekusi** untuk debugging

## Compatibility

- **DIgSILENT PowerFactory 2021** (semua SP: base, SP1, SP2, SP3, SP4)
- **DIgSILENT PowerFactory 2022** (semua SP: base, SP1, SP2)
- Python 3.8, 3.9, 3.10, 3.11 (tergantung versi PowerFactory)
- Windows (tested on Windows 10/11)

### Catatan Versi:
- PowerFactory 2021 biasanya support Python 3.8, 3.9, 3.10
- PowerFactory 2022 biasanya support Python 3.8, 3.9, 3.10, 3.11
- Script otomatis mencari path kedua versi

## Next Steps

- Tambahkan lebih banyak template script
- Implement async execution
- Add callback untuk monitoring progress
- Create GUI untuk management script
