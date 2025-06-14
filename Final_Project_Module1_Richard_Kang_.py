# Dictionary to store student data
data_nilai = {}

# Input helper
def input_data():
    return {
        "nama": input("Masukkan nama siswa: "),
        "umur": input("Masukkan umur siswa: "),
        "jenis_kelamin": input("Masukkan jenis kelamin siswa (L/P): "),
        "alamat": input("Masukkan alamat siswa: "),
        "nilai": input("Masukkan nilai siswa: ")
    }

# Create
def tambah_data():
    id = input("Masukkan ID siswa: ")
    if id in data_nilai:
        print("Siswa sudah ada.")
    else:
        data_nilai[id] = input_data()
        print("Data berhasil ditambahkan.")

# Read
def tampilkan_data():
    while True:
        print("\n=== Report Nilai Siswa ===")
        print("1. Semua Data")
        print("2. Data Tertentu")
        print("3. Kembali")
        pilihan = input("Pilih menu (1-3): ")

        if pilihan == "1":
            if not data_nilai:
                print("Belum ada data.")
            else:
                for id, info in data_nilai.items():
                    print(f"{id}: {info}")
        elif pilihan == "2":
            id = input("Masukkan ID siswa: ")
            print(f"{id}: {data_nilai.get(id, 'Data tidak ditemukan')}")
        elif pilihan == "3":
            break
        else:
            print("Pilihan tidak valid.")

# Update
def ubah_data():
    id = input("Masukkan ID siswa yang ingin diubah: ")
    if id not in data_nilai:
        print("Data siswa tidak ditemukan.")
        return

    print(f"{id}: {data_nilai[id]}")
    if input("Lanjutkan update? (Y/N): ").lower() != "y":
        return

    fields = list(data_nilai[id].keys())
    for i, key in enumerate(fields, start=1):
        print(f"{i}. Ubah {key}")
    pilihan = input("Pilih field yang ingin diubah (1-5): ")

    if pilihan.isdigit() and 1 <= int(pilihan) <= len(fields):
        field = fields[int(pilihan) - 1]
        data_nilai[id][field] = input(f"Masukkan {field} baru: ")
        print("Data berhasil diupdate.")
    else:
        print("Pilihan tidak valid.")

# Delete
def hapus_data():
    id = input("Masukkan ID siswa yang ingin dihapus: ")
    if data_nilai.pop(id, None):
        print("Data berhasil dihapus.")
    else:
        print("Siswa tidak ditemukan.")

# Fitur Tambahan
from openpyxl import Workbook
# Export Ke Excel
def export_ke_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Siswa"

    # Header
    ws.append(["ID", "Nama", "Umur", "Jenis Kelamin", "Alamat", "Nilai"])

    # Isi data
    for id, info in data_nilai.items():
        ws.append([
            id,
            info["nama"],
            info["umur"],
            info["jenis_kelamin"],
            info["alamat"],
            info["nilai"]
        ])

    wb.save("data_siswa.xlsx")
    print("Data berhasil diekspor ke 'data_siswa.xlsx'")


# Main Menu
def menu():
    while True:
        print("\n=== Menu Utama ===")
        print("1. Tambah Data")
        print("2. Tampilkan Data")
        print("3. Ubah Data")
        print("4. Hapus Data")
        print("5. Ekspor Data ke Excel")
        print("6. Keluar")
        pilihan = input("Pilih menu (1-6): ")

        if pilihan == "1":
            tambah_data()
        elif pilihan == "2":
            tampilkan_data()
        elif pilihan == "3":
            ubah_data()
        elif pilihan == "4":
            hapus_data()
        elif pilihan == "5":
            export_ke_excel()
        elif pilihan == "6":
            print("Terima kasih.")
            break
        else:
            print("Pilihan tidak valid.")

# Run the program
menu()
