from menu_siswa import menu_siswa
from menu_tagihan import menu_tagihan
from menu_pembayaran import menu_pembayaran
from menu_laporan_servis import menu_laporan_servis


# MENGELOLA MENU UTAMA
def menu_utama():
      while True:
            print("\nMenu utama Perpustakaan: ")
            print("1. Siswa")
            print("2. Tagihan")
            print("3. Pembayaran")
            print("4. Laporan servis")
            print("5. Keluar")

            pilihan = input("Pilih menu: ")

            if pilihan == '1':
                  menu_siswa()
            elif pilihan == '2':
                  menu_tagihan()
            elif pilihan == '3':
                  menu_pembayaran()
            elif pilihan == '4':
                  menu_laporan_servis()
            elif pilihan == '5':
                  print("Terimakasih telah menggunakan sistem pembayaran SPP.")
                  break
            
# Menjalankan menu_utama itu sebagai program utama
if __name__ == "__main__":
    menu_utama()
