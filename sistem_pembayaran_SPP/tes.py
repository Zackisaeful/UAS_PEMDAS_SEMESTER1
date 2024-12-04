from datetime import datetime

# mengambil bulan saat ini
tanggal_penagihan = datetime.today().strftime('%d-%m-%Y')
print(tanggal_penagihan)

bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", 
         "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

nomor_bulan = int(datetime.today().strftime('%m'))
print(bulan[nomor_bulan - 1])  