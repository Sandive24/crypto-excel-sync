import fungsi_coingecko
import fungsi_coinmarketcap
import fungsi_reward
import time


# Fungsi untuk menghitung waktu yang telah berlalu
def elapsed_time(start, end):
    elapsed = end - start
    minutes, seconds = divmod(elapsed, 60)
    return f"{int(minutes)} Menit, {int(seconds)} Detik"

# Mulai pengukuran waktu
start_time = time.time()




# PATH 1
excel_path_1 = r"C:\My DATA\Gitub Project\price_update\doc\tes.xlsx"
print(f"\nPath 1 : {excel_path_1}")
fungsi_coingecko.run(excel_path_1)
print("\n")
fungsi_coinmarketcap.main(excel_path_1)

# # PATH 2
# print("\n"*5)
# excel_path_2 = r"C:\My DATA\Gitub Project\price_update\coba.xlsx"
# print(f"Path 2 : {excel_path_2}")
# fungsi_coingecko.run(excel_path_2)
# print("\n")
# fungsi_coinmarketcap.main(excel_path_2)



# UPDATE REWARD
fungsi_reward.main()



# Akhiri pengukuran waktu
end_time = time.time()

# Cetak total waktu pemrosesan
print("\n"*5)
print(f"Processing Time: {elapsed_time(start_time, end_time)}")