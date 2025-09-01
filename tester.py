import cProfile
import pstats
import openpyxl

# --- ПУТЬ К МЕДЛЕННОМУ ИСХОДНОМУ ФАЙЛУ ---
file_to_profile = 'source/Financial memorandum SW_29.08.2025.xlsx' # ЗАМЕНИТЕ НА РЕАЛЬНЫЙ ПУТЬ

def load_standard():
    """Загружает файл в стандартном режиме."""
    wb = openpyxl.load_workbook(file_to_profile, data_only=True)
    wb.close()

def load_readonly():
    """Загружает файл в режиме read_only."""
    wb = openpyxl.load_workbook(file_to_profile, data_only=True, read_only=True)
    wb.close()

print(f"--- Профилирование стандартной загрузки файла {file_to_profile} ---")
cProfile.run('load_standard()', 'stats_standard.prof')

print(f"\n--- Профилирование read_only загрузки файла {file_to_profile} ---")
cProfile.run('load_readonly()', 'stats_readonly.prof')

# --- Вывод результатов ---
print("\n\n--- ТОП-10 самых медленных функций (стандартный режим) ---")
p_standard = pstats.Stats('stats_standard.prof')
p_standard.sort_stats('tottime').print_stats(10)

print("\n--- ТОП-10 самых медленных функций (read_only режим) ---")
p_readonly = pstats.Stats('stats_readonly.prof')
p_readonly.sort_stats('tottime').print_stats(10)