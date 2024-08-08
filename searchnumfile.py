import os
import textract

def search_in_file(file_path, search_number):
    try:
        # Извлекаем текст из файла
        text = textract.process(file_path).decode('utf-8')
        # Проверяем, содержится ли искомый номер в тексте
        return search_number in text
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return False

def search_number_in_files(root_folder, search_number):
    results = []
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            # Проверяем расширения файлов и избегаем временных файлов
            if filename.endswith(('.docx', '.xlsx')) and not filename.startswith('~$'):
                if search_in_file(file_path, search_number):
                    results.append(file_path)
    return results

# Пример использования
root_folder = r'C:\Users\Emil.Amiruldayev\Desktop\HTT distribution list'  # Исправьте на абсолютный путь
search_number = 'RZ8KC1K871D'  # Номер, который необходимо найти
found_files = search_number_in_files(root_folder, search_number)

if found_files:
    print("Найдено в следующих файлах:")
    for file in found_files:
        print(file)
else:
    print("Номер не найден.")
