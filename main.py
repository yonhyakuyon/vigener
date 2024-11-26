import pandas as pd


def read_vigenere_alphabet(excel_file):
    """Читает алфавит Виженера из Excel-файла."""
    df = pd.read_excel(excel_file, header=None)
    return df.values


def vigenere_encrypt(message, key, alphabet):
    """Шифрует сообщение методом Виженера."""
    encrypted_message = []
    key = key.lower()
    key_length = len(key)
    alphabet_dict = {letter: i for i, letter in enumerate(alphabet[0])}

    key_indices = [alphabet_dict[char] for char in key]

    for i, char in enumerate(message.lower()):
        if char in alphabet_dict:
            row_index = alphabet_dict[key[i % key_length]]
            col_index = alphabet_dict[char]
            encrypted_char = alphabet[row_index][col_index]
            encrypted_message.append(encrypted_char)
        else:
            # Если символ не в алфавите, оставить без изменений
            encrypted_message.append(char)

    return "".join(encrypted_message)


def append_to_text_file(file, encrypted_message):
    """Добавляет зашифрованное сообщение в текстовый файл на новую строку."""
    with open(file, "a", encoding="utf-8") as f:
        f.write(encrypted_message + "\n")


def main():
    # Считываем алфавит Виженера из Excel
    excel_file = input("Введите путь к Excel-файлу с алфавитом Виженера: ")
    alphabet = read_vigenere_alphabet(excel_file)

    # Считываем сообщение из текстового файла
    text_file = input("Введите путь к текстовому файлу с сообщением: ")
    with open(text_file, "r", encoding="utf-8") as file:
        message = file.read().strip()

    # Пользователь вводит ключ
    key = input("Введите ключ для шифрования: ").strip()

    # Шифруем сообщение
    encrypted_message = vigenere_encrypt(message, key, alphabet)

    # Выводим зашифрованное сообщение
    print("\nЗашифрованное сообщение:")
    print(encrypted_message)

    # Сохраняем зашифрованное сообщение в текстовый файл
    output_file = input(
        "Введите путь для сохранения зашифрованного сообщения (текстовый файл): "
    )
    append_to_text_file(output_file, encrypted_message)
    print(f"Зашифрованное сообщение добавлено в файл: {output_file}")


if __name__ == "__main__":
    main()
