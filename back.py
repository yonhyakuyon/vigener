import pandas as pd


def read_vigenere_alphabet(excel_file):
    """Читает алфавит Виженера из Excel-файла."""
    df = pd.read_excel(excel_file, header=None)
    return df.values


def vigenere_decrypt(encrypted_message, key, alphabet):
    """Расшифровывает сообщение методом Виженера."""
    decrypted_message = []
    key = key.lower()
    key_length = len(key)
    alphabet_dict = {letter: i for i, letter in enumerate(alphabet[0])}

    key_indices = [alphabet_dict[char] for char in key]

    for i, char in enumerate(encrypted_message.lower()):
        if char in alphabet_dict:
            row_index = alphabet_dict[key[i % key_length]]
            col_index = alphabet[row_index].tolist().index(char)
            decrypted_char = alphabet[0][col_index]
            decrypted_message.append(decrypted_char)
        else:
            # Если символ не в алфавите, оставить без изменений
            decrypted_message.append(char)

    return "".join(decrypted_message)


def decrypt_main():
    # Считываем алфавит Виженера из Excel
    excel_file = input("Введите путь к Excel-файлу с алфавитом Виженера: ")
    alphabet = read_vigenere_alphabet(excel_file)

    # Считываем зашифрованное сообщение из текстового файла
    encrypted_file = input(
        "Введите путь к текстовому файлу с зашифрованным сообщением: "
    )
    with open(encrypted_file, "r", encoding="utf-8") as file:
        encrypted_message = file.read().strip()

    # Пользователь вводит ключ
    key = input("Введите ключ для расшифровки: ").strip()

    # Расшифровываем сообщение
    decrypted_message = vigenere_decrypt(encrypted_message, key, alphabet)

    # Выводим расшифрованное сообщение
    print("\nРасшифрованное сообщение:")
    print(decrypted_message)

    # Сохраняем расшифрованное сообщение в текстовый файл
    output_file = input(
        "Введите путь для сохранения расшифрованного сообщения (текстовый файл): "
    )
    append_to_text_file(output_file, decrypted_message)
    print(f"Расшифрованное сообщение добавлено в файл: {output_file}")


if __name__ == "__main__":
    decrypt_main()
