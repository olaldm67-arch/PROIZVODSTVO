import uuid

# --- Константы ---
STATUSES = ["Черновик", "На согласовании", "Утверждён", "Отклонён"]
ROLES = ["Автор", "Согласующий", "Администратор"]

# --- Классы ---
class User:
    def __init__(self, name, role):
        self.id = str(uuid.uuid4())
        self.name = name
        self.role = role

class Document:
    def __init__(self, title, content, author):
        self.id = str(uuid.uuid4())
        self.title = title
        self.content = content
        self.author = author
        self.status = STATUSES[0]

# --- Хранилища в памяти ---
users = []
documents = []

# --- Пример пользователей ---
users.append(User("Асхаб", "Автор"))
users.append(User("Тамаев", "Согласующий"))

# --- Функции ---
def add_document():
    author = next((u for u in users if u.role == "Автор"), None)
    if not author:
        print("Нет доступного автора.")
        return

    title = input("Введите название документа: ")
    content = input("Введите содержимое документа: ")

    doc = Document(title, content, author)
    documents.append(doc)
    print(f"Документ '{doc.title}' создан автором {author.name}.")

def list_documents():
    if not documents:
        print("Документов нет.")
        return

    for doc in documents:
        print(f"[{doc.id[:6]}] {doc.title} | Статус: {doc.status} | Автор: {doc.author.name}")

def change_status():
    list_documents()
    doc_id_part = input("Введите первые символы ID документа для изменения статуса: ")

    doc = next((d for d in documents if d.id.startswith(doc_id_part)), None)
    if not doc:
        print("Документ не найден.")
        return

    current_index = STATUSES.index(doc.status)
    if doc.status in ["Утверждён", "Отклонён"]:
        print("Изменение статуса невозможно.")
        return

    next_status = STATUSES[current_index + 1]
    confirm = input(f"Перевести '{doc.title}' в статус '{next_status}'? (y/n): ").lower()
    if confirm == 'y':
        doc.status = next_status
        print(f"Статус документа обновлён: {doc.status}")

# --- Меню ---
def main():
    while True:
        print("\n--- ЭДО Меню ---")
        print("1. Добавить документ")
        print("2. Список документов")
        print("3. Перевести документ в следующий статус")
        print("4. Выход")
        choice = input("Выберите действие: ")

        if choice == "1":
            add_document()
        elif choice == "2":
            list_documents()
        elif choice == "3":
            change_status()
        elif choice == "4":
            print("Выход из программы.")
            break
        else:
            print("Неверный выбор.")

if __name__ == "__main__":
    main()