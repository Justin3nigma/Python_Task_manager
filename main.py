import pandas as pd
import pickle

def display_menu():
    print("==== To-Do List Manager ====")
    print("1. Add Task")
    print("2. View Tasks")
    print("3. Mark Task as Completed")
    print("4. Delete Task")
    print("5. Import Tasks from Excel")
    print("6. Export Tasks to Excel")
    print("7. Quit")
    print("============================")
def add_task(todo_list):
    task = input("Enter task description: ")
    todo_list.append({"task": task, "completed": False})
    print("Task added successfully!")

def view_tasks(todo_list):
    if not todo_list:
        print("No tasks found.")
    else:
        for index, task in enumerate(todo_list):
            status = "Completed" if task["completed"] else "Not Completed"
            print(f"{index+1}. {task['task']} [{status}]")

def mark_task_completed(todo_list):
    task_number = int(input("Enter task number to mark as completed: "))
    if task_number < 1 or task_number > len(todo_list):
        print("Invalid task number.")
    else:
        task = todo_list[task_number - 1]
        task["completed"] = True
        print("Task marked as completed!")

def delete_task(todo_list):
    task_number = int(input("Enter task number to delete: "))
    if task_number < 1 or task_number > len(todo_list):
        print("Invalid task number.")
    else:
        del todo_list[task_number - 1]
        print("Task deleted successfully!")

def import_tasks(todo_list):
    file_path = input("Enter the file path of the Excel file: ")
    try:
        df = pd.read_excel(file_path)
        tasks = df["Task"].tolist()
        for task in tasks:
            todo_list.append({"task": task, "completed": False})
        print("Tasks imported successfully!")
    except FileNotFoundError:
        print("File not found.")
    except Exception as e:
        print("Error occurred while importing tasks:", str(e))

def export_tasks(todo_list):
    print("Enter the file path to save the Excel file")
    print("ex) C:/Users/YourUsername/Documents/task_list.xlsx")
    file_path = input("Path: ")
    try:
        tasks = [task["task"] for task in todo_list]
        df = pd.DataFrame({"Task": tasks})
        df.to_excel(file_path, index=False)
        print("Tasks exported successfully!")
    except Exception as e:
        print("Error occurred while exporting tasks:", str(e))

def save_tasks(todo_list):
    with open("tasks.pickle", "wb") as file:
        pickle.dump(todo_list, file)
    print("Tasks saved successfully!")

def load_tasks():
    try:
        with open("tasks.pickle", "rb") as file:
            return pickle.load(file)
    except FileNotFoundError:
        return []

def main():
    todo_list = load_tasks()

    while True:
        display_menu()
        choice = input("Enter your choice: ")
        print("============================")
        if choice == "1":
            add_task(todo_list)
        elif choice == "2":
            view_tasks(todo_list)
        elif choice == "3":
            mark_task_completed(todo_list)
        elif choice == "4":
            delete_task(todo_list)
        elif choice == "5":
            import_tasks(todo_list)
        elif choice == "6":
            export_tasks(todo_list)
        elif choice == "7":
            save_tasks(todo_list)
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
