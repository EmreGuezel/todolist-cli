import openpyxl

def save_to_excel(tasks, filename="tasks.xlsx"):
    # Create a workbook and a sheet
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Tasks"

    # Add headers
    sheet.append(["Task", "Status"])

    # Add tasks
    for task in tasks:
        status = "Completed" if task["completed"] else "Continue..."
        sheet.append([task["task"], status])

    # Save the workbook
    wb.save(filename)
    print(f"Tasks saved to {filename}")

tasks = []

def menu():
    print("\nPICK ONE AND START")
    print("1. Add task")
    print("2. List task")
    print("3. Complete task")
    print("4. Delete task")
    print("5. EXIT\n")
    
def add(tasks):
    task = input("ADD YOUR TASK: ")
    tasks.append({"task": task, "completed" : False})
    print(f"{task} added.")

def list_task(tasks):
    if not tasks:
        print("There is no task to do.")
    else:
        for idx, task in enumerate(tasks, 1):
            status = "Completed" if task["completed"] else "Continue..."
            print(f"{idx}. {task['task']} [{status}]")
            
def delete(tasks):
    list_task(tasks)
    task_number = int(input("Which number u want to delete: ")) - 1
    deleted_task = tasks.pop(task_number)
    print(f"'{deleted_task['task']}' deleted.")
    
def complete(tasks):
    list_task(tasks)
    task_number = int(input("Which number u want to mark: ")) - 1
    tasks[task_number]["completed"] = True
    print(f"'{tasks[task_number]['task']}' marked.")
    
def main():
    tasks = []
    while True:
        menu()
        choice = input("Choose one: ")

        match choice:
            case "1":
                add(tasks)
            case "2":
                list_task(tasks)
            case "3":
                complete(tasks)
            case "4":
                delete(tasks)
            case "5":
                print("Exiting...")
                break
            case _:
                print("invalid")

if __name__ == "__main__":
    main()