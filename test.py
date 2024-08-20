import sys

tasks = []

def menu():
    print("\nPICK ONE AND START")
    print("1. Add task")
    print("2. List task")
    print("3. Mark task")
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
    
def mark(tasks):
    list_task(tasks)
    task_number = int(input("Which number u want to mark: ")) - 1
    tasks[task_number]["completed"] = True
    print(f"'{tasks[task_number]['task']}' marked.")
    
add(tasks)
mark(tasks)
list_task(tasks)