import openpyxl
import win32com.client as win32
import os
import logging

# Configure logging
logging.basicConfig(filename='tasks.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class TaskManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.tasks = self.load_from_excel()

    def save_to_excel(self):
        try:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title = "Tasks"
            sheet.append(["Task", "Status"])

            for task in self.tasks:
                status = "Completed" if task["completed"] else "Continue..."
                sheet.append([task["task"], status])

            wb.save(self.filepath)
            logging.info(f"Tasks saved to {self.filepath}")

            # Refresh Excel file
            excel = win32.Dispatch("Excel.Application")
            workbook = excel.Workbooks.Open(self.filepath)
            excel.CalculateFull()
            workbook.Save()
            workbook.Close()
            excel.Quit()
        except Exception as e:
            logging.error(f"Failed to save tasks: {e}")

    def load_from_excel(self):
        tasks = []
        if os.path.exists(self.filepath):
            try:
                wb = openpyxl.load_workbook(self.filepath)
                sheet = wb.active

                for row in sheet.iter_rows(min_row=2, values_only=True):
                    task, status = row
                    tasks.append({"task": task, "completed": status == "Completed"})

                logging.info(f"Tasks loaded from {self.filepath}")
            except Exception as e:
                logging.error(f"Failed to load tasks: {e}")
        else:
            logging.info(f"{self.filepath} not found. Starting with an empty task list.")
        return tasks

    def add_task(self, task):
        if task.strip() == "":
            print("Task cannot be empty.")
            logging.warning("Attempted to add an empty task.")
            return
        self.tasks.append({"task": task, "completed": False})
        self.save_to_excel()
        print(f"'{task}' added.")
        logging.info(f"Task added: {task}")

    def list_tasks(self):
        if not self.tasks:
            print("There is no task to do.")
        else:
            for idx, task in enumerate(self.tasks, 1):
                status = "Completed" if task["completed"] else "Continue..."
                print(f"{idx}. {task['task']} [{status}]")


    def delete_task(self, task_number):
        if 0 <= task_number < len(self.tasks):
            deleted_task = self.tasks.pop(task_number)
            self.save_to_excel()
            print(f"'{deleted_task['task']}' deleted.")
            logging.info(f"Task deleted: {deleted_task['task']}")
        else:
            print("Invalid task number.")
            logging.warning("Invalid task number for deletion.")

    def complete_task(self, task_number):
        if 0 <= task_number < len(self.tasks):
            self.tasks[task_number]["completed"] = True
            self.save_to_excel()
            print(f"'{self.tasks[task_number]['task']}' marked as completed.")
            logging.info(f"Task completed: {self.tasks[task_number]['task']}")
        else:
            print("Invalid task number.")
            logging.warning("Invalid task number for completion.")

def menu():
    print("\nPICK ONE AND START")
    print("1. Add task")
    print("2. List tasks")
    print("3. Complete task")
    print("4. Delete task")
    print("5. EXIT\n")

def main():
    filepath = r"C:\Users\emreg\OneDrive\Masaüstü\deneme\todolist-cli\tasks.xlsx"
    task_manager = TaskManager(filepath)

    while True:
        menu()
        choice = input("Choose one: ")

        if choice == "1":
            task = input("ADD YOUR TASK: ")
            task_manager.add_task(task)
        elif choice == "2":
            task_manager.list_tasks()
        elif choice == "3":
            task_manager.list_tasks()
            try:
                task_number = int(input("Which number do you want to mark as completed: ")) - 1
                task_manager.complete_task(task_number)
            except ValueError:
                print("Invalid input. Please enter a number.")
                logging.warning("Invalid input for completing a task.")
        elif choice == "4":
            task_manager.list_tasks()
            try:
                task_number = int(input("Which number do you want to delete: ")) - 1
                task_manager.delete_task(task_number)
            except ValueError:
                print("Invalid input. Please enter a number.")
                logging.warning("Invalid input for deleting a task.")
        elif choice == "5":
            task_manager.save_to_excel()
            print("Exiting...")
            logging.info("Exiting application.")
            break
        else:
            print("Invalid choice. Please try again.")
            logging.warning("Invalid menu choice.")

if __name__ == "__main__":
    main()
