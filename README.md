# To-Do List Manager

This is a simple To-Do List Manager implemented in Python. It allows you to add, view, mark as completed, delete tasks, import tasks from an Excel file, export tasks to an Excel file, and save tasks using pickle.

## Requirements
- Python 3.x

## Installation and Setup

1. Clone or download the project to your local machine.

2. Install the required external library by running the following command in your command prompt or terminal:

- pip install pandas

This command installs the pandas library, which is needed for importing and exporting tasks from/to Excel.

3. Once the installation is complete, you can run the script using the following command:

- python main.py

## Usage
Follow the on-screen menu to interact with the To-Do List Manager.

### Menu Options

1. **Add Task**: Adds a new task to the to-do list.
2. **View Tasks**: Displays all the tasks in the to-do list.
3. **Mark Task as Completed**: Marks a task as completed.
4. **Delete Task**: Deletes a task from the to-do list.
5. **Import Tasks from Excel**: Imports tasks from an Excel file and adds them to the to-do list.
6. **Export Tasks to Excel**: Exports the tasks in the to-do list to an Excel file.
7. **Quit**: Saves the tasks and exits the program.

## Data Persistence
The tasks are saved using the pickle module in a file named `tasks.pickle`. The tasks are loaded automatically when the program starts.

Please note that this implementation is meant to be a simple example and does not include extensive error handling or advanced features. Feel free to modify and enhance it according to your needs.

For any further assistance or questions, feel free to reach out.

You can copy and paste the above formatted text into your GitHub README file. Make sure to include it within the backticks (```) to maintain the markdown formatting.
