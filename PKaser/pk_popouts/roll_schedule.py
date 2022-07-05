from tkinter import messagebox


def roll_schedule():
    schedule = """
                Roll Schedule

    Quick Roll: 10:30am & 3:30pm (AST)
    Web Roll: 10am & 3pm (eOps)
    Biscuit Roll: Once a Month (eOps)
    """
    messagebox.showinfo(message='{}'.format(schedule))