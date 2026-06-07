import tkinter as tk

from main_controller import MainController
from main_view import MainView
from models.main_model import MainModel


def main() -> None:
    root = tk.Tk()

    model = MainModel()
    view = MainView(root)
    MainController(model, view)

    root.mainloop()


if __name__ == "__main__":
    main()
