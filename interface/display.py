import tkinter as tk
import customtkinter as ctk

# Outside function that will center a new pop up window relative to the main window
def center_new_window(main_window: ctk.CTkFrame,
                      new_window: ctk.CTkFrame) -> None:

    # Set the geomtry of the new window
    def set_geometry() -> None:

        # Get x and y position of the main window
        main_x = main_window.winfo_rootx()
        main_y = main_window.winfo_rooty()

        # Main window dimensions
        main_width = 800
        main_height = 580

        # Update idle task to avoid bugs in changing geometry
        new_window.update_idletasks()

        # Get the new window dimensions
        new_window_width = new_window.winfo_width()
        new_window_height = new_window.winfo_height()

        # Calculate new the position of the new window so that it is centered
        x = main_x + (main_width - new_window_width) // 2
        y = main_y + (main_height - new_window_height) // 2

        # Apply new position
        new_window.geometry(f"{new_window_width}x{new_window_height}+{int(x)}+{int(y)}")

    # Add delay to avoid bugs in setting new window position
    new_window.after(70, set_geometry)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        """"""""""""""""""""""""
        #   MAIN APP WINDOW    #
        """"""""""""""""""""""""

        # configure window
        self.title("Community Minerals Universal Tool")
        self.geometry(self.center_main_window(800, 580))
        self.resizable(False, False)

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

def main() -> None:

    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
