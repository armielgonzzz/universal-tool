import os
import customtkinter as ctk
import threading
import webbrowser
from customtkinter import filedialog
from tools.phone_cleanup_tool.clean_up import main as run_clean_up

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

def center_main_window(screen: ctk.CTkFrame, width: int, height: int) -> str:

    # Get width and height of main window
    screen_width = screen.winfo_screenwidth()
    screen_height = screen.winfo_screenheight()

    # Calculate position to center window
    x = int((screen_width/2) - (width/2))
    y = int((screen_height/2) - (height/1.5))

    return f"{width}x{height}+{x}+{y}"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ########################
        #                      #
        #   MAIN APP WINDOW    #
        #                      #
        ########################

        # Configure window
        self.title("Lead Management Tools")
        self.geometry(center_main_window(self, 800, 580))
        self.resizable(False, False)

        # Configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Left side frame
        self.left_side_frame = ctk.CTkFrame(self)
        self.left_side_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        self.left_side_frame.grid_columnconfigure(0, weight=1)
        self.left_side_frame.grid_rowconfigure(0, weight=1)

        # Tool Options Frame
        self.tool_options_frame = ctk.CTkScrollableFrame(self.left_side_frame)
        self.tool_options_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        self.tool_options_frame.grid_columnconfigure(0, weight=1)

        # Tool Window
        self.tool_window_frame = ctk.CTkFrame(self)
        self.tool_window_frame.grid_rowconfigure(0, weight=1)
        self.tool_window_frame.grid_columnconfigure(0, weight=1)
        self.tool_window_frame.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

        # Select from the tool buttons label
        self.tool_options_label = ctk.CTkLabel(self.tool_options_frame,
                                               text='Select a tool',
                                               font=ctk.CTkFont(
                                                   size=24,
                                                   weight='bold'
                                               ))
        self.tool_options_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        welcome_label = ctk.CTkLabel(self.tool_window_frame,
                                     text="Welcome to\nLead Management\nTools",
                                     font=ctk.CTkFont(
                                         size=36,
                                         weight='bold'))
        welcome_label.grid(row=0, column=0, padx=10, pady=10)

        self.documentation_button = ctk.CTkButton(self.left_side_frame,
                                                     text='Official Documentation',
                                                     command=self.open_link,
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343',
                                                     font=ctk.CTkFont(
                                                         weight='bold'
                                                     ))
        self.documentation_button.grid(row=1, column=0, padx=10, pady=(0,10), stick='nsew')

        self.clean_phone_tool_button = ctk.CTkButton(self.tool_options_frame,
                                                     text='Phone Number Cleanup Tool',
                                                     command=self.display_phone_clean_tool,
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343')
        self.clean_phone_tool_button.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')

    
    def open_link(self):
        webbrowser.open("https://github.com/armielgonzzz/universal-tool/blob/main/README.md")
    
    ##############################
    #                            #
    # PHONE NUMBER CLEAN UP TOOL #
    #                            #
    ##############################

    def display_phone_clean_tool(self):

        self.input_file_check = False
        self.save_path_check = False

        phone_clean_frame = ctk.CTkFrame(self.tool_window_frame)
        phone_clean_frame.grid_rowconfigure(7, weight=1)
        phone_clean_frame.grid_columnconfigure(0, weight=1)
        phone_clean_frame.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')

        label = ctk.CTkLabel(phone_clean_frame,
                             text="Phone Number Clean Up Tool",
                             font=ctk.CTkFont(
                                 size=36,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        select_file_button = ctk.CTkButton(phone_clean_frame,
                                           text='Select files to process',
                                           command=lambda: self.open_select_file(phone_clean_frame),
                                           fg_color='#5b5c5c',
                                           hover_color='#424343')
        select_file_button.grid(row=1, column=0, padx=10, pady=5)

        define_save_path_button = ctk.CTkButton(phone_clean_frame,
                                                text='Save output files to',
                                                command=lambda: self.select_save_directory(phone_clean_frame),
                                                fg_color='#5b5c5c',
                                                hover_color='#424343')
        define_save_path_button.grid(row=4, column=0, padx=10, pady=5)


    def open_select_file(self, frame: ctk.CTkFrame):

        self.file_paths = filedialog.askopenfilenames(title="Select a file",
                                               filetypes=[("CSV Files", "*.csv"),
                                                          ("Excel Files", "*.xlsx"),
                                                          ("All Files", "*.*")])
        if self.file_paths:

            files_to_process_label = ctk.CTkLabel(frame,
                                                  text='Files to Process',
                                                  font=ctk.CTkFont(
                                                      size=18,
                                                      weight='bold'
                                                  ))
            files_to_process_label.grid(row=2, column=0, padx=10, pady=5)

            select_files_frame = ctk.CTkScrollableFrame(frame,height=100)
            select_files_frame.grid_columnconfigure(0, weight=1)
            select_files_frame.grid(row=3, column=0, padx=30, pady=5, sticky='nsew')

            for index, file in enumerate(self.file_paths):
                file_name = os.path.basename(file)
                selected_files_label = ctk.CTkLabel(select_files_frame,
                                                text=file_name,
                                                wraplength=400)
                selected_files_label.grid(row=index, column=0, padx=10, pady=3, sticky='nsew')

            print(f"Files selected: {self.file_paths}")
            self.input_file_check = True

            if self.input_file_check and self.save_path_check:
                run_tool_button = ctk.CTkButton(frame,
                                                text='RUN TOOL',
                                                command=self.trigger_tool,
                                                height=36,
                                                width=240,
                                                fg_color='#d99125',
                                                hover_color='#ae741e',
                                                text_color='#141414',
                                                corner_radius=50,
                                                font=ctk.CTkFont(
                                                    size=18,
                                                    weight='bold'
                                                ))
                run_tool_button.grid(row=7, column=0, padx=10, pady=5)
    
    def select_save_directory(self, frame: ctk.CTkFrame):

        self.save_path = filedialog.askdirectory(title="Select save directory")
        if self.save_path:
            save_path_label = ctk.CTkLabel(frame,
                                       text=None,
                                       corner_radius=10,
                                       fg_color='#292929',
                                       wraplength=400)
            save_path_label.grid(row=6, column=0, padx=30, pady=5, sticky='nsew', ipadx=8, ipady=8)

            save_directory = ctk.CTkLabel(frame,
                                          text='Save Directory',
                                          font=ctk.CTkFont(
                                              size=18,
                                              weight='bold'
                                              ))
            save_directory.grid(row=5, column=0, padx=10, pady=5)
            save_path_label.configure(text=f"{self.save_path}")
            print(f"Directory selected: {self.save_path}")
            self.save_path_check = True

            if self.input_file_check and self.save_path_check:
                run_tool_button = ctk.CTkButton(frame,
                                                text='RUN TOOL',
                                                command=self.trigger_tool,
                                                height=36,
                                                width=240,
                                                fg_color='#d99125',
                                                hover_color='#ae741e',
                                                text_color='#141414',
                                                corner_radius=50,
                                                font=ctk.CTkFont(
                                                    size=18,
                                                    weight='bold'
                                                ))
                run_tool_button.grid(row=7, column=0, padx=10, pady=5)
    
    def tool_running_window(self):
        
        tool_run_window = ctk.CTkToplevel()
        center_new_window(self, tool_run_window)
        tool_run_window.resizable(False, False)
        tool_run_window.geometry("400x200")
        tool_run_window.grid_columnconfigure(0, weight=1)
        tool_run_window.grid_rowconfigure(0, weight=1)
        tool_run_window.attributes('-topmost', True)
        tool_run_window.title("Run Tool")

        tool_run_label = ctk.CTkLabel(tool_run_window,
                                      text="Tool is running in background.\nPlease wait for the tool to finish",
                                      font=ctk.CTkFont(
                                          size=16))
        tool_run_label.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        return tool_run_window
    
    def trigger_tool(self):
        window = self.tool_running_window()
        threading.Thread(target=self.run_clean_up_with_callback, args=(self.file_paths, self.save_path, window)).start()

    def run_clean_up_with_callback(self, file_paths, save_path, window):
        try:
            run_clean_up(file_paths, save_path)
            self.tool_result(result=True)

        except:
            self.tool_result(result=False)
        
        finally:
            window.destroy()

    def tool_result(self, result: bool=False):
        
        message: str = ""
        if result:
            message = "SUCCESSFULLY processed all files"
        else:
            message = f"Tool run FAILED"
        
        tool_result_window = ctk.CTkToplevel()
        center_new_window(self, tool_result_window)
        tool_result_window.resizable(False, False)
        tool_result_window.geometry("400x200")
        tool_result_window.grid_rowconfigure(0, weight=1)
        tool_result_window.grid_columnconfigure(0, weight=1)
        tool_result_window.attributes('-topmost', True)
        tool_result_window.title("Run Tool")
        tool_run_label = ctk.CTkLabel(tool_result_window,
                                      text=message,
                                      wraplength=300,
                                      font=ctk.CTkFont(
                                          size=14,
                                          weight='normal'))
        tool_run_label.grid(row=0, column=0, padx=10, pady=(15, 5), sticky="nsew")
        tool_run_button = ctk.CTkButton(tool_result_window,
                                        text="OK",
                                        fg_color='#5b5c5c',
                                        hover_color='#424343',
                                        command=lambda: tool_result_window.destroy())
        tool_run_button.grid(row=1, column=0, padx=10, pady=(5, 15))


def main() -> None:

    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
