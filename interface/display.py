import os
import customtkinter as ctk
import threading
import webbrowser
from customtkinter import filedialog
from tools.text_inactive_tool.text_inactive import main as run_text_inactive
from tools.phone_cleanup_tool.clean_up import main as run_clean_up
from tools.pipedrive_automation_tool.pipedrive_automation import main as run_automation
from tools.autodialer_cleanup_tool.cleanup_autodialer import main as run_autodialer
from tools.missing_deals_tool.missing_deals import main as run_missing_deals
from tools.missing_deals_tool.lookup import main as run_missing_deals_lookup
from tools.marketing_cleanup_tool.marketing_clean_up import main as run_marketing_cleanup
from tools.autodialer_cleanup_tool.cleaner_file_automation import dropbox_authentication
from tools.autodialer_cleanup_tool.cleaner_file_automation import main as update_list_cleaner_file


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
        self.clean_phone_tool_button.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')
        self.clean_phone_tool_button.bind("<Button-1>", lambda event: self.track_button_click(1))

        self.text_inactive_button = ctk.CTkButton(self.tool_options_frame,
                                                     text='Text Inactive Tool',
                                                     command=lambda:self.show_frame(TextInactive),
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343')
        self.text_inactive_button.grid(row=2, column=0, padx=10, pady=5, sticky='nsew')
        self.text_inactive_button.bind("<Button-1>", lambda event: self.track_button_click(2))

        self.pipedrive_automation_button = ctk.CTkButton(self.tool_options_frame,
                                                     text='Pipedrive Automation Tool',
                                                     command=lambda:self.show_frame(PipedriveAutomation),
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343')
        self.pipedrive_automation_button.grid(row=3, column=0, padx=10, pady=5, sticky='nsew')
        self.pipedrive_automation_button.bind("<Button-1>", lambda event: self.track_button_click(3))

        self.autodialer_button = ctk.CTkButton(self.tool_options_frame,
                                                     text='AutoDialer Cleanup Tool',
                                                     command=lambda:self.show_frame(AutoDialerCleaner),
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343')
        self.autodialer_button.grid(row=4, column=0, padx=10, pady=5, sticky='nsew')
        self.autodialer_button.bind("<Button-1>", lambda event: self.track_button_click(4))

        self.missing_deals_button = ctk.CTkButton(self.tool_options_frame,
                                                     text='Missing Deals Text Tool',
                                                     command=lambda:self.show_frame(MissingDealsText),
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343')
        self.missing_deals_button.grid(row=5, column=0, padx=10, pady=5, sticky='nsew')
        self.missing_deals_button.bind("<Button-1>", lambda event: self.track_button_click(5))

        self.marketing_cleanup_button = ctk.CTkButton(self.tool_options_frame,
                                                     text='Marketing Cleanup Tool',
                                                     command=lambda:self.show_frame(MarketingCleanupTool),
                                                     fg_color='#5b5c5c',
                                                     hover_color='#424343')
        self.marketing_cleanup_button.grid(row=6, column=0, padx=10, pady=5, sticky='nsew')
        self.marketing_cleanup_button.bind("<Button-1>", lambda event: self.track_button_click(6))

        self.clicked_button_id = ctk.IntVar()
        self.current_frame = None
        self.input_file_check, self.save_path_check = False, False
        self.show_frame(InitialFrame)

    ##################
    #                #
    # HELPER METHODS #
    #                #
    ##################

    def display_checklist(self, window):
        
        # Function to check if all checkboxes are checked
        def check_all_selected():
            if all(state.get() for state in checkbox_states):
                confirm_button.configure(state="normal", fg_color='#5b5c5c')
            else:
                confirm_button.configure(state="disabled", fg_color='#424343')

        window.destroy()

        checklist_window = ctk.CTkToplevel()
        center_new_window(self, checklist_window)
        checklist_window.resizable(False, False)
        checklist_window.geometry("470x400")
        checklist_window.grid_columnconfigure(0, weight=1)
        checklist_window.grid_rowconfigure(0, weight=1)
        checklist_window.attributes('-topmost', True)
        checklist_window.protocol("WM_DELETE_WINDOW", lambda: None)
        checklist_window.title("Output Checklist")
        checklist_window.grid_rowconfigure(1, weight=1)
        checklist_window.grid_columnconfigure(0, weight=1)
        checklist_scrollable_frame = ctk.CTkScrollableFrame(checklist_window,
                                                            corner_radius=20)
        checklist_scrollable_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew", ipadx=5, ipady=5)

        checkbox_states = []
        output_checklist_dict = {
            1 : ["Correct output column names",
                 "No Duplicates in phone_number column",
                 "Correctness of tagging per lead from\ncolumn reason_for_removal",
                 "If tagged 'in_pipedrive is Y', verify if value in column\nin_pipedrive is Y from input file",
                 "If tagged 'rc_pd is Yes', verify if value in column\n rc_pd is Yes from input file",
                 "If tagged 'Both type & carrier_type are Landline', verify\nif both columns type & carrier_type are Landline from input file",
                 "If tagged 'text_opt_in is No', verify if column text_opt_in\nis No from input file",
                 "If tagged 'contact_deal_id Not Empty', verify if column\ncontact_deal_id is empty from input file",
                 "If tagged 'contact_deal_status Not Empty', verify if column\ncontact_deal_status is empty from input file",
                 "If tagged 'contact_person_id Not Empty', verify if column\ncontact_person_id is empty from input file",
                 "If tagged 'phone_number_deal_id Not Empty', verify if column\nphone_number_deal_id is empty from input file",
                 "If tagged 'phone_number_deal_status Not Empty', verify if\ncolumn phone_number_deal_status is empty from input file",
                 "If tagged 'RVM - Last RVM Date - last 7 days from tool run\ndate', verify if the date from column RVM - Last RVM Date is\n7 days ago from today from input file",
                 "If tagged 'Latest Text Marketing Date (Sent) - last 7 days\nfrom tool run date', verify if the date from column Latest\nText Marketing Date (Sent) is 7 days ago from today from input file",
                 "If tagged 'Rolling 30 Days Rvm Count and Rolling 30 Days\nText Marketing Count - total >= 3', verify if the total count\nof columns Rolling 30 Days Rvm Count and Rolling 30 Days Text\nMarketing Count is less than or equal to 3 from input file",
                 "If tagged 'Deal - ID Not Empty', verify if column\nDeal - ID is empty from input file",
                 "If tagged 'Deal - Text Opt-in is No', verify if column\nDeal - Text Opt-in is No from input file",
                 "For leads with empty reason_for_removal values, verify\nif all of the previous column conditions are not satisfied",
                 'Multiple tagging for reason_for_removal\nwith separator of " , "',
                 "Output file name and format"
                 ],
            
            2: [
                "Correct output column names",
                "Two output files if tool created new deals\n(FU and New Deals)",
                "No blank values for column Deal ID in FU file",
                'Multiple Deal IDs in FU file with separator of " | "',
                '"Text Inactive - For Review" value in Deal - Label column',
                "Proper row value and format in columns Deal - Title,\nDeal - Deal Summary, Person - Name and Note (if any)",
                "Person - Phone, Person - Phone 1 not blank",
                "Importing of output files to Pipedrive is successful"
                ],
            
            3: [
                "Correct output column names",
                "No duplicate values from column Deal - ID",
                "Only unique serial numbers from column Deal - Serial Number",
                "Deal - Deal Summary value should be 'Completed'",
                "Proper formatting of values from column\nPerson - Mailing Address (should have '..., USA' at the end)",
                "For non-empty values from columns 'Person - Email' and\n'Person - Phone', verify if the same values are reflected for\ncolumns 'Person - Email 1' and 'Person - Phone 1' respectively"
                ],
            
            4: [
                "AutoDialer List"
               ],

            5: [
                "Missing Deals Text List"
               ],
            
            6: [
                "Marketing Cleanup Tool List"
               ]
        }
        checkbox_labels = output_checklist_dict[self.clicked_button_id.get()]

        checklist_label = ctk.CTkLabel(checklist_window,
                                       text="Verify the following from the output file(s)",
                                       font=ctk.CTkFont(
                                           size=18,
                                           weight='bold'
                                       ))
        checklist_label.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        # Create checkboxes for each item in the list
        for label in checkbox_labels:
            state = ctk.BooleanVar()
            checkbox = ctk.CTkCheckBox(checklist_scrollable_frame,
                                       text=label,
                                       variable=state,
                                       fg_color='green',
                                       hover_color='green',
                                       command=check_all_selected)
            checkbox.pack(pady=5, anchor='w')
            checkbox_states.append(state)
        
        confirm_button = ctk.CTkButton(checklist_window,
                                       text="Confirm",
                                       state="disabled",
                                       fg_color='#424343',
                                       hover_color='#424343',
                                       command=checklist_window.destroy)
        confirm_button.grid(row=2, column=0, padx=5, pady=(5,10))

    def track_button_click(self, button_id):
        self.clicked_button_id.set(button_id)

    def show_frame(self, frame_class):
        """Destroys the current frame and replaces it with a new one."""
        if self.current_frame is not None:
            self.current_frame.destroy()
        self.current_frame = frame_class(self.tool_window_frame, self)

    
    def open_link(self):
        webbrowser.open("https://github.com/armielgonzzz/universal-tool/blob/main/README.md")
    
    ##############################
    #                            #
    # PHONE NUMBER CLEAN UP TOOL #
    #                            #
    ##############################

    def display_phone_clean_tool(self):

        self.input_file_check, self.save_path_check, self.cleaner_file_check = False, False, False

        self.current_frame = ctk.CTkFrame(self.tool_window_frame)
        self.current_frame.grid_rowconfigure(7, weight=1)
        self.current_frame.grid_columnconfigure(0, weight=1)
        self.current_frame.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')

        label = ctk.CTkLabel(self.current_frame,
                             text="Phone Number Clean Up Tool",
                             font=ctk.CTkFont(
                                 size=30,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        update_cleaner_button = ctk.CTkButton(self.current_frame,
                                              text="Select list cleaner file",
                                              fg_color='#5b5c5c',
                                              hover_color='#424343',
                                              command=lambda:self.select_list_cleaner_file(frame=self.current_frame, func=run_clean_up))
        update_cleaner_button.grid(row=1, column=0, padx=5, pady=5, sticky="ns")
        
        select_file_button = ctk.CTkButton(self.current_frame,
                                           text='Select files to process',
                                           command=lambda: self.open_select_file(frame=self.current_frame, func=run_clean_up),
                                           fg_color='#5b5c5c',
                                           hover_color='#424343')
        select_file_button.grid(row=3, column=0, padx=10, pady=5)

        define_save_path_button = ctk.CTkButton(self.current_frame,
                                                text='Save output files to',
                                                command=lambda: self.select_save_directory(frame=self.current_frame, func=run_clean_up),
                                                fg_color='#5b5c5c',
                                                hover_color='#424343')
        define_save_path_button.grid(row=5, column=0, padx=10, pady=5)


    def open_select_file(self, frame: ctk.CTkFrame, func):

        self.file_paths = filedialog.askopenfilenames(title="Select a file",
                                               filetypes=[("All Files", "*.*")])
        if self.file_paths:

            select_files_frame = ctk.CTkScrollableFrame(frame,height=100)
            select_files_frame.grid_columnconfigure(0, weight=1)
            select_files_frame.grid(row=4, column=0, padx=30, pady=5, sticky='nsew')

            for index, file in enumerate(self.file_paths):
                file_name = os.path.basename(file)
                selected_files_label = ctk.CTkLabel(select_files_frame,
                                                text=file_name,
                                                wraplength=400)
                selected_files_label.grid(row=index, column=0, padx=10, pady=3, sticky='nsew')

            print(f"Files selected: {self.file_paths}")
            self.input_file_check = True

            if self.input_file_check and self.save_path_check and self.cleaner_file:
                run_tool_button = ctk.CTkButton(frame,
                                                text='RUN TOOL',
                                                command=lambda: self.trigger_tool(func, self.cleaner_file, self.file_paths, self.save_path),
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
    
    def select_save_directory(self, frame: ctk.CTkFrame, func):

        self.save_path = filedialog.askdirectory(title="Select save directory")
        if self.save_path:
            save_path_label = ctk.CTkLabel(frame,
                                       text=None,
                                       corner_radius=10,
                                       fg_color='#292929',
                                       wraplength=400)
            save_path_label.grid(row=6, column=0, padx=30, pady=5, sticky='nsew', ipadx=8, ipady=8)

            save_path_label.configure(text=f"{self.save_path}")
            print(f"Directory selected: {self.save_path}")
            self.save_path_check = True

            if self.input_file_check and self.save_path_check and self.cleaner_file:
                run_tool_button = ctk.CTkButton(frame,
                                                text='RUN TOOL',
                                                command=lambda: self.trigger_tool(func, self.cleaner_file, self.file_paths, self.save_path),
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
    
    def select_list_cleaner_file(self, frame: ctk.CTkFrame, func):
        self.cleaner_file = filedialog.askopenfilename(title="Select list cleaner file",
                                                        filetypes=[("All Files", "*.*")])
        if self.cleaner_file:
            cleaner_file_label = ctk.CTkLabel(frame,
                                              text=f"{os.path.basename(self.cleaner_file)}",
                                              fg_color="transparent")
            cleaner_file_label.grid(row=2, column=0, padx=5, pady=5)
            self.cleaner_file_check = True
            if self.input_file_check and self.save_path_check and self.cleaner_file_check:
                run_tool_button = ctk.CTkButton(frame,
                                                text='RUN TOOL',
                                                command=lambda: self.trigger_tool(func, self.cleaner_file, self.file_paths, self.save_path),
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
    
    def trigger_tool(self, func, *args, **kwargs):
        window = self.tool_running_window()
        threading.Thread(target=self.run_clean_up_with_callback, args=(window, func, *args), kwargs=kwargs).start()

    def run_clean_up_with_callback(self, window, func, *args, **kwargs):
        try:
            func(*args, **kwargs)
            self.tool_result(result=True)

        except:
            self.tool_result(result=False)
        
        finally:
            window.destroy()

    def tool_result(self, result: bool=False):

        tool_result_window = ctk.CTkToplevel()
        center_new_window(self, tool_result_window)
        tool_result_window.resizable(False, False)
        tool_result_window.geometry("400x200")
        tool_result_window.grid_rowconfigure(0, weight=1)
        tool_result_window.grid_columnconfigure(0, weight=1)
        tool_result_window.attributes('-topmost', True)
        tool_result_window.title("Run Tool")

        tool_run_label = ctk.CTkLabel(tool_result_window,
                                      text="SUCCESSFULLY processed all files" if result else "Tool run FAILED",
                                      wraplength=300,
                                      font=ctk.CTkFont(
                                          size=14,
                                          weight='normal'))
        tool_run_label.grid(row=0, column=0, padx=10, pady=(15, 5), sticky="nsew")
        tool_run_button = ctk.CTkButton(tool_result_window,
                                        text="OK",
                                        fg_color='#5b5c5c',
                                        hover_color='#424343',
                                        command=lambda:
                                        self.display_checklist(tool_result_window) if result else tool_result_window.destroy())
        tool_run_button.grid(row=1, column=0, padx=10, pady=(5, 15))

class InitialFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        welcome_label = ctk.CTkLabel(self,
                                     text="Welcome to\nLead Management\nTools",
                                     font=ctk.CTkFont(
                                         size=36,
                                         weight='bold'))
        welcome_label.grid(row=0, column=0, padx=10, pady=10)


class TextInactive(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.controller.input_file_check = False
        self.controller.save_path_check = False
        self.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        self.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(self,
                             text="Text Inactive Tool",
                             font=ctk.CTkFont(
                                 size=30,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        select_file_button = ctk.CTkButton(self,
                                           text='Select files to process',
                                           command=lambda: self.controller.open_select_file(frame=self,
                                                                                            func=run_text_inactive),
                                           fg_color='#5b5c5c',
                                           hover_color='#424343')
        select_file_button.grid(row=1, column=0, padx=10, pady=5)

        define_save_path_button = ctk.CTkButton(self,
                                                text='Save output files to',
                                                command=lambda: self.controller.select_save_directory(frame=self,
                                                                                                      func=run_text_inactive),
                                                fg_color='#5b5c5c',
                                                hover_color='#424343')
        define_save_path_button.grid(row=4, column=0, padx=10, pady=5)

class PipedriveAutomation(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.controller.input_file_check = False
        self.controller.save_path_check = False
        self.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        self.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(self,
                             text="Pipedrive Automation Tool",
                             font=ctk.CTkFont(
                                 size=30,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        select_file_button = ctk.CTkButton(self,
                                           text='Select files to process',
                                           command=lambda: self.controller.open_select_file(frame=self,
                                                                                            func=run_automation),
                                           fg_color='#5b5c5c',
                                           hover_color='#424343')
        select_file_button.grid(row=1, column=0, padx=10, pady=5)

        define_save_path_button = ctk.CTkButton(self,
                                                text='Save output files to',
                                                command=lambda: self.controller.select_save_directory(frame=self,
                                                                                                      func=run_automation),
                                                fg_color='#5b5c5c',
                                                hover_color='#424343')
        define_save_path_button.grid(row=4, column=0, padx=10, pady=5)

class AutoDialerCleaner(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.controller.input_file_check = False
        self.controller.save_path_check = False
        self.cleaner_file = False
        self.files_to_clean = False
        self.save_path = False
        self.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(8, weight=1)

        label = ctk.CTkLabel(self,
                             text="AutoDialer Cleanup Tool",
                             font=ctk.CTkFont(
                                 size=30,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=(10,0), sticky='nsew')

        update_cleaner_button = ctk.CTkButton(self,
                                            text="Update list cleaner file",
                                            fg_color='#5b5c5c',
                                            hover_color='#424343',
                                            command=lambda:self.update_list_cleaner())
        update_cleaner_button.grid(row=1, column=0, padx=5, pady=5, sticky="ns")

        update_cleaner_button = ctk.CTkButton(self,
                                              text="Select list cleaner file",
                                              fg_color='#5b5c5c',
                                              hover_color='#424343',
                                              command=lambda:self.select_cleaner_file(self))
        update_cleaner_button.grid(row=2, column=0, padx=5, pady=5, sticky="ns")

        list_button = ctk.CTkButton(self,
                                    text="Select files to clean",
                                    fg_color='#5b5c5c',
                                    hover_color='#424343',
                                    command=lambda:self.select_files_to_clean(self))
        list_button.grid(row=4, column=0, padx=5, pady=5, sticky="ns")

        list_button = ctk.CTkButton(self,
                                    text="Save output files to",
                                    fg_color='#5b5c5c',
                                    hover_color='#424343',
                                    command=lambda:self.select_save_path(self))
        list_button.grid(row=6, column=0, padx=5, pady=5, sticky="ns")
    
    def select_cleaner_file(self, window):

        self.cleaner_file = filedialog.askopenfilename(title="Select list cleaner file",
                                                        filetypes=[("All Files", "*.*")])
        if self.cleaner_file:
            cleaner_file_label = ctk.CTkLabel(window,
                                              text=f"{os.path.basename(self.cleaner_file)}",
                                              fg_color="transparent")
            cleaner_file_label.grid(row=3, column=0, padx=5, pady=5)
            self.check_run()

    def select_files_to_clean(self, window):

        self.files_to_clean = filedialog.askopenfilenames(title="Select files to clean",
                                                          filetypes=[("All Files", "*.*")])
        if self.files_to_clean:
            files_to_clean_frame = ctk.CTkScrollableFrame(window,
                                                          height=80)
            files_to_clean_frame.grid(row=5, column=0, padx=5, pady=5, sticky="ew")
            files_to_clean_frame.grid_columnconfigure(0, weight=1)

            for i, file in enumerate(self.files_to_clean):
                file_name = os.path.basename(file)
                selected_files_label = ctk.CTkLabel(files_to_clean_frame,
                                                    text=file_name,
                                                    wraplength=400)
                selected_files_label.grid(row=i, column=0, padx=10, pady=3, sticky='nsew')       
            self.check_run()         
    
    def select_save_path(self, window):
        self.save_path = filedialog.askdirectory(title="Save directory")
        if self.save_path:
            save_path_label = ctk.CTkLabel(window,
                                           text=f"{self.save_path}",
                                           fg_color="transparent",
                                           wraplength=400)
            save_path_label.grid(row=7, column=0, padx=5, pady=5)
            self.check_run()

    def check_run(self):
        if self.cleaner_file and self.files_to_clean and self.save_path:
            run_tool_button = ctk.CTkButton(self,
                                            text='RUN TOOL',
                                            height=36,
                                            width=240,
                                            fg_color='#d99125',
                                            hover_color='#ae741e',
                                            text_color='#141414',
                                            corner_radius=50,
                                            font=ctk.CTkFont(size=18, weight='bold'),
                                            command=lambda:self.controller.trigger_tool(run_autodialer,
                                                                                        self.cleaner_file,
                                                                                        self.files_to_clean,
                                                                                        self.save_path))
            run_tool_button.grid(row=8, column=0, padx=10, pady=5)

    def update_list_cleaner(self):

        def submit_action(window):
            user_input = input_field.get()
            if user_input:
                window.destroy()
                self.controller.trigger_tool(update_list_cleaner_file, user_input)
                

        # Create the authentication frame (a new top-level window)
        authentication_frame = ctk.CTkToplevel(self)
        authentication_frame.resizable(False, False)
        authentication_frame.geometry("470x180")
        authentication_frame.grid_columnconfigure(0, weight=1)
        authentication_frame.grid_rowconfigure((2), weight=1)  # Adjust grid row configuration
        authentication_frame.attributes('-topmost', True)
        authentication_frame.title("Dropbox Authentication")

        # Button to open authentication link
        open_link = ctk.CTkButton(authentication_frame,
                                text="Open Authentication Link",
                                fg_color='#5b5c5c',
                                hover_color='#424343',
                                command=lambda: dropbox_authentication())
        open_link.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        input_field = ctk.CTkEntry(authentication_frame, placeholder_text="Enter the value here")
        input_field.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')

        submit_button = ctk.CTkButton(authentication_frame,
                                      text="Update List Cleaner",
                                      fg_color='#d99125',
                                      hover_color='#ae741e',
                                      text_color='#141414',
                                      corner_radius=50,
                                      font=ctk.CTkFont(size=18, weight='bold'),
                                      command=lambda: submit_action(authentication_frame))
        submit_button.grid(row=2, column=0, padx=10, pady=20, sticky='nsew')



class MissingDealsText(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.controller.input_file_check = False
        self.controller.save_path_check = False
        self.files_to_process = False
        self.save_path = False
        self.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((5), weight=1)

        label = ctk.CTkLabel(self,
                             text="Missing Deals Text Tool",
                             font=ctk.CTkFont(
                                 size=30,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        select_file_button = ctk.CTkButton(self,
                                           text='Select files to process',
                                           command=lambda: self.select_files_to_clean(self),
                                           fg_color='#5b5c5c',
                                           hover_color='#424343')
        select_file_button.grid(row=1, column=0, padx=10, pady=5)

        define_save_path_button = ctk.CTkButton(self,
                                                text='Save output files to',
                                                command=lambda: self.select_save_path(self),
                                                fg_color='#5b5c5c',
                                                hover_color='#424343')
        define_save_path_button.grid(row=3, column=0, padx=10, pady=5)

    def select_files_to_clean(self, window):

        self.files_to_process = filedialog.askopenfilenames(title="Select files to clean",
                                                          filetypes=[("All Files", "*.*")])
        if self.files_to_process:
            files_to_clean_frame = ctk.CTkScrollableFrame(window,
                                                          height=80)
            files_to_clean_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
            files_to_clean_frame.grid_columnconfigure(0, weight=1)

            for i, file in enumerate(self.files_to_process):
                file_name = os.path.basename(file)
                selected_files_label = ctk.CTkLabel(files_to_clean_frame,
                                                    text=file_name,
                                                    wraplength=400)
                selected_files_label.grid(row=i, column=0, padx=10, pady=3, sticky='nsew')       
            self.check_run()         
    
    def select_save_path(self, window):
        self.save_path = filedialog.askdirectory(title="Save directory")
        if self.save_path:
            save_path_label = ctk.CTkLabel(window,
                                           text=f"{self.save_path}",
                                           fg_color="transparent",
                                           wraplength=400)
            save_path_label.grid(row=4, column=0, padx=5, pady=5)
            self.check_run()

    
    def check_run(self):
        if self.files_to_process and self.save_path:
            run_buttons_frame = ctk.CTkFrame(self,
                                             fg_color='transparent')
            run_buttons_frame.grid(row=5, column=0, padx=10, pady=5)
            run_buttons_frame.grid_rowconfigure(0, weight=1)
            run_buttons_frame.grid_columnconfigure((0,1), weight=1)

            run_tool_button = ctk.CTkButton(run_buttons_frame,
                                            text='RUN TOOL',
                                            height=36,
                                            width=240,
                                            fg_color='#d99125',
                                            hover_color='#ae741e',
                                            text_color='#141414',
                                            corner_radius=50,
                                            font=ctk.CTkFont(size=18, weight='bold'),
                                            command=lambda:self.controller.trigger_tool(run_missing_deals,
                                                                                        self.files_to_process,
                                                                                        self.save_path))
            run_tool_button.grid(row=0, column=0, padx=10, pady=5)

            run_tool_button = ctk.CTkButton(run_buttons_frame,
                                            text='RUN LOOK UP',
                                            height=36,
                                            width=240,
                                            # fg_color='#d99125',
                                            # hover_color='#ae741e',
                                            # text_color='#141414',
                                            corner_radius=50,
                                            font=ctk.CTkFont(size=18, weight='bold'),
                                            command=lambda:self.controller.trigger_tool(run_missing_deals_lookup,
                                                                                        self.files_to_process,
                                                                                        self.save_path))
            run_tool_button.grid(row=0, column=1, padx=10, pady=5)

class MarketingCleanupTool(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.controller.input_file_check = False
        self.controller.save_path_check = False
        self.pipedrive_file = False
        self.files_to_clean = False
        self.save_path = False
        self.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(7, weight=1)

        label = ctk.CTkLabel(self,
                             text="Marketing Cleanup Tool",
                             font=ctk.CTkFont(
                                 size=30,
                                 weight='bold'
                             ))
        label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        list_cleaner_button = ctk.CTkButton(self,
                                            text="Select pipedrive file",
                                            fg_color='#5b5c5c',
                                            hover_color='#424343',
                                            command=lambda:self.select_cleaner_file(self))
        list_cleaner_button.grid(row=1, column=0, padx=5, pady=5, sticky="ns")

        list_button = ctk.CTkButton(self,
                                    text="Select files to clean",
                                    fg_color='#5b5c5c',
                                    hover_color='#424343',
                                    command=lambda:self.select_files_to_clean(self))
        list_button.grid(row=3, column=0, padx=5, pady=5, sticky="ns")

        list_button = ctk.CTkButton(self,
                                    text="Save output files to",
                                    fg_color='#5b5c5c',
                                    hover_color='#424343',
                                    command=lambda:self.select_save_path(self))
        list_button.grid(row=5, column=0, padx=5, pady=5, sticky="ns")
    
    def select_cleaner_file(self, window):

        self.pipedrive_file = filedialog.askopenfilename(title="Select list cleaner file",
                                                        filetypes=[("All Files", "*.*")])
        if self.pipedrive_file:
            cleaner_file_label = ctk.CTkLabel(window,
                                              text=f"{os.path.basename(self.pipedrive_file)}",
                                              fg_color="transparent")
            cleaner_file_label.grid(row=2, column=0, padx=5, pady=5)
            self.check_run()

    def select_files_to_clean(self, window):

        self.files_to_clean = filedialog.askopenfilenames(title="Select files to clean",
                                                          filetypes=[("All Files", "*.*")])
        if self.files_to_clean:
            files_to_clean_frame = ctk.CTkScrollableFrame(window,
                                                          height=80)
            files_to_clean_frame.grid(row=4, column=0, padx=5, pady=5, sticky="ew")
            files_to_clean_frame.grid_columnconfigure(0, weight=1)

            for i, file in enumerate(self.files_to_clean):
                file_name = os.path.basename(file)
                selected_files_label = ctk.CTkLabel(files_to_clean_frame,
                                                    text=file_name,
                                                    wraplength=400)
                selected_files_label.grid(row=i, column=0, padx=10, pady=3, sticky='nsew')       
            self.check_run()         
    
    def select_save_path(self, window):
        self.save_path = filedialog.askdirectory(title="Save directory")
        if self.save_path:
            save_path_label = ctk.CTkLabel(window,
                                           text=f"{self.save_path}",
                                           fg_color="transparent",
                                           wraplength=400)
            save_path_label.grid(row=6, column=0, padx=5, pady=5)
            self.check_run()

    def check_run(self):
        if self.pipedrive_file and self.files_to_clean and self.save_path:
            run_tool_button = ctk.CTkButton(self,
                                            text='RUN TOOL',
                                            height=36,
                                            width=240,
                                            fg_color='#d99125',
                                            hover_color='#ae741e',
                                            text_color='#141414',
                                            corner_radius=50,
                                            font=ctk.CTkFont(size=18, weight='bold'),
                                            command=lambda:self.controller.trigger_tool(run_marketing_cleanup,
                                                                                        self.files_to_clean,
                                                                                        self.pipedrive_file,
                                                                                        self.save_path))
            run_tool_button.grid(row=7, column=0, padx=10, pady=5)


def main() -> None:

    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
