"""
Application Switcher for Windows
================================

Keyboard-based application switcher for Windows using window title matching.

Author:
    
    Corvin Gröning

GitHub:

    https://github.com/cgroening/WinAppSwitcher

This module provides a command-line interface to quickly switch between open
Windows applications by pressing predefined keyboard shortcuts. It uses the
Windows API via the `pywin32` library to list, detect, and activate open
windows based on partial title matching.

The tool displays an interactive table in the terminal showing which keys are
bound to which applications. When the user presses a valid key, the
corresponding application window is brought to the foreground (restored from
minimized if needed).

Designed for productivity workflows, the AppSwitcher can be customized easily
via the `KEY_BINDINGS` dictionary.

The script end if any other key than a letter is typed.

It's recommended to leave the command window running this script open and use
AutoHotKey to bring it to the front.

Example for an AutoHotKey script (bound to ALT+Space):

```
#SingleInstance force
!Space::
    if WinExist("Python App Switcher")
        WinActivate
return
```

"""

import win32gui
import win32con
import msvcrt
import os
from texttable import Texttable   # type: ignore


# Key bindings that map a single character to an application window title
# (or a part of it)
KEY_BINDINGS: dict[str, str] = {
    'A': '',
    'B': '',
    'C': '',
    'D': '',
    'E': 'Excel',
    'F': '',
    'G': '',
    'H': '',
    'I': 'Notepad++',
    'J': '',
    'K': '',
    'L': '',
    'M': 'Outlook',
    'N': 'Microsoft Teams',
    'O': '',
    'P': 'PDF-XChange Editor',
    'Q': '',
    'R': '',
    'S': 'Edge',
    'T': 'Total Commander',
    'U': '',
    'V': 'Visual Studio Code',
    'W': 'Word',
    'X': '',
    'Y': 'OneNote',
    'Z': ''
}


class AppSwitcher:
    """
    This class allows users to activate application windows via keyboard
    shortcuts.

    It displays a table of assigned shortcut keys and their corresponding
    applications, listens for key input and activates the matching window if
    found.

    Attributes:
        table: The rendered table showing key-to-app mappings.
        open_windows: Reserved for potential future window state caching.
    """
    table_head: list[str]
    table_data: list[list[str]]



    def __init__(self):
        """
        Initializes the AppSwitcher, creates the shortcut table and starts the
        key input loop.
        """
        os.system("title Python App Switcher")  # Set title of console window
        self.create_table()
        self.start_loop()

    def clean_key_bindings(self, bindings: dict[str, str]) -> dict[str, str]:
        """Removes all key bindings where the key or value is an empty string
        (or whitespace).

        Args:
            bindings: The original key bindings dictionary.

        Returns:
            dict: A new dictionary with only valid key-value pairs.
        """
        return {
            key: value
            for key, value in bindings.items()
            if key.strip() != '' and value.strip() != ''
        }

    def create_table(self) -> None:
        """
        Creates a table of key bindings using `Texttable`.

        The table is structured in two columns per row (each with key + app
        name), forming a 4-column layout total.
        """
        # Setup table and columns
        table = Texttable()
        table.set_cols_align(['l', 'l', 'l', 'l'])
        table.set_cols_valign(['m', 'm', 'm', 'm'])
        head = ['Key', 'Application', 'Key', 'Application']

        # Get data and add header
        rows: list[list[str]] = []  # type: ignore
        data_list = list(self.clean_key_bindings(KEY_BINDINGS).items())
        row_count = len(data_list)
        row_count_new = int(row_count / 2) + (1 if row_count % 2 != 0 else 0)

        # Fill cells with data
        for i in range(row_count_new):
            # Column 1
            col_a, col_b = data_list[i]

            # Column 2
            second_column_index = i + row_count_new
            if second_column_index < row_count:
                col_c, col_d = data_list[second_column_index]
            else:
                col_c, col_d = '', ''

            rows.append([col_a, col_b, col_c, col_d])

        rows.insert(0, head)
        table.add_rows(rows)

        self.table = table

    def start_loop(self) -> None:
        """
        Starts the main loop that waits for key input.

        Displays the key-to-app table and listens for user input.
        When a valid key is pressed, attempts to activate the corresponding
        application. Exits the script if any other key that is letter is typed.
        """
        while True:
            # Clear console window and print table with the key-to-app mappings
            clear = lambda: os.system('cls')  # noqa
            clear()
            print(self.table.draw())

            # Get user input; use package msvcrt instead of input() so that
            # the pressing the ENTER key is not necessary
            key_input = msvcrt.getch().decode().capitalize()

            if not key_input.isalpha():
                break

            # Activate window corresponding to the pressed key
            if key_input in KEY_BINDINGS.keys():
                self.activate_app(KEY_BINDINGS[key_input])


    def activate_app(self, target: str) -> None:
        """
        Activates a window matching the given application name.

        Args:
            target: The name (or partial title) of the application window to
            activate.
        """
        if not target:  # Skip if target is empty
            return

        found = False

        def enum_callback(window_handle, _):
            """
            Callback for `EnumWindows` to find and focus the matching window.
            """
            nonlocal found  # Use variable from wrapping function

            # Only process windows that are currently visible to the user.
            # Hidden or background system windows are ignored.
            if win32gui.IsWindowVisible(window_handle):
                title = win32gui.GetWindowText(window_handle).upper()
                if target.upper() in title:
                    print(f'Activating window: {title}')

                    # Check if window is minimized
                    placement = win32gui.GetWindowPlacement(window_handle)
                    is_minimized = placement[1] == win32con.SW_SHOWMINIMIZED

                    try:
                        if is_minimized:
                            # If minimized, use the new method
                            win32gui.ShowWindow(window_handle,
                                                win32con.SW_RESTORE)

                        # Always set as foreground window
                        win32gui.SetForegroundWindow(window_handle)
                        found = True
                    except Exception as e:
                        print(f'Error activating window: {e}')

                    return False  # Stop enumeration

            return True  # Continue

        # List all top level windows and run callback
        try:
            win32gui.EnumWindows(enum_callback, None)
        except Exception as e:
            print(f'An error occurred while checking windows: {e}')

        if not found and target:
            print(f'Window "{target}" not found.')


_ = AppSwitcher()
