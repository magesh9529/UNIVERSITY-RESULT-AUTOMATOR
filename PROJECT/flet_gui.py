import os
import re
import time
import asyncio

import flet as ft

from config import SEMESTER_EXAM_CODES
from result_fetcher import ResultFetcher
from result_parser import ResultParser
from excel_generator import ExcelGenerator


# ============================================================
# CONFIGURATION
# ============================================================

FAIL_GRADES = {
    "F",
    "RA",
    "U",
    "FAIL",
    "AB"
}


# ============================================================
# MAIN APPLICATION
# ============================================================

class UniversityResultApp:

    def __init__(self, page: ft.Page):

        self.page = page

        # ----------------------------
        # Result data
        # ----------------------------
        self.results = []
        self.register_numbers = []

        # ----------------------------
        # Input
        # ----------------------------
        self.semester = None
        self.exam_code = None
        self.report_file_name = ""

        # ----------------------------
        # Counters
        # ----------------------------
        self.total_students = 0
        self.processed_students = 0
        self.success_count = 0
        self.failed_count = 0

        # ----------------------------
        # Timing
        # ----------------------------
        self.start_time = None

        # ----------------------------
        # Excel
        # ----------------------------
        self.output_file = None

        self.setup_page()
        self.show_fetch_page()


    # ========================================================
    # PAGE SETTINGS
    # ========================================================

    def setup_page(self):

        self.page.title = (
            "Pondicherry University Result Automator"
        )

        self.page.window.width = 1200
        self.page.window.height = 760

        self.page.padding = 0
        self.page.bgcolor = "#F5F8FC"

        self.page.theme_mode = ft.ThemeMode.LIGHT


    # ========================================================
    # HEADER
    # ========================================================

    def create_header(self):

        return ft.Container(

            height=100,
            padding=20,
            bgcolor="white",

            content=ft.Row(

                controls=[

                    ft.Row(

                        controls=[

                            ft.Container(

                                width=55,
                                height=55,

                                bgcolor="#E8F1FF",
                                border_radius=30,

                                alignment=ft.Alignment.CENTER,

                                content=ft.Icon(
                                    ft.Icons.SCHOOL,
                                    size=32,
                                    color="#103E91"
                                )

                            ),

                            ft.Column(

                                controls=[

                                    ft.Text(
                                        "PONDICHERRY UNIVERSITY",
                                        size=22,
                                        weight=ft.FontWeight.BOLD,
                                        color="#092B70"
                                    ),

                                    ft.Text(
                                        "Examination Results",
                                        size=17,
                                        weight=ft.FontWeight.BOLD
                                    ),

                                    ft.Text(
                                        "Result Automation to Excel",
                                        size=12,
                                        color="#667085"
                                    )

                                ],

                                spacing=1

                            )

                        ]

                    ),

                    ft.Container(expand=True),

                    ft.Column(

                        controls=[

                            ft.Text(
                                "University Result Automator",
                                size=12,
                                color="#163B80"
                            ),

                            ft.Text(
                                "v1.0.0",
                                size=11,
                                color="#667085"
                            )

                        ],

                        spacing=2,

                        horizontal_alignment=(
                            ft.CrossAxisAlignment.END
                        )

                    )

                ]

            )

        )


        # ========================================================
        # SIDEBAR
        # ========================================================
        # ========================================================
    # CLEAR START REGISTER ERROR
    # ========================================================

    def clear_start_register_error(self, e):

        self.start_register_field.error_text = None

        self.page.update()


    # ========================================================
    # CLEAR END REGISTER ERROR
    # ========================================================

    def clear_end_register_error(self, e):

        self.end_register_field.error_text = None

        self.page.update()
        
    def create_sidebar(self, selected):

        def menu_button(
            title,
            icon,
            page_name
        ):

            active = (
                selected == page_name
            )

            return ft.Container(

                border_radius=8,

                bgcolor=(
                    "#246BFD"
                    if active
                    else None
                ),

                padding=12,

                on_click=lambda e: (
                    self.navigate(page_name)
                ),

                content=ft.Row(

                    controls=[

                        ft.Icon(
                            icon,
                            color="white",
                            size=20
                        ),

                        ft.Text(

                            title,

                            color="white",
                            size=14,

                            weight=(
                                ft.FontWeight.BOLD
                                if active
                                else ft.FontWeight.NORMAL
                            )

                        )

                    ],

                    spacing=12

                )

            )


        return ft.Container(

            width=220,
            bgcolor="#073B5C",
            padding=20,

            content=ft.Column(

                controls=[

                    ft.Container(height=20),

                    ft.Container(

                        alignment=ft.Alignment.CENTER,

                        content=ft.Icon(
                            ft.Icons.SCHOOL,
                            size=65,
                            color="white"
                        )

                    ),

                    ft.Container(height=20),

                    menu_button(
                        "Home",
                        ft.Icons.HOME,
                        "home"
                    ),

                    menu_button(
                        "Fetch Results",
                        ft.Icons.DESCRIPTION,
                        "fetch"
                    ),

                    menu_button(
                        "View Results",
                        ft.Icons.TABLE_CHART,
                        "results"
                    ),

                    menu_button(
                        "Settings",
                        ft.Icons.SETTINGS,
                        "settings"
                    ),

                    menu_button(
                        "About",
                        ft.Icons.INFO,
                        "about"
                    )

                ],

                spacing=8

            )

        )


    # ========================================================
    # NAVIGATION
    # ========================================================

    def navigate(self, page_name):

        if page_name in (
            "home",
            "fetch"
        ):

            self.show_fetch_page()

        elif page_name == "results":

            self.show_results_page()

        elif page_name == "settings":

            self.show_settings_page()

        elif page_name == "about":

            self.show_about_page()


    # ========================================================
    # PAGE LAYOUT
    # ========================================================

    def set_page(
        self,
        selected,
        content
    ):

        self.page.clean()

        self.page.add(

            ft.Row(

                controls=[

                    self.create_sidebar(
                        selected
                    ),

                    ft.Column(

                        controls=[

                            self.create_header(),

                            ft.Container(
                                content=content,
                                padding=25,
                                expand=True
                            )

                        ],

                        expand=True,
                        spacing=0

                    )

                ],

                expand=True,
                spacing=0

            )

        )


    # ========================================================
    # FETCH PAGE
    # ========================================================

    def show_fetch_page(self):

        self.start_register_field = ft.TextField(

            label="Starting Register Number",

            hint_text="Example: 23TN0047",

            prefix_icon=ft.Icons.BADGE,

            width=300,

            border_radius=8,

            max_length=8,

            on_change=self.clear_start_register_error

        )


        self.end_register_field = ft.TextField(

            label="Ending Register Number",

            hint_text="Example: 23TN0050",

            prefix_icon=ft.Icons.BADGE,

            width=300,

            border_radius=8,

            max_length=8,

            on_change=self.clear_end_register_error

        )


        self.semester_dropdown = ft.Dropdown(

            label="Semester",

            hint_text="Select Semester",

            width=300,

            options=[

                ft.dropdown.Option(

                    key=str(i),

                    text=(
                        f"Semester {i} "
                        f"({SEMESTER_EXAM_CODES[i]})"
                    )

                )

                for i in range(1, 9)

            ]

        )


        # ====================================================
        # REPORT FILE NAME
        # ====================================================

        self.file_name_field = ft.TextField(

            label="Report File Name",

            hint_text="Example: CSE_Sem6_June_2026",

            prefix_icon=ft.Icons.DESCRIPTION,

            width=300,

            border_radius=8

        )


        fetch_button = ft.ElevatedButton(

            content=ft.Row(

                controls=[

                    ft.Icon(
                        ft.Icons.SEARCH
                    ),

                    ft.Text(
                        "Start Fetching Results",
                        size=15,
                        weight=ft.FontWeight.BOLD
                    )

                ],

                tight=True

            ),

            height=50,

            on_click=self.start_fetching

        )


        input_card = ft.Container(

            bgcolor="white",

            border_radius=14,

            padding=25,

            content=ft.Column(

                controls=[

                    ft.Row(

                        controls=[

                            ft.Container(

                                width=55,
                                height=55,

                                bgcolor="#E8F1FF",

                                border_radius=10,

                                alignment=ft.Alignment.CENTER,

                                content=ft.Icon(

                                    ft.Icons.DESCRIPTION,

                                    size=30,

                                    color="#103E91"

                                )

                            ),

                            ft.Column(

                                controls=[

                                    ft.Text(

                                        "Fetch Student Results",

                                        size=22,

                                        weight=ft.FontWeight.BOLD

                                    ),

                                    ft.Text(

                                        "Enter the register range, semester and report name.",

                                        color="#667085"

                                    )

                                ],

                                spacing=3

                            )

                        ]

                    ),

                    ft.Divider(),

                    ft.Text(

                        "Register Number Range",

                        size=14,

                        weight=ft.FontWeight.BOLD

                    ),

                    ft.Row(

                        controls=[

                            self.start_register_field,

                            self.end_register_field

                        ],

                        wrap=True,

                        spacing=20

                    ),

                    ft.Text(

                        "Result Information",

                        size=14,

                        weight=ft.FontWeight.BOLD

                    ),

                    ft.Row(

                        controls=[

                            self.semester_dropdown,

                            self.file_name_field

                        ],

                        wrap=True,

                        spacing=20

                    ),

                    ft.Container(

                        bgcolor="#E8F4FF",

                        border_radius=8,

                        padding=15,

                        content=ft.Row(

                            controls=[

                                ft.Icon(

                                    ft.Icons.INFO_OUTLINE,

                                    color="#1677FF"

                                ),

                                ft.Text(

                                    "Results are fetched one by one. "
                                    "Excel is created only when you click "
                                    "'Download Excel Report'.",

                                    size=12

                                )

                            ]

                        )

                    ),

                    ft.Container(height=5),

                    ft.Row(

                        controls=[
                            fetch_button
                        ],

                        alignment=(
                            ft.MainAxisAlignment.CENTER
                        )

                    )

                ],

                spacing=18

            )

        )


        content = ft.Column(

            controls=[

                ft.Text(

                    "Result Automation",

                    size=27,

                    weight=ft.FontWeight.BOLD

                ),

                ft.Text(

                    "Fetch Pondicherry University examination results.",

                    color="#667085"

                ),

                ft.Container(height=10),

                input_card

            ],

            scroll=ft.ScrollMode.AUTO

        )


        self.set_page(
            "fetch",
            content
        )


    # ========================================================
    # START FETCHING
    # ========================================================

# ========================================================
# START FETCHING
# ========================================================

    async def start_fetching(self, e):

        # ----------------------------------------------------
        # CLEAR PREVIOUS FIELD ERRORS
        # ----------------------------------------------------

        self.start_register_field.error_text = None
        self.end_register_field.error_text = None
        self.semester_dropdown.error_text = None
        self.file_name_field.error_text = None

        self.page.update()


        # ----------------------------------------------------
        # GET USER INPUT
        # ----------------------------------------------------

        start = (
            self.start_register_field.value
            or ""
        ).strip().upper()


        end = (
            self.end_register_field.value
            or ""
        ).strip().upper()


        semester_value = (
            self.semester_dropdown.value
        )


        file_name = (
            self.file_name_field.value
            or ""
        ).strip()


        # ----------------------------------------------------
        # EMPTY FIELD VALIDATION
        # ----------------------------------------------------

        has_error = False


        if not start:

            self.start_register_field.error_text = (
                "Starting Register Number is required."
            )

            has_error = True


        if not end:

            self.end_register_field.error_text = (
                "Ending Register Number is required."
            )

            has_error = True


        if not semester_value:

            self.semester_dropdown.error_text = (
                "Please select a semester."
            )

            has_error = True


        if not file_name:

            self.file_name_field.error_text = (
                "Report File Name is required."
            )

            has_error = True


        if has_error:

            self.page.update()

            self.show_snackbar(
                "Please correct the highlighted fields."
            )

            return


        # ----------------------------------------------------
        # START REGISTER NUMBER LENGTH VALIDATION
        # ----------------------------------------------------

        if len(start) != 8:

            self.start_register_field.error_text = (
                "Starting Register Number must be exactly 8 characters."
            )

            self.page.update()

            self.show_snackbar(
                "Starting Register Number must be exactly 8 characters."
            )

            return


        # ----------------------------------------------------
        # END REGISTER NUMBER LENGTH VALIDATION
        # ----------------------------------------------------

        if len(end) != 8:

            self.end_register_field.error_text = (
                "Ending Register Number must be exactly 8 characters."
            )

            self.page.update()

            self.show_snackbar(
                "Ending Register Number must be exactly 8 characters."
            )

            return


        # ----------------------------------------------------
        # REGISTER NUMBER FORMAT VALIDATION
        #
        # Example:
        # 23TN0047
        #
        # Must end with a number because range generation
        # depends on the numeric part.
        # ----------------------------------------------------

        start_match = re.match(
            r"^(.*?)(\d+)$",
            start
        )


        if not start_match:

            self.start_register_field.error_text = (
                "Invalid Starting Register Number format."
            )

            self.page.update()

            self.show_snackbar(
                "Invalid Starting Register Number format."
            )

            return


        end_match = re.match(
            r"^(.*?)(\d+)$",
            end
        )


        if not end_match:

            self.end_register_field.error_text = (
                "Invalid Ending Register Number format."
            )

            self.page.update()

            self.show_snackbar(
                "Invalid Ending Register Number format."
            )

            return


        # ----------------------------------------------------
        # GENERATE REGISTER NUMBERS
        # ----------------------------------------------------

        registers = self.generate_register_numbers(
            start,
            end
        )


        if not registers:

            self.start_register_field.error_text = (
                "Check the register number range."
            )

            self.end_register_field.error_text = (
                "Check the register number range."
            )

            self.page.update()

            self.show_snackbar(
                "Invalid register number range. "
                "Both register numbers must have the same prefix "
                "and the starting number must be smaller than "
                "or equal to the ending number."
            )

            return


        # ----------------------------------------------------
        # CLEAN REPORT FILE NAME
        # ----------------------------------------------------

        if file_name.lower().endswith(
            ".xlsx"
        ):

            file_name = (
                file_name[:-5]
            )


        # Remove invalid Windows filename characters
        file_name = re.sub(

            r'[<>:"/\\|?*]',

            "_",

            file_name

        )


        if not file_name.strip():

            self.file_name_field.error_text = (
                "Please enter a valid Report File Name."
            )

            self.page.update()

            self.show_snackbar(
                "Please enter a valid Report File Name."
            )

            return


        # ----------------------------------------------------
        # STORE REPORT FILE NAME
        # ----------------------------------------------------

        self.report_file_name = (
            file_name
        )


        # ----------------------------------------------------
        # STORE SEMESTER
        # ----------------------------------------------------

        self.semester = int(
            semester_value
        )


        self.exam_code = (

            SEMESTER_EXAM_CODES[
                self.semester
            ]

        )


        # ----------------------------------------------------
        # STORE REGISTER NUMBERS
        # ----------------------------------------------------

        self.register_numbers = (
            registers
        )


        # ----------------------------------------------------
        # RESET CURRENT RUN
        # ----------------------------------------------------

        self.results = []

        self.total_students = len(
            registers
        )

        self.processed_students = 0

        self.success_count = 0

        self.failed_count = 0

        self.output_file = None


        # ----------------------------------------------------
        # SHOW PROCESSING PAGE
        # ----------------------------------------------------

        self.show_processing_page()


        # Give Flet time to display the processing page
        await asyncio.sleep(
            0.1
        )


        # ----------------------------------------------------
        # START SEQUENTIAL PROCESSING
        # ----------------------------------------------------

        await self.process_all_students()
        # ========================================================
        # GENERATE REGISTER NUMBERS
        # ========================================================

    def generate_register_numbers(
        self,
        start,
        end
    ):

        start_match = re.match(

            r"^(.*?)(\d+)$",

            start

        )


        end_match = re.match(

            r"^(.*?)(\d+)$",

            end

        )


        if (
            not start_match
            or not end_match
        ):

            return None


        start_prefix = (
            start_match.group(1)
        )


        end_prefix = (
            end_match.group(1)
        )


        if (
            start_prefix
            != end_prefix
        ):

            return None


        start_number_text = (
            start_match.group(2)
        )


        end_number_text = (
            end_match.group(2)
        )


        start_number = int(
            start_number_text
        )


        end_number = int(
            end_number_text
        )


        if (
            start_number
            > end_number
        ):

            return None


        number_length = len(
            start_number_text
        )


        register_numbers = []


        # NORMAL FOR LOOP
        for number in range(
            start_number,
            end_number + 1
        ):

            register_number = (

                start_prefix

                + str(number).zfill(
                    number_length
                )

            ).upper()


            register_numbers.append(
                register_number
            )


        return register_numbers


    # ========================================================
    # PROCESSING PAGE
    # ========================================================

    def show_processing_page(self):

        self.progress_bar = ft.ProgressBar(

            value=0,

            height=12,

            border_radius=8

        )


        self.progress_percentage = ft.Text(

            "0%",

            size=17,

            weight=ft.FontWeight.BOLD

        )


        self.processing_text = ft.Text(

            "Preparing...",

            color="#475467"

        )


        self.completed_text = ft.Text(

            (
                f"0 / "
                f"{self.total_students} "
                f"students completed"
            ),

            color="#475467"

        )


        self.eta_text = ft.Text(

            "Elapsed: 0 sec   |   ETA: Calculating...",

            color="#667085",

            size=13

        )


        self.total_card_value = ft.Text(

            str(self.total_students),

            size=25,

            weight=ft.FontWeight.BOLD

        )


        self.processed_card_value = ft.Text(

            "0",

            size=25,

            weight=ft.FontWeight.BOLD

        )


        self.success_card_value = ft.Text(

            "0",

            size=25,

            weight=ft.FontWeight.BOLD,

            color="#14804A"

        )


        self.failed_card_value = ft.Text(

            "0",

            size=25,

            weight=ft.FontWeight.BOLD,

            color="#D92D20"

        )


        self.status_column = ft.Column(

            controls=[],

            spacing=6

        )


        progress_card = ft.Container(

            bgcolor="white",

            padding=22,

            border_radius=12,

            content=ft.Column(

                controls=[

                    ft.Row(

                        controls=[

                            ft.Container(

                                width=50,
                                height=50,

                                bgcolor="#E8F1FF",

                                border_radius=10,

                                alignment=(
                                    ft.Alignment.CENTER
                                ),

                                content=ft.Icon(

                                    ft.Icons.AUTORENEW,

                                    size=30,

                                    color="#1769E8"

                                )

                            ),

                            ft.Column(

                                controls=[

                                    ft.Text(

                                        "Fetching Results...",

                                        size=22,

                                        weight=(
                                            ft.FontWeight.BOLD
                                        )

                                    ),

                                    ft.Text(

                                        "Students are being processed one by one.",

                                        color="#667085"

                                    )

                                ],

                                spacing=2

                            )

                        ]

                    ),

                    ft.Container(height=5),

                    ft.Row(

                        controls=[

                            ft.Container(

                                content=(
                                    self.progress_bar
                                ),

                                expand=True

                            ),

                            self.progress_percentage

                        ],

                        spacing=15

                    ),

                    ft.Row(

                        controls=[

                            self.processing_text,

                            ft.Container(
                                expand=True
                            ),

                            self.completed_text

                        ]

                    ),

                    self.eta_text

                ],

                spacing=12

            )

        )


        cards = ft.Row(

            controls=[

                self.summary_card(

                    "Total Students",

                    self.total_card_value,

                    ft.Icons.GROUPS,

                    "#1769E8"

                ),

                self.summary_card(

                    "Processed",

                    self.processed_card_value,

                    ft.Icons.AUTORENEW,

                    "#7A5AF8"

                ),

                self.summary_card(

                    "Successful",

                    self.success_card_value,

                    ft.Icons.CHECK_CIRCLE,

                    "#12B76A"

                ),

                self.summary_card(

                    "Unavailable",

                    self.failed_card_value,

                    ft.Icons.CANCEL,

                    "#F04438"

                )

            ],

            wrap=True

        )


        status_card = ft.Container(

            bgcolor="white",

            padding=20,

            border_radius=12,

            content=ft.Column(

                controls=[

                    ft.Text(

                        "Current Status",

                        size=18,

                        weight=ft.FontWeight.BOLD

                    ),

                    ft.Divider(),

                    self.status_column

                ]

            )

        )


        content = ft.Column(

            controls=[

                progress_card,

                cards,

                status_card

            ],

            spacing=15,

            scroll=ft.ScrollMode.AUTO

        )


        self.set_page(
            "fetch",
            content
        )


    # ========================================================
    # SUMMARY CARD
    # ========================================================

    def summary_card(
        self,
        title,
        value_control,
        icon,
        icon_color
    ):

        return ft.Container(

            width=205,

            bgcolor="white",

            padding=18,

            border_radius=10,

            content=ft.Row(

                controls=[

                    ft.Container(

                        width=45,
                        height=45,

                        border_radius=10,

                        bgcolor="#F2F4F7",

                        alignment=(
                            ft.Alignment.CENTER
                        ),

                        content=ft.Icon(

                            icon,

                            size=26,

                            color=icon_color

                        )

                    ),

                    ft.Column(

                        controls=[

                            ft.Text(

                                title,

                                size=11,

                                color="#667085"

                            ),

                            value_control

                        ],

                        spacing=2

                    )

                ],

                spacing=12

            )

        )


    # ========================================================
    # PROCESS ONE STUDENT
    # ========================================================

    def process_one_student(
        self,
        register_number
    ):

        try:

            fetcher = ResultFetcher(

                register_number,

                self.exam_code

            )


            response = (
                fetcher.fetch_result()
            )


            if not response:

                return None


            if (
                response.get(
                    "error_code"
                )
                == -1
            ):

                return None


            data = response.get(
                "data",
                {}
            )


            html = data.get(
                "html"
            )


            if not html:

                return None


            parser = ResultParser(
                html
            )


            student = (
                parser.parse_student_info()
            )


            subjects = (
                parser.parse_subjects()
            )


            if student is None:

                return None


            if not subjects:

                return None


            return {

                "student": student,

                "subjects": subjects

            }


        except Exception as error:

            print(

                f"Error processing "
                f"{register_number}: "
                f"{error}"

            )

            return None


    # ========================================================
    # PROCESS ALL STUDENTS
    #
    # NO MULTITHREADING
    # NORMAL FOR LOOP
    # ========================================================

    async def process_all_students(self):

        self.start_time = (
            time.time()
        )


        for register_number in (
            self.register_numbers
        ):

            # -----------------------------------------------
            # CURRENT STUDENT
            # -----------------------------------------------

            self.processing_text.value = (

                f"Processing: "
                f"{register_number}"

            )


            self.page.update()


            # Let GUI draw current student
            await asyncio.sleep(0.05)


            # -----------------------------------------------
            # FETCH RESULT
            # -----------------------------------------------

            result = (
                self.process_one_student(
                    register_number
                )
            )


            # -----------------------------------------------
            # COMPLETED
            # -----------------------------------------------

            self.processed_students += 1


            if result:

                self.results.append(
                    result
                )

                self.success_count += 1

                status = "Success"


            else:

                self.failed_count += 1

                status = "Failed"


            # -----------------------------------------------
            # UPDATE GUI
            # -----------------------------------------------

            self.update_processing_ui(

                register_number,

                status

            )


            # Allow progress repaint
            await asyncio.sleep(0.05)


        # -----------------------------------------------
        # SORT RESULTS
        # -----------------------------------------------

        self.results.sort(

            key=lambda result:
                result[
                    "student"
                ][
                    "register_number"
                ]

        )


        # IMPORTANT:
        # DO NOT CREATE EXCEL HERE

        await asyncio.sleep(0.4)


        self.show_success_page()


    # ========================================================
    # UPDATE PROGRESS + ETA
    # ========================================================

    def update_processing_ui(
        self,
        register_number,
        status
    ):

        if self.total_students == 0:
            return


        # ----------------------------------------------------
        # PROGRESS
        # ----------------------------------------------------

        progress = (

            self.processed_students

            / self.total_students

        )


        percentage = (
            progress * 100
        )


        self.progress_bar.value = (
            progress
        )


        self.progress_percentage.value = (

            f"{percentage:.0f}%"

        )


        self.completed_text.value = (

            f"{self.processed_students} / "
            f"{self.total_students} "
            f"students completed"

        )


        # ----------------------------------------------------
        # ELAPSED TIME
        # ----------------------------------------------------

        elapsed_time = (

            time.time()

            - self.start_time

        )


        # ----------------------------------------------------
        # ETA
        # ----------------------------------------------------

        if self.processed_students > 0:

            average_time = (

                elapsed_time

                / self.processed_students

            )


            remaining_students = (

                self.total_students

                - self.processed_students

            )


            eta_seconds = (

                average_time

                * remaining_students

            )


            if remaining_students == 0:

                eta_display = (
                    "Completed"
                )


            elif eta_seconds >= 60:

                minutes = int(
                    eta_seconds // 60
                )

                seconds = int(
                    eta_seconds % 60
                )

                eta_display = (

                    f"{minutes} min "
                    f"{seconds} sec"

                )


            else:

                eta_display = (

                    f"{eta_seconds:.0f} sec"

                )


            self.eta_text.value = (

                f"Elapsed: "
                f"{elapsed_time:.1f} sec"
                f"   |   "
                f"ETA: "
                f"{eta_display}"

            )


        # ----------------------------------------------------
        # COUNTERS
        # ----------------------------------------------------

        self.processed_card_value.value = (

            str(
                self.processed_students
            )

        )


        self.success_card_value.value = (

            str(
                self.success_count
            )

        )


        self.failed_card_value.value = (

            str(
                self.failed_count
            )

        )


        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        if status == "Success":

            status_icon = (
                ft.Icons.CHECK_CIRCLE
            )

            status_color = (
                "#12B76A"
            )

            status_message = (
                "Result fetched successfully"
            )


        else:

            status_icon = (
                ft.Icons.CANCEL
            )

            status_color = (
                "#F04438"
            )

            status_message = (
                "Result unavailable"
            )


        status_row = ft.Container(

            bgcolor="#F9FAFB",

            border_radius=6,

            padding=8,

            content=ft.Row(

                controls=[

                    ft.Text(

                        (
                            f"["
                            f"{self.processed_students}"
                            f"/"
                            f"{self.total_students}"
                            f"]"
                        ),

                        size=12,

                        color="#667085"

                    ),

                    ft.Text(

                        register_number,

                        weight=(
                            ft.FontWeight.BOLD
                        ),

                        size=12

                    ),

                    ft.Container(
                        expand=True
                    ),

                    ft.Icon(

                        status_icon,

                        color=status_color,

                        size=18

                    ),

                    ft.Text(

                        status_message,

                        size=12,

                        color=status_color

                    )

                ]

            )

        )


        self.status_column.controls.insert(

            0,

            status_row

        )


        # Show latest 7
        self.status_column.controls = (

            self.status_column.controls[
                :7
            ]

        )


        self.page.update()


    # ========================================================
    # DOWNLOAD / SAVE EXCEL
    # ========================================================

    def save_excel(self):

        # ----------------------------------------------------
        # CHECK RESULTS
        # ----------------------------------------------------

        if not self.results:

            self.show_snackbar(
                "No results available to download."
            )

            return


        # ----------------------------------------------------
        # OUTPUT FOLDER
        # ----------------------------------------------------

        os.makedirs(

            "output",

            exist_ok=True

        )


        # ----------------------------------------------------
        # USER FILE NAME
        # ----------------------------------------------------

        base_name = (
            self.report_file_name
        )


        file_name = (
            base_name
            + ".xlsx"
        )


        file_path = os.path.join(

            "output",

            file_name

        )


        # ----------------------------------------------------
        # DON'T REPLACE OLD FILE
        # ----------------------------------------------------

        counter = 1


        while os.path.exists(
            file_path
        ):

            file_name = (

                f"{base_name}_"
                f"{counter}.xlsx"

            )


            file_path = os.path.join(

                "output",

                file_name

            )


            counter += 1


        # ----------------------------------------------------
        # GENERATE EXCEL
        # ----------------------------------------------------

        try:

            generator = ExcelGenerator(
                file_path
            )


            generator.create_excel(

                self.results,

                self.semester

            )


            self.output_file = (
                file_path
            )


            self.show_snackbar(

                f"Report created successfully: "
                f"{file_name}"

            )


        except Exception as error:

            self.show_snackbar(

                f"Unable to create report: "
                f"{error}"

            )


    # ========================================================
    # SUCCESS PAGE
    # ========================================================

    def show_success_page(self):

        elapsed = 0


        if self.start_time:

            elapsed = (

                time.time()

                - self.start_time

            )


        content = ft.Column(

            horizontal_alignment=(
                ft.CrossAxisAlignment.CENTER
            ),

            controls=[

                ft.Container(height=10),

                # ------------------------------------------------
                # SUCCESS ICON
                # ------------------------------------------------

                ft.Container(

                    width=110,
                    height=110,

                    border_radius=55,

                    bgcolor="#DDF7E8",

                    alignment=(
                        ft.Alignment.CENTER
                    ),

                    content=ft.Icon(

                        ft.Icons.CHECK_CIRCLE,

                        size=75,

                        color="#20B15A"

                    )

                ),

                ft.Text(

                    "Process Completed!",

                    size=27,

                    weight=ft.FontWeight.BOLD

                ),

                ft.Text(

                    "All register numbers have been processed.",

                    color="#667085"

                ),

                # ------------------------------------------------
                # COUNTS
                # ------------------------------------------------

                ft.Row(

                    controls=[

                        self.small_stat(

                            "Total Students",

                            self.total_students

                        ),

                        self.small_stat(

                            "Successful",

                            self.success_count

                        ),

                        self.small_stat(

                            "Unavailable",

                            self.failed_count

                        )

                    ],

                    alignment=(
                        ft.MainAxisAlignment.CENTER
                    ),

                    wrap=True

                ),

                # ------------------------------------------------
                # TIME
                # ------------------------------------------------

                ft.Text(

                    (
                        f"Total processing time: "
                        f"{elapsed:.1f} seconds"
                    ),

                    color="#667085"

                ),

                # ------------------------------------------------
                # REPORT INFORMATION
                # ------------------------------------------------

                ft.Container(

                    width=700,

                    bgcolor="#E8F4FF",

                    padding=18,

                    border_radius=10,

                    content=ft.Row(

                        controls=[

                            ft.Icon(

                                ft.Icons.DESCRIPTION,

                                color="#1769E8",

                                size=30

                            ),

                            ft.Column(

                                controls=[

                                    ft.Text(

                                        "Excel Report Ready",

                                        weight=(
                                            ft.FontWeight.BOLD
                                        ),

                                        size=15

                                    ),

                                    ft.Text(

                                        (
                                            "Report name: "
                                            f"{self.report_file_name}.xlsx"
                                        ),

                                        color="#475467"

                                    ),

                                    ft.Text(

                                        (
                                            "Click Download Excel Report "
                                            "to generate the file."
                                        ),

                                        color="#667085",

                                        size=11

                                    )

                                ],

                                spacing=2

                            )

                        ]

                    )

                ),

                # ------------------------------------------------
                # BUTTONS
                # ------------------------------------------------

                ft.Row(

                    controls=[

                        ft.ElevatedButton(

                            "Download Excel Report",

                            icon=ft.Icons.DOWNLOAD,

                            on_click=lambda e:
                                self.save_excel()

                        ),

                        ft.ElevatedButton(

                            "View Results",

                            icon=(
                                ft.Icons.TABLE_CHART
                            ),

                            on_click=lambda e:
                                self.show_results_page()

                        ),

                        ft.ElevatedButton(

                            "Fetch New Results",

                            icon=ft.Icons.REFRESH,

                            on_click=lambda e:
                                self.show_fetch_page()

                        )

                    ],

                    alignment=(
                        ft.MainAxisAlignment.CENTER
                    ),

                    wrap=True

                )

            ],

            spacing=15,

            scroll=ft.ScrollMode.AUTO

        )


        self.set_page(
            "results",
            content
        )


    # ========================================================
    # SMALL STAT CARD
    # ========================================================

    def small_stat(
        self,
        title,
        value
    ):

        return ft.Container(

            width=190,

            bgcolor="white",

            padding=18,

            border_radius=10,

            content=ft.Column(

                controls=[

                    ft.Text(

                        title,

                        color="#667085",

                        size=12

                    ),

                    ft.Text(

                        str(value),

                        size=25,

                        weight=(
                            ft.FontWeight.BOLD
                        )

                    )

                ],

                spacing=2

            )

        )


    # ========================================================
    # CALCULATE SGPA
    # ========================================================

    def calculate_sgpa(
        self,
        subjects
    ):

        total_credit = 0
        total_credit_point = 0


        for subject in subjects:

            try:

                credit = float(
                    subject["credit"]
                )


                credit_point = float(
                    subject["credit_point"]
                )


                total_credit += credit

                total_credit_point += (
                    credit_point
                )


            except (
                ValueError,
                TypeError,
                KeyError
            ):

                continue


        if total_credit == 0:

            return 0


        sgpa = (

            total_credit_point

            / total_credit

        )


        return round(
            sgpa,
            2
        )


    # ========================================================
    # STUDENT RESULT STATUS
    # ========================================================

    def get_student_result_status(
        self,
        subjects
    ):

        for subject in subjects:

            grade = (

                str(
                    subject.get(
                        "grade",
                        ""
                    )
                )
                .strip()
                .upper()

            )


            if grade in FAIL_GRADES:

                return "Fail"


        return "Pass"


    # ========================================================
    # RESULTS PAGE
    # ========================================================

    def show_results_page(self):

        if not self.results:

            content = ft.Column(

                controls=[

                    ft.Text(

                        "Fetched Results Preview",

                        size=25,

                        weight=(
                            ft.FontWeight.BOLD
                        )

                    ),

                    ft.Container(

                        bgcolor="white",

                        padding=30,

                        border_radius=12,

                        content=ft.Column(

                            controls=[

                                ft.Icon(

                                    ft.Icons.INFO_OUTLINE,

                                    size=45,

                                    color="#1769E8"

                                ),

                                ft.Text(

                                    "No results have been fetched yet.",

                                    size=16

                                ),

                                ft.ElevatedButton(

                                    "Fetch Results",

                                    on_click=lambda e:
                                        self.show_fetch_page()

                                )

                            ]

                        )

                    )

                ]

            )


            self.set_page(
                "results",
                content
            )

            return


        # ----------------------------------------------------
        # CREATE STUDENT TABLE
        # ----------------------------------------------------

        rows = []


        for index, result in enumerate(

            self.results,

            start=1

        ):

            student = (
                result["student"]
            )


            subjects = (
                result["subjects"]
            )


            sgpa = (
                self.calculate_sgpa(
                    subjects
                )
            )


            result_status = (
                self.get_student_result_status(
                    subjects
                )
            )


            if result_status == "Pass":

                status_bg = "#DDF7E8"

                status_color = "#16803D"


            else:

                status_bg = "#FEE4E2"

                status_color = "#D92D20"


            rows.append(

                ft.DataRow(

                    cells=[

                        ft.DataCell(

                            ft.Text(
                                str(index)
                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                student[
                                    "register_number"
                                ]

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                student[
                                    "student_name"
                                ]

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                student.get(
                                    "exam_name",
                                    ""
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                str(
                                    len(subjects)
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(
                                str(sgpa)
                            )

                        ),

                        ft.DataCell(

                            ft.Container(

                                bgcolor=status_bg,

                                padding=6,

                                border_radius=6,

                                content=ft.Text(

                                    result_status,

                                    color=status_color

                                )

                            )

                        ),

                        ft.DataCell(

                            ft.IconButton(

                                icon=(
                                    ft.Icons.VISIBILITY
                                ),

                                tooltip=(
                                    "View Details"
                                ),

                                on_click=lambda e,
                                current_result=result:
                                    self.show_student_dialog(
                                        current_result
                                    )

                            )

                        )

                    ]

                )

            )


        student_table = ft.DataTable(

            columns=[

                ft.DataColumn(
                    ft.Text("S.No")
                ),

                ft.DataColumn(
                    ft.Text(
                        "Register Number"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Student Name"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Exam Name"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Subjects"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "SGPA"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Result"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Action"
                    )
                )

            ],

            rows=rows

        )


        # ----------------------------------------------------
        # DOWNLOAD BUTTON
        # ----------------------------------------------------

        download_button = ft.ElevatedButton(

            "Download Excel Report",

            icon=ft.Icons.DOWNLOAD,

            on_click=lambda e:
                self.save_excel()

        )


        content = ft.Column(

            controls=[

                ft.Row(

                    controls=[

                        ft.Column(

                            controls=[

                                ft.Text(

                                    "Fetched Results Preview",

                                    size=25,

                                    weight=(
                                        ft.FontWeight.BOLD
                                    )

                                ),

                                ft.Text(

                                    (
                                        "Student-wise and "
                                        "subject-wise result analysis."
                                    ),

                                    color="#667085"

                                )

                            ],

                            spacing=2

                        ),

                        ft.Container(
                            expand=True
                        ),

                        download_button

                    ]

                ),

                ft.Row(

                    controls=[

                        self.small_stat(

                            "Total Students",

                            self.total_students

                        ),

                        self.small_stat(

                            "Successfully Fetched",

                            self.success_count

                        ),

                        self.small_stat(

                            "Unavailable",

                            self.failed_count

                        )

                    ],

                    wrap=True

                ),

                ft.Container(

                    bgcolor="white",

                    border_radius=12,

                    padding=15,

                    content=ft.Column(

                        controls=[

                            ft.Text(

                                "Student Wise Results",

                                size=19,

                                weight=(
                                    ft.FontWeight.BOLD
                                )

                            ),

                            ft.Row(

                                controls=[
                                    student_table
                                ],

                                scroll=(
                                    ft.ScrollMode.AUTO
                                )

                            )

                        ]

                    )

                ),

                self.create_subject_analysis()

            ],

            spacing=15,

            scroll=ft.ScrollMode.AUTO

        )


        self.set_page(
            "results",
            content
        )


    # ========================================================
    # SUBJECT ANALYSIS DATA
    # ========================================================

    def get_subject_analysis(self):

        analysis = {}


        for result in self.results:

            subjects = (
                result["subjects"]
            )


            for subject in subjects:

                subject_name = (

                    str(
                        subject[
                            "subject_name"
                        ]
                    )
                    .strip()

                )


                grade = (

                    str(
                        subject.get(
                            "grade",
                            ""
                        )
                    )
                    .strip()
                    .upper()

                )


                if (
                    subject_name
                    not in analysis
                ):

                    analysis[
                        subject_name
                    ] = {

                        "appeared": 0,

                        "passed": 0,

                        "failed": 0

                    }


                analysis[
                    subject_name
                ][
                    "appeared"
                ] += 1


                if grade in FAIL_GRADES:

                    analysis[
                        subject_name
                    ][
                        "failed"
                    ] += 1


                else:

                    analysis[
                        subject_name
                    ][
                        "passed"
                    ] += 1


        return analysis


    # ========================================================
    # SUBJECT ANALYSIS UI
    # ========================================================

    def create_subject_analysis(self):

        analysis = (
            self.get_subject_analysis()
        )


        table_rows = []

        visual_rows = []


        for (
            subject_name,
            data
        ) in analysis.items():

            appeared = (
                data["appeared"]
            )

            passed = (
                data["passed"]
            )

            failed = (
                data["failed"]
            )


            if appeared > 0:

                percentage = (

                    passed

                    / appeared

                    * 100

                )

            else:

                percentage = 0


            # ------------------------------------------------
            # TABLE
            # ------------------------------------------------

            table_rows.append(

                ft.DataRow(

                    cells=[

                        ft.DataCell(

                            ft.Text(
                                subject_name
                            )

                        ),

                        ft.DataCell(

                            ft.Text(
                                str(appeared)
                            )

                        ),

                        ft.DataCell(

                            ft.Text(
                                str(passed)
                            )

                        ),

                        ft.DataCell(

                            ft.Text(
                                str(failed)
                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                f"{percentage:.1f}%",

                                weight=(
                                    ft.FontWeight.BOLD
                                ),

                                color=(
                                    "#16803D"
                                    if failed == 0
                                    else "#1769E8"
                                )

                            )

                        )

                    ]

                )

            )


            # ------------------------------------------------
            # VISUAL BAR
            # ------------------------------------------------

            progress_value = (

                percentage
                / 100

            )


            visual_rows.append(

                ft.Container(

                    bgcolor="#F9FAFB",

                    border_radius=8,

                    padding=12,

                    content=ft.Column(

                        controls=[

                            ft.Row(

                                controls=[

                                    ft.Text(

                                        subject_name,

                                        weight=(
                                            ft.FontWeight.BOLD
                                        ),

                                        size=12

                                    ),

                                    ft.Container(
                                        expand=True
                                    ),

                                    ft.Text(

                                        f"Passed: {passed}",

                                        color="#16803D",

                                        size=11

                                    ),

                                    ft.Text(

                                        f"Failed: {failed}",

                                        color="#D92D20",

                                        size=11

                                    )

                                ]

                            ),

                            ft.ProgressBar(

                                value=(
                                    progress_value
                                ),

                                height=10,

                                border_radius=6

                            ),

                            ft.Text(

                                (
                                    "100% Result"
                                    if failed == 0
                                    else
                                    f"{percentage:.1f}% Pass"
                                ),

                                size=11,

                                color=(
                                    "#16803D"
                                    if failed == 0
                                    else "#475467"
                                )

                            )

                        ],

                        spacing=5

                    )

                )

            )


        subject_table = ft.DataTable(

            columns=[

                ft.DataColumn(
                    ft.Text(
                        "Subject"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Appeared"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Passed"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Failed"
                    )
                ),

                ft.DataColumn(
                    ft.Text(
                        "Pass %"
                    )
                )

            ],

            rows=table_rows

        )


        return ft.Container(

            bgcolor="white",

            padding=20,

            border_radius=12,

            content=ft.Column(

                controls=[

                    ft.Row(

                        controls=[

                            ft.Icon(

                                ft.Icons.BAR_CHART,

                                color="#1769E8"

                            ),

                            ft.Text(

                                "Subject Wise Result Analysis",

                                size=20,

                                weight=(
                                    ft.FontWeight.BOLD
                                )

                            )

                        ]

                    ),

                    ft.Text(

                        (
                            "Shows how many students passed "
                            "or failed each subject."
                        ),

                        color="#667085"

                    ),

                    ft.Divider(),

                    ft.Text(

                        "Pass Percentage Overview",

                        size=16,

                        weight=(
                            ft.FontWeight.BOLD
                        )

                    ),

                    ft.Column(

                        controls=visual_rows,

                        spacing=8

                    ),

                    ft.Divider(),

                    ft.Text(

                        "Detailed Subject Result",

                        size=16,

                        weight=(
                            ft.FontWeight.BOLD
                        )

                    ),

                    ft.Row(

                        controls=[
                            subject_table
                        ],

                        scroll=(
                            ft.ScrollMode.AUTO
                        )

                    )

                ],

                spacing=12

            )

        )


    # ========================================================
    # STUDENT DETAILS DIALOG
    # ========================================================

    def show_student_dialog(
        self,
        result
    ):

        student = (
            result["student"]
        )


        subjects = (
            result["subjects"]
        )


        subject_rows = []


        for index, subject in enumerate(

            subjects,

            start=1

        ):

            subject_rows.append(

                ft.DataRow(

                    cells=[

                        ft.DataCell(

                            ft.Text(
                                str(index)
                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                subject.get(
                                    "subject_name",
                                    ""
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                subject.get(
                                    "type",
                                    ""
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                subject.get(
                                    "credit",
                                    ""
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                subject.get(
                                    "grade_point",
                                    ""
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                subject.get(
                                    "credit_point",
                                    ""
                                )

                            )

                        ),

                        ft.DataCell(

                            ft.Text(

                                subject.get(
                                    "grade",
                                    ""
                                )

                            )

                        )

                    ]

                )

            )


        sgpa = (
            self.calculate_sgpa(
                subjects
            )
        )


        result_status = (
            self.get_student_result_status(
                subjects
            )
        )


        dialog = ft.AlertDialog(

            modal=True,

            title=ft.Text(

                (
                    "Student Result - "
                    + student[
                        "register_number"
                    ]
                )

            ),

            content=ft.Container(

                width=850,

                content=ft.Column(

                    controls=[

                        ft.Container(

                            bgcolor="#E8F4FF",

                            padding=15,

                            border_radius=8,

                            content=ft.Column(

                                controls=[

                                    ft.Text(

                                        (
                                            "Register Number : "
                                            + student[
                                                "register_number"
                                            ]
                                        ),

                                        weight=(
                                            ft.FontWeight.BOLD
                                        )

                                    ),

                                    ft.Text(

                                        (
                                            "Student Name : "
                                            + student[
                                                "student_name"
                                            ]
                                        )

                                    ),

                                    ft.Text(

                                        (
                                            "Semester : "
                                            + str(
                                                self.semester
                                            )
                                        )

                                    ),

                                    ft.Text(

                                        (
                                            "SGPA : "
                                            + str(sgpa)
                                        )

                                    ),

                                    ft.Text(

                                        (
                                            "Result : "
                                            + result_status
                                        )

                                    ),

                                    ft.Text(

                                        (
                                            "Exam : "
                                            + student.get(
                                                "exam_name",
                                                ""
                                            )
                                        )

                                    )

                                ],

                                spacing=4

                            )

                        ),

                        ft.Row(

                            controls=[

                                ft.DataTable(

                                    columns=[

                                        ft.DataColumn(
                                            ft.Text("S.No")
                                        ),

                                        ft.DataColumn(
                                            ft.Text("Subject")
                                        ),

                                        ft.DataColumn(
                                            ft.Text("Type")
                                        ),

                                        ft.DataColumn(
                                            ft.Text("Credit")
                                        ),

                                        ft.DataColumn(
                                            ft.Text(
                                                "Grade Point"
                                            )
                                        ),

                                        ft.DataColumn(
                                            ft.Text(
                                                "Credit Point"
                                            )
                                        ),

                                        ft.DataColumn(
                                            ft.Text("Grade")
                                        )

                                    ],

                                    rows=(
                                        subject_rows
                                    )

                                )

                            ],

                            scroll=(
                                ft.ScrollMode.AUTO
                            )

                        )

                    ],

                    tight=True,

                    scroll=(
                        ft.ScrollMode.AUTO
                    )

                )

            ),

            actions=[

                ft.TextButton(

                    "Close",

                    on_click=lambda e:
                        self.close_dialog(
                            dialog
                        )

                )

            ]

        )


        self.page.overlay.append(
            dialog
        )


        dialog.open = True


        self.page.update()


    # ========================================================
    # CLOSE DIALOG
    # ========================================================

    def close_dialog(
        self,
        dialog
    ):

        dialog.open = False

        self.page.update()


    # ========================================================
    # SETTINGS PAGE
    # ========================================================

    def show_settings_page(self):

        content = ft.Column(

            controls=[

                ft.Text(

                    "Settings",

                    size=26,

                    weight=(
                        ft.FontWeight.BOLD
                    )

                ),

                ft.Container(

                    bgcolor="white",

                    padding=25,

                    border_radius=12,

                    content=ft.Column(

                        controls=[

                            ft.Text(

                                "Processing Settings",

                                size=18,

                                weight=(
                                    ft.FontWeight.BOLD
                                )

                            ),

                            ft.Divider(),

                            ft.Text(
                                "Processing Mode: Sequential"
                            ),

                            ft.Text(
                                "Students are processed one by one using a normal for loop."
                            ),

                            ft.Text(
                                "Multithreading: Disabled"
                            ),

                            ft.Text(
                                "Excel output folder: output/"
                            ),

                            ft.Text(
                                "Excel is created only when Download Excel Report is clicked."
                            ),

                            ft.Text(
                                "Existing Excel files are not overwritten."
                            ),

                            ft.Text(
                                "Progress and ETA are updated after each student."
                            )

                        ],

                        spacing=10

                    )

                )

            ]

        )


        self.set_page(
            "settings",
            content
        )


    # ========================================================
    # ABOUT PAGE
    # ========================================================

    def show_about_page(self):

        content = ft.Column(

            controls=[

                ft.Text(

                    "About",

                    size=26,

                    weight=(
                        ft.FontWeight.BOLD
                    )

                ),

                ft.Container(

                    bgcolor="white",

                    padding=25,

                    border_radius=12,

                    content=ft.Column(

                        controls=[

                            ft.Icon(

                                ft.Icons.SCHOOL,

                                size=55,

                                color="#1769E8"

                            ),

                            ft.Text(

                                "University Result Automator",

                                size=21,

                                weight=(
                                    ft.FontWeight.BOLD
                                )

                            ),

                            ft.Text(
                                "Version 1.0.0"
                            ),

                            ft.Divider(),

                            ft.Text(

                                (
                                    "Python + Flet application "
                                    "for fetching Pondicherry "
                                    "University examination results."
                                )

                            ),

                            ft.Text(
                                "Features:"
                            ),

                            ft.Text(
                                "• Register number range processing"
                            ),

                            ft.Text(
                                "• Sequential for-loop processing"
                            ),

                            ft.Text(
                                "• Live progress percentage"
                            ),

                            ft.Text(
                                "• Elapsed time and ETA"
                            ),

                            ft.Text(
                                "• User-defined report name"
                            ),

                            ft.Text(
                                "• Separate Excel download"
                            ),

                            ft.Text(
                                "• Student-wise results"
                            ),

                            ft.Text(
                                "• Subject-wise result analysis"
                            ),

                            ft.Text(
                                "• SGPA calculation"
                            ),

                            ft.Text(
                                "• Excel report generation"
                            )

                        ],

                        spacing=8

                    )

                )

            ],

            scroll=(
                ft.ScrollMode.AUTO
            )

        )


        self.set_page(
            "about",
            content
        )


    # ========================================================
    # SNACKBAR
    # ========================================================

    def show_snackbar(
        self,
        message
    ):

        snackbar = ft.SnackBar(

            content=ft.Text(
                message
            )

        )


        self.page.overlay.append(
            snackbar
        )


        snackbar.open = True


        self.page.update()


# ============================================================
# FLET MAIN
# ============================================================

def main(page: ft.Page):

    UniversityResultApp(
        page
    )


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":

    ft.run(main)