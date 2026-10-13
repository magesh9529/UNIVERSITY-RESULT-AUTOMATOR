
import os
import re
import time 

from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

from config import SEMESTER_EXAM_CODES
from result_fetcher import ResultFetcher
from result_parser import ResultParser
from excel_generator import ExcelGenerator

class UniversityResultAutomator:

    def __init__(self):
        self.start_register = None
        self.end_register = None
        self.semester = None
        self.exam_code = None
        
        self.register_numbers = []
        self.results = []

    # Get user input for register number and semester
    def get_user_input(self):
        self.start_register = input("Enter the starting Register Number: ")
        self.end_register = input("Enter the ending Register Number: ")
        try:
            self.semester = int(input("Enter your Semester (1-8): "))
        except ValueError:
            print("semester must be an integer. Please enter a number between 1 and 8.")
            return False
        
        if self.semester not in SEMESTER_EXAM_CODES:
            print("Invalid semester. Please enter a number between 1 and 8.")
            return False

        self.exam_code = SEMESTER_EXAM_CODES[self.semester]
        return True

    # Generate a list of register numbers based on the user input
    def generate_register_numbers(self):
        start_match = re.match(r"^(.*?)(\d+)$", self.start_register)
        end_match = re.match(r"^(.*?)(\d+)$", self.end_register)

        if not start_match or not end_match:
            print("Invalid register number format. Please ensure they end with numbers.")
            return False

        start_prefix = start_match.group(1)
        end_prefix = end_match.group(1)

        start_number_text = start_match.group(2)
        end_number_text = end_match.group(2)

        if start_prefix != end_prefix:
            print("Register number prefixes do not match. Please ensure they have the same prefix.")
            return False

        start_number = int(start_number_text)
        end_number = int(end_number_text)

        if start_number > end_number:
            print("Starting register number must be less than or equal to the ending register number.")
            return False

        number_length = len(start_number_text)

        for number in range(start_number, end_number + 1):
            register_number = (start_prefix + str(number).zfill(number_length)).upper()
            self.register_numbers.append(register_number)
        print("\nTotal Register Numbers: ",len(self.register_numbers))
        return True 

    # IF WE ARE NOT USING MULTI THREADING WE DONT NEED THIS FUCTION :
    # BECAUSE FOR LOOP IS SEQUENTIAL AND BUT MULTI THREADING IS CONCURRENT 

    # def process_one_student(self, register_number):
    #     try:
    #         fetcher = ResultFetcher(register_number, self.exam_code)
    #         response = fetcher.fetch_result()

    #         if not response:
    #             print(f"Failed to fetch result for {register_number}. Skipping...")
    #             return None
            
    #         if response.get("error_code") == -1:
    #             print(register_number, "->", response.get("data", {}).get("html", "Result not available."))
    #             return None
            
    #         html = response["data"]["html"]
    #         parser = ResultParser(html)
    #         student = parser.parse_student_info()
    #         subjects = parser.parse_subjects()

    #         if student is None:
    #             print(f"Student details not found for {register_number}.")
    #             return None
            
    #         if not subjects:
    #             print(f"Subjects details not found for {register_number}.")
    #             return None

    #         return {
    #             "student": student,
    #             "subjects": subjects
    #             }
        
    #     except Exception as error:
    #             print(f"An error occurred while processing {register_number}: {error}")




    # Fetch and Parse all results for the generated register numbers
    def process_students(self):
        # total = len(self.register_numbers) # USED TO DISPLAY THE PROGRESS
        # print(f"\n Processing {total} students... ")

        # #Number of threads
        # with ThreadPoolExecutor(max_workers = 2) as executor:
        #     futures = {
        #         executor.submit(self.process_one_student, register_number): register_number
        #         for register_number in self.register_numbers
        #     }

        #     #tqdm progress bar
        #     with tqdm(
        #         total = total,
        #         desc = "Fetching Results",
        #         unit = "student"
        #     ) as progress_bar:
        #         for future in as_completed(futures):
        #             register_number = futures[future]
        #             try:
        #                 result = future.result()
        #                 if result:
        #                     self.results.append(result)
        #             except Exception as error:
        #                 print(f"\nError processing " f"{register_number}: "f"{error}") 

        #             progress_bar.update(1)

        # return True
                    

        total = len(self.register_numbers) # USED TO DISPLAY THE PROGRESS

        for index, register_number in enumerate(self.register_numbers, start=1):
            print(f"\n[{index}/{total}] Processing Register Number: {register_number}")

            #create a ResultFetcher instance for each register number
            fetcher = ResultFetcher(register_number, self.exam_code)
            response = fetcher.fetch_result()

            # No response
            if not response:
                print(f"Failed to fetch result for {register_number}. Skipping...")
                continue

            # University returned an error
            if response.get("error_code") == -1:
                print(register_number, "->", response.get("data", {}).get("html", "Result not available."))
                continue

            try:
                html = response["data"]["html"]
                parser = ResultParser(html)
                student = parser.parse_student_info()
                subjects = parser.parse_subjects()

                if student is None:
                    print(f"Student details not found for {register_number}.")
                    continue
                if not subjects:
                    print(f"Subjects details not found for {register_number}.")
                    continue

                self.results.append({
                    "student": student,
                    "subjects": subjects
                })

                print(f"Result fetched and parsed successfully for {register_number}.")

            except Exception as error:
                print(f"An error occurred while processing {register_number}: {error}")

            time.sleep(0.5) # delay between requests to avoid overwhelming the server

        return True

    def generate_excel(self):
        if not self.results:
            print("No results to generate Excel. Please fetch results first.")
            return 
        os.makedirs("output", exist_ok=True)
        file_name = f"output/University_Results_{self.semester}.xlsx"

        excel_generator = ExcelGenerator(file_name)
        excel_generator.create_excel(self.results, self.semester)

        print("\nExcel saved successfully at:", file_name)

    def run(self):
        if not self.get_user_input():
            return

        if not self.generate_register_numbers():
            return

        self.process_students()
        print("\nSuccessfully Fetched :",len(self.results),)
        print("Failed/Unavailable :",len(self.register_numbers)-len(self.results))

        self.generate_excel()

        print("\nProcess completed successfully!")            

if __name__ == "__main__":
    app = UniversityResultAutomator()
    app.run()