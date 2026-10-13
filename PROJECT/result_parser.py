from bs4 import BeautifulSoup

class ResultParser:

    def __init__(self,html):
        self.html = html
        self.soup = BeautifulSoup(html, 'html.parser')

    def parse_student_info(self):

        student_tables = self.soup.find_all(
            "table",
            id = "student_info"
        )
        if not student_tables:
            return None # Return None if the table is not found

        # Get all rows in the first table
        rows = student_tables[0].find_all("tr") 

        # Get the exam name from the first row and first cell
        exam_name = rows[0].find_all("td")[0].get_text(" ", strip=True)

        # Get the text of the second row and strip whitespace
        register_text = rows[1].get_text(" ", strip = True) 

        # Get the student name from the second row and first cell
        student_name = rows[2].get_text(" ", strip=True)


        register_number = register_text.split(":", 1)[1].strip()  # Extract the register number from the text

        student_name = student_name.split(":", 1)[1].strip()  # Extract the student name from the text

        return{
            "register_number" : register_number,
            "student_name" : student_name,
            "exam_name" : exam_name
        }

    def parse_subjects(self):
        subject_table = self.soup.find("table", id = "results_subject_table")

        if subject_table is None:
            return None  # Return None if the table is not found

        subjects = []

        rows = subject_table.find_all("tr")

        for row in rows[1:]:
            columns = row.find_all("td")

            # if len(columns)<7:
            #     continue  # Skip rows that don't have enough columns

            subject = {
                "subject_name" : columns[1].get_text(" ", strip =True),
                "type" : columns[2].get_text(" ", strip = True),
                "credit": columns[3].get_text(" ", strip = True),
                "grade_point": columns[4].get_text(" ", strip = True),
                "credit_point": columns[5].get_text(" ", strip = True),
                "grade": columns[6].get_text(" ", strip = True)
            }
            subjects.append(subject)

        return subjects

if __name__ == "__main__":

    from result_fetcher import ResultFetcher
    from config import SEMESTER_EXAM_CODES


    # Fetch result from university

    fetcher = ResultFetcher(
        "23TN0047",
        SEMESTER_EXAM_CODES[6]
    )

    response = fetcher.fetch_result()


    if response and response["error_code"] != -1:

        html = response["data"]["html"]


        # Parse HTML

        parser = ResultParser(html)


        student = parser.parse_student_info()

        subjects = parser.parse_subjects()


        print("\nSTUDENT DETAILS")
        print("----------------")

        print(student)


        print("\nSUBJECT RESULTS")
        print("----------------")


        for subject in subjects:

            print(
                subject["subject_name"],
                "->",
                subject["grade"],
                "->",
                subject["credit"],
                "->",
                subject["grade_point"],
                "->",
                subject["credit_point"]
            )

    else:

        print("Result not available.")
    