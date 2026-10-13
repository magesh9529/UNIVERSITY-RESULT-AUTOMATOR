import requests
from config import BASE_URL, RESULT_API_URL, SEMESTER_EXAM_CODES


class ResultFetcher:

    def __init__(self, register_number, exam_code):
        self.register_number = register_number
        self.exam_code = exam_code

        # create a session - automatically stores cookies received from the first request and sends them with the second request.
        self.session = requests.Session()

    
    def fetch_result(self):

        # Data needed for the requests
        params = {
            "r" : self.register_number,
            "e" : self.exam_code,
        }
        data = {
            "r" : self.register_number,
            "e" : self.exam_code,
            "ct" : ""
        }

        try:
            # STEP 1: Open normal result page
            page_response = self.session.get(
                BASE_URL, 
                params = params,
                timeout=15
                )
            page_response.raise_for_status()  # Raise an exception for HTTP errors
            
            # return response.json()  # Return the JSON content of the result page
            # return response.text  # Return the HTML content of the result page

            # STEP 2: Call actual result API

            headers = {
                "X-Requested-With" : "XMLHttpRequest", # This header is often used to identify Ajax requests. It can be used by the server to return different content for Ajax requests.
                "Referer" : page_response.url  # set the Referer header to the URL of the first request
            }
            response = self.session.post(
                RESULT_API_URL,
                data = data,
                headers = headers,
                timeout=15
            )
            response.raise_for_status()  # Raise an exception for HTTP errors
            
            # STEP 3: Convert JSON response to Python dictionary and return it
            result =  response.json()  # Return the JSON content of the result page
            return result
        
        except requests.RequestException as error:
            print(f"Error fetching result: {error}")
            return None

if __name__ == "__main__":

    result = ResultFetcher("23TN0047", SEMESTER_EXAM_CODES[6])
    fetched_result = result.fetch_result()

    if fetched_result:
        print("Result fetched successfully!")
        # You can process the fetched_result further here
        print(fetched_result) #[:2000] # Print the first 2000 characters of the result for demonstration

    else:
        print("Failed to fetch the result.")

'''
#URL Construction using "f" Strings
register_number = "23TN0047",
semster_number = 6
semester_code_mapping = {    
    1 : "A",
    2 : "B",
    3 : "C",
    4 : "D",
    5 : "E",
    6 : "F",
    7 : "G",
    8 : "H"
}
semseter_code = semester_code_mapping[semster_number]
url = f"https://exam.pondiuni.edu.in/results/result.php?r={register_number}&e={semseter_code}"
'''