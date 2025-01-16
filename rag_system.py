import pandas as pd
import openai

class EmployeeDataProcessor:
    """
    A class to process employee data from an Excel file
    and generate summaries using OpenAI's GPT API.
    """
    def __init__(self, excel_path):
        """
        Initializes the EmployeeDataProcessor with the Excel file path.
        """
        self.excel_path = excel_path
        self.data = None
        self.summary = None

    def read_excel(self):
        """
        Reads the Excel file into a pandas DataFrame.
        """
        try:
            # Read the Excel file
            self.data = pd.read_excel(self.excel_path)
            print("Data successfully loaded.")
        except Exception as e:
            print(f"Error reading Excel file: {e}")

    def process_data(self):
        """
        Processes the data to compute average salary and department distribution.
        """
        if self.data is not None:
            try:
                # Calculate average salary
                avg_salary = self.data['Salary'].mean()
                # Count employees in each department
                department_distribution = self.data['Department'].value_counts().to_dict()

                # Store the summary
                self.summary = {
                    "average_salary": avg_salary,
                    "department_distribution": department_distribution
                }
                print("Data successfully processed.")
            except KeyError as e:
                print(f"Error processing data. Missing column: {e}")
        else:
            print("No data to process. Please load the Excel file first.")

   def summarize_with_gpt(self, api_key):
    """
    Generates a natural language summary using OpenAI's GPT API.
    """
    if self.summary is not None:
        # Prepare the GPT prompt
        prompt = (
            f"The data contains employee details. "
            f"The average salary is {self.summary['average_salary']:.2f}. "
            f"The department distribution is as follows: {self.summary['department_distribution']}. "
            f"Please generate a summary in plain English."
        )
        
        # Set OpenAI API key
        openai.api_key = api_key
        try:
            # Call OpenAI GPT API
            response = openai.Completion.create(
                engine="text-davinci-003",
                prompt=prompt,
                max_tokens=150
            )
            # Return the text response
            return response.choices[0].text.strip()
        except Exception as e:
            print(f"Error using GPT API: {e}")
    else:
        print("No summary available. Please process the data first.")
    return None


if __name__ == "__main__":
    # Step 1: Initialize the processor
    processor = EmployeeDataProcessor("employee_data.xlsx")  # Replace with your Excel file path
    
    # Step 2: Read Excel file
    print("Reading Excel file...")
    processor.read_excel()
    
    # Step 3: Process the data
    print("Processing data...")
    processor.process_data()
    
    # Step 4: Generate summary using GPT
    print("Generating summary with GPT...")
    api_key = "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"  # Replace with your OpenAI API key
    summary = processor.summarize_with_gpt(api_key)
    
    # Step 5: Print the generated summary
    if summary:
        print("\nGenerated Summary:")
        print(summary)
