import openpyxl
from datetime import datetime

print ( "Welcome To Bidaa AL Mutwaa School ")
print("Visitor data is protected in accordance with Bidaa Al Mutawa School's digital safety policy.")
print("Done by the programmers team  at Bada Al Mutawa School ")
print("-------------------------------------------------------")
def record_visitors():
    """Records visitor information (name, job title, reason for visit, entry time) and saves it to an Excel file and a text file."""

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write header row
    sheet.append(["Visitor Name", "Job Title", "Reason for Visit", "Entry Time"])

    with open("visitor_log.txt", "w") as file:
        file.write("Visitor Name | Job Title | Reason for Visit | Entry Time\n")
        file.write("-" * 70 + "\n")

    while True:
        visitor_name = input("Enter visitor name (type 'exit' to quit): ")

        if visitor_name.lower() == "exit":
            break

        job_title = input("Enter visitor job title: ")
        reason_for_visit = input("Enter reason for visit: ")
        entry_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append([visitor_name, job_title, reason_for_visit, entry_time])

        with open("visitor_log.txt", "a") as file:
            file.write(f"{visitor_name:15} | {job_title:20} | {reason_for_visit:25} | {entry_time}\n")

    # Save the workbook
    workbook.save("visitor_log.xlsx")

    print("Visitor information recorded successfully.")

    # Print the contents of the text file
    with open("visitor_log.txt", "r") as file:
        print(file.read())

if __name__ == "__main__":
    record_visitors()