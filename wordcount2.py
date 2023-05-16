import os
import csv
import re
import textract
from natsort import natsorted

# Prompt the user to enter the folder path
folder_path = input("Enter the folder path containing the Word documents: ")

# Get a list of all the Word documents in the folder
word_files = [f.strip() for f in os.listdir(folder_path) if f.endswith('.docx')]

# Print the list of word files
print("\nGiven Files:")
for file in word_files:
    print(file)

# Sorting function
# def sort_filenames(filename):
#     parts = re.split('-|\s', filename, maxsplit=2)
#     first_num = int(parts[0]) if len(parts) > 0 and parts[0].isdigit() else 0
#     second_num = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
#     return (first_num, second_num)

# Sort the word_files list using natsort
sorted_files = natsorted(word_files)

# Print the list of word files after sorting
print("\nFiles sorted:")
for file in sorted_files:
    print(file)

# Initialize a dictionary to store the results for each student
results = {}

# Process each Word document in the folder
for file_name in sorted_files:
    # Get the full path of the Word document
    file_path = os.path.join(folder_path, file_name)

    # Extract the text from the Word document
    text = textract.process(file_path).decode('utf-8')

    # Identify the students who wrote comments
    student_names = re.findall(r'\[(.*?)\]', text)

    # Loop through each student name and count their comments
    for student_name in student_names:
        # Replace line breaks that aren't at the end of a sentence with a space
        text = re.sub(r'(?<!\.)\n', ' ', text)
        # Check if this is the first time we're seeing this student name
        if student_name not in results:
            results[student_name] = {
                'word_count': 0,
                'comment_count': 0,
                'files': {file_name: {'word_count': 0, 'comment_count': 0}}
            }

        # Define a regular expression for the section titles
        section_title_regex = r'Discussion Points\s*\(at least one per person\)\s*'

        # Define a regular expression for the comments
        comment_regex = f'{student_name}\]\s*(.*?)(?=\[|$)'

        # Extract the comments
        comments = re.findall(comment_regex, text, re.DOTALL)

        # Exclude the section titles from the comments
        comments = [re.sub(section_title_regex, '', comment).strip() for comment in comments]

        # Count the number of comments made by the student and total number of words
        word_count = 0
        comment_count = 0
        for comment in comments:
            words = comment.split()
            if len(words) > 0:  # only count the comment if it contains words
                comment_count += 1
                word_count += len(words)
                # for i, word in enumerate(words, start=1):
                #      print(f"{i}.{word}")

        # Count the total number of words written by the student
        word_count = sum(len(comment.split()) for comment in comments)
        # for comment in comments:
        #     words = comment.split()
        #     for i, word in enumerate(words, start=1):
        #         print(f"{i}.{word}")
        #     word_count = len(words)

        # Add the results for this student to the dictionary
        results[student_name]['word_count'] += word_count
        results[student_name]['comment_count'] += comment_count
        results[student_name]['files'][file_name] = {
            'word_count': word_count,
            'comment_count': comment_count
        }

# Sort the student names alphabetically
student_names = sorted(results.keys())

# Create a CSV file to store the results
with open('analysis.csv', mode='w', newline='') as csv_file:
    writer = csv.writer(csv_file)

    # Write the header row
    header_row = ['Student Name', 'Total Word Count', 'Total Comment Count']
    for file_name in sorted_files:
        header_row.append(f'{file_name} Count')
    writer.writerow(header_row)

    # Write the data rows
    for student_name in student_names:
        total_word_count = 0
        total_comment_count = 0
        row = [student_name]
        for file_name in sorted_files:  # use sorted_files instead of word_files
            if file_name in results[student_name]['files']:
                file_results = results[student_name]['files'][file_name]
                total_word_count += file_results['word_count']
                total_comment_count += file_results['comment_count']
                # Combine word count and comment count into a single cell, separated by a comma
                row.append(f"{file_results['word_count']}, {file_results['comment_count']}")
            else:
                row.append('')
        row.insert(1, total_word_count)
        row.insert(2, total_comment_count)
        writer.writerow(row)

print(f"\nNumber of students: {len(results)}")
print("Analysis complete. Results written to analysis.csv.")
