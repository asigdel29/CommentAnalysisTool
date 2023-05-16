# CommentAnalysisTool

A Python script designed to process Word documents and extract comments made by students.It streamlines the process of analyzing student comments and participation in academic or research settings. By providing detailed metrics, it enables educators and researchers to gain valuable insights into student engagement and involvement in discussions or collaborative activities.

Features:

Extracts comments: The tool utilizes the textract library to extract text from Word documents, focusing on comments made by students within square brackets, such as "[Student Name] Comment text".

Calculates word count and comment count: It counts the number of words and comments made by each student, considering comments as distinct contributions. It excludes section titles from the comment count to ensure accurate results.

Supports sorting and organizing: The script sorts the Word files in the provided folder in a natural order, ensuring the correct sequencing of files. It also organizes the results by student name and provides a breakdown of counts for each student and document.

Generates a CSV report: The tool generates a CSV file named "analysis.csv" that summarizes the analysis results. The report includes student names, total word count, total comment count, and individual counts for each document.
