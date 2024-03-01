Face Recognition Attendance System

This Python script is designed to recognize faces in a webcam feed and mark attendance automatically in an Excel sheet. It utilizes the face_recognition library to recognize faces and OpenCV for video capture and display.

Features
Face Recognition: The script uses the face_recognition library to recognize faces in real-time webcam feed.
Attendance Marking: Automatically marks attendance in an Excel sheet for recognized faces.
Date and Time Logging: Logs the date and time of attendance alongside student names.
Installation
Clone this repository to your local machine:

bash
Copy code
git clone https://github.com/q-ms8/Attendance-System-Using-Face-Recognition.git
Install the required dependencies:

Copy code
pip install opencv-python numpy pandas openpyxl face-recognition
Usage
Ensure that your webcam is connected and functional.

Run the Python script:

Copy code
python face_recognition_attendance.py
The webcam feed will open, and faces will be recognized in real-time.

Press 'q' to exit the application.

Configuration
Image Directory: Images of students should be stored in the ImagesAttendance directory.
Excel Sheet: The attendance will be marked in the attendance_sheet.xlsx Excel file. Make sure it exists in the same directory as the script.
Troubleshooting
If the script fails to recognize faces, ensure that your webcam is properly configured and well-lit.
Make sure the images of students are clear and well-captured for better recognition accuracy.
Contributing
Contributions are welcome! If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.


