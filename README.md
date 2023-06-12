# Automated Study Planner using Pomodoro Technique

This project focuses on developing an automated study planner that incorporates the Pomodoro technique and tracks study time across various subjects. It leverages data visualization tools like PivotTable to effectively present study information. Furthermore, data analysis is conducted on the study log data to gain valuable insights.


## Features
- Automatic study plan generation based on the Pomodoro technique
- Recording and tracking study time for each subject
- Visualization of study progress through tables and graphs
- Flexibility to study without following the Pomodoro technique
- Easy input of start and end times for study sessions
- Handling of breaks between study intervals and sessions

## Spreadsheet Structure

The study planner utilizes a spreadsheet with the following column names:

1.	DATE: Date of the study session
2.	STARTING TIME: Start time of the session
3.	ENDING TIME: End time of the session
4.	DURATION: Duration of the session
5.	COURSE: Subject or course being studied
6.	RESULT: Outcome of the study session (+ for completed, - for incomplete)

## Pomodoro Technique

In the Pomodoro technique, each study session consists of 4 intervals, with each interval being 25 minutes long. The breaks between intervals are as follows:

- 5-minute break between the first three intervals
- 7-minute break between the third and fourth intervals

If the user wishes to have multiple study sessions, a 15-minute break is provided between sessions. However, if the user prefers to study without following the Pomodoro technique, they can still use the automated study planner.

## Usage

To use the automated study planner with the Pomodoro technique, follow these steps:

1. First, we press Ctrl+J keys, and activate the macro that organizes the study plan using the Pomodoro technique.


2. Enter the desired starting time and the number of study sessions in the program.


![image](https://github.com/romka516/SmartPomodoroPlanner/assets/101732278/d1355c07-eea6-41ce-8ad5-4a7207e9eebb)
![image](https://github.com/romka516/SmartPomodoroPlanner/assets/101732278/bbc8262d-2132-4634-8749-8740ce0e4dce)



3. The program will automatically fill in the date, break times between intervals, and break times between sessions.


![image](https://github.com/romka516/SmartPomodoroPlanner/assets/101732278/ad025769-07a6-473f-82b4-291787f19baf)



4. Fill in the RESULT column with "+" for successfully completed sessions and "-" for incomplete sessions.


![image](https://github.com/romka516/SmartPomodoroPlanner/assets/101732278/67e59a2b-fe4a-4998-9f45-1a213dccb3d6)


However, If you wish to use automated study planner without Pomodor technique, follow these steps:

1. Press Ctrl+P keys to activate the Macro


2. Enter the desired starting time

![image](https://github.com/romka516/SmartPomodoroPlanner/assets/101732278/abc06f4c-efcd-459e-a90b-e2beb39517f4)



3.Single-row will be created as in the image
![image](https://github.com/romka516/SmartPomodoroPlanner/assets/101732278/2f5a380d-44a5-4d28-af0c-2e5264e1ffe5)


Upon completing your study session, please fill in the course you studied and the time at which you finished studying.


## License

This project is licensed under the [MIT License](LICENSE).
