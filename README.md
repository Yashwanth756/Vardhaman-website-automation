<h1>About the project</h1>
<p> 
The project is specifically tailored for faculty at Vardhaman College of Engineering, streamlining the process of extracting crucial student information such as CGPA, SGPA, backlogs, and other academic data. This system offers an intuitive and efficient solution, saving time and effort by automating data retrieval and presenting it in a user-friendly interface.
<br><br>This project uses these package for running :<br>
Selenium, a powerful web automation tool, to facilitate the login process using student credentials and extract essential data from a designated website. This allows for seamless interaction with web elements and automated navigation.

To handle data efficiently, **Pandas** is utilized for reading from and writing to Excel files (.xlsx), enabling easy manipulation and analysis of the extracted information.

The user experience is enhanced through the implementation of **Tkinter**, a standard GUI toolkit for Python, which creates an intuitive and visually appealing graphical user interface. This interface simplifies user interactions, making the functionality of the application more accessible and user-friendly.
</p>
<h3>Running the file</h3>
<p>To run the program use the command python day2final.py</p>
<p>Wait until the  tkinter window appears
<ol>
 <li>Give the output file name which will create the xlsx file in the current folder with given name</li> 
  <li>Click on choose file button - It pops up a file dialogue menu select a xlsx file which has students credentials as specified the format of datag.xlsx file</li>
  <li>Click on submit button</li>
</ol>
</p>
<h3>OUTPUT</h3>
<b>After the completion of program it outputs 3 files</b>
<ol>
  <li>fileNameMarks.xlsx</li>
  <li>fileNameMarks_backlogs_code_list.xlsx</li>
  <li>fileNmaeMarks_backlogs_SubjectName_list.xlsx</li>
</ol>

<p><B>fileNameMarks.xlsx</B> -  It consits of 
<ol>
  <li>rollNo</li>
  <li>sgpa 1</li>
  <li>remaining sgpa</li>
  <li>cgpa</li>
  <li>Backlogs count</li>
  <li>Backolgs subject code</li>
  <li>Backlogs subject name</li>
</ol>
</p>

<p><B>fileNameMarks_backlogs_code_list.xlsx</B> - It consits of 
<ol>
  <li>Backlogs code</li>
  <li>Backloge count for each subject</li>
</ol>
</p>

<p><b>fileNmaeMarks_backlogs_SubjectName_list.xlsx</b> - It consists of 
<ol>
  <li>Subject name</li>
  <li>Subject backlog count</li>
</ol>
</p>
