# excel-webpage-printscreen

this python script search-screenshot.py: 

- gets search parameters from one excel file 
- makes an url for the google search pages using the words in the excel file
- search for example "personname" AND "word1" OR "word2" OR ... OR "word-n"
- opens google search pages on webbrowser tabs and takes screenshots of them. (screenshots will be taken of the whole screen, so open a webbrowser window to full the screen)
- opens a catpicture page to let you know the script is finished with the screenshots

- saves results as pictures in a results subfolder
- saves 1 excel for each person name telling what search was done and when

- this script has almost no error handling (yet). It requires an input at the end, so you know that if you get the line asking for the end input, the program didn't crash. If you don't get it, something went wrong.
