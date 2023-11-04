# WordFileTranslator
- This is a program by Can SARMAN that allows you to translate Word documents from one language to another with ease and accuracy.
- It uses the Google Translate API to perform the translation.
- It has a graphical user interface (GUI) that lets you select a Word file to be translated, choose the source and destination languages, and start the translation process.
- It also shows the progress of the translation, the path of the translated file, and the path of the log file that records any errors or exceptions that occur during the translation.
- It creates a copy of the original Word file and replaces the text in each paragraph with the translated text.
- It uses a chunk size of 1024 characters to avoid exceeding the Google Translate API limit.
- It allows you to cancel the translation at any time and deletes the partially translated file.
- It handles various errors and exceptions that may occur while opening, saving, or translating the Word file and displays appropriate messages to the user.
- It uses the docx module to manipulate Word documents and the tkinter module to create the GUI.
