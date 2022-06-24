using RPA_Tech.Models;

Notepad notepad = new Notepad();

string notepadTextInput = "Hello World";
notepad.inputNotepadText(notepadTextInput);

notepad.saveNotepad();
Thread.Sleep(500);

notepad.SaveAsWindowInput = notepad.getSaveAsTextBox();
string fileNameToSave = "HelloWorld";
notepad.saveFile(fileNameToSave);

Thread.Sleep(500);
notepad.NotepadProcess.Kill();
Console.WriteLine("Operation Complete. Press any key to exit.");