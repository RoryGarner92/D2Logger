Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\t\testfile.txt", True)
a.WriteLine("This is a test.")
a.Close