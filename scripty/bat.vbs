Set fs = CreateObject("Scripting.FileSystemObject")
' Set a = fs.OpenTextFile("C:\Users\Rory\Desktop\D2Logger\new.txt", 8, true)

Set a = fs.OpenTextFile("C:\Users\Rory\Desktop\D2Logger\scripty\testfile.txt", 8 , True)
a.WriteLine("This is a test.")
a.Close