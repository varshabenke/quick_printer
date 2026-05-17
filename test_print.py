import win32api

file_path = r"C:\Users\VarshaBenke\Desktop\dl.jpeg"

win32api.ShellExecute(
    0,
    "print",
    file_path,
    None,
    ".",
    0
)

print("Printing Started")