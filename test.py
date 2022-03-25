import win32print

# hello = win32print.OpenPrinter("EPSON L3150 Series")
file = "Print.xlsx"
pHandle = win32print.GetDefaultPrinter()
printer = win32print.OpenPrinter(pHandle)
job = win32print.StartDocPrinter(file, None, "RAW")
# win32print.StartDocPrinter(printer)
win32print.WritePrinter(printer, "Print Me Puhleeezzz!")
win32print.EndPagePrinter(printer)
print(pHandle)
print(printer)