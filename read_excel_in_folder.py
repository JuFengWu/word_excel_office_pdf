import  os

# get current path
currentPath = os.path.dirname(os.path.abspath(__file__))
print(currentPath)

excelList = []
# get all files in foler
for path, subdirs, files in os.walk("excel_data\execl_data1"):
    for name in files:
        subFoilderFilePath = os.path.join(path, name)
        print(subFoilderFilePath)
        root, extension = os.path.splitext(subFoilderFilePath)
        print("root is " + root + ", extension is "+str(extension))
        if extension == ".xlsx":
            excelList.append(subFoilderFilePath)
print(excelList)

