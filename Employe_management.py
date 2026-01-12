import os
from openpyxl import Workbook, load_workbook
import msvcrt

class empDB:
    def __init__(self, filename, headers):
        self.filename = filename

        if not os.path.isfile(filename):
            self.wb = Workbook()
            self.ws = self.wb.active
            self.ws.append(headers)
            self.headers = headers
            self.wb.save(filename)
        else:
            self.wb = load_workbook(filename)
            self.ws = self.wb.active
            self.headers = [cell.value for cell in self.ws[1]]

    def add(self, empdata):
        for i in empdata:
            if i == "":
                print("Please provide enough data")
                print(headers)
                return
        
        self.ws.append(empdata)
        self.wb.save(self.filename)
        print("Data Added")

    def find(self, ID):
        for row in self.ws.iter_rows(min_row=1, values_only=True):
            if row[0] == ID:
                print(row)
                return
        print(f"The data for 'EMPID:{id}' is not present in database")

    def modify(self, ID):
        modData = ["" for _ in range(len(headers)-1)]
        count = 1
        for row in self.ws.iter_rows(min_row=1,):
            if row[0].value == ID:
                print("Leave Empty if not want to modify")
                modData[0] = input("Enter New Name: ")
                modData[1] = input("Enter New Dept: ")
                modData[2] = input("Enter New Sal: ")
                for data in modData:
                    if data == "":
                        count +=1
                        continue    
                    row[count].value = data
                    count +=1
                self.wb.save(self.filename)
                print("Data Modified")
                return
        print(f"Data for 'EMPId:{id}' not present in database")
        return

    def delete(self, ID):
        for row_num, row in enumerate(self.ws.iter_rows(min_row=2), start=2):
            if row[0].value == ID:
                print([cell.value for cell in row])
                choice = input("Are you sure do you want to delete this data Y/N: ")
                if "y" == choice.lower():
                    self.ws.delete_rows(row_num)
                    print("Data Deleted")
                else:
                    print("Data not deleted")
                self.wb.save(self.filename)
                return
        print(f"The data for 'EMPID:{id}' is not present in database")
        return



if __name__ == "__main__":
    headers = ["EMPID","EmpName","Dept","Salary"]
    while True:
        print("1. Add Employee")
        print("2. Find Employee")
        print("3. Modify Employee")
        print("4. Delete Employee")
        print("5. Exit")

        db = empDB("Employee_databse.xlsx", headers=headers)

        ch = msvcrt.getch().decode()
        match ch:
            case "1":
                data = dict.fromkeys(headers, "")
                for i in data.keys():
                    d = input(f'Enter {i} : ')
                    data[i] = d
                db.add(list(data.values()))
            case "2":
                id = input("Enter EmpID: ")
                db.find(id)
                
            case "3":
                id = input("Enter EmpID: ")
                db.modify(id)
    
            case "4":
                id = input("Enter Empid: ")
                db.delete(id)

            case "5":
                break
