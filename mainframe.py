import HierarchyMaker
from MappingNonLFF import mappingnonlffaktion
from MappingLFFRobot import mainloop
from test import testing
import openpyxl as xl
import subprocess
import os
from pathlib import Path
import win32com.client
import pythoncom


def get_user_choice():
    while True:
        try:
            ans = int(input("\n#####################################################\n\nChoose what u wanna do: \n1. create hierarchy, \n2. create nonLFF mapping template, or \n3. use the LFF mapping rob0t, or \n4. exit cuz im a dummy \n5. git pull (pls do this everytime i beg you) \n6. git coomit/push \n7. open GCOH Mapping Log \n8. Open GCOH Structure file \n9. Open Hierarchy Builder \n\n(ENTER AN INTEGER FOR THE LOVE OF GOD) \n#######################################################\n"))
            if 1 <= ans <= 9:
                return ans
            else:
                print("Invalid input. Please enter a number between 1 and 9.")
        except ValueError:
            print("Please ENTER an INTEGER.")


def run_hierarchy_maker():
    try:
        os.system('cls')
        print("running hierarchy maker")

        pandas_path = HierarchyMaker.INPUT_FILE

        # Attach to existing Excel or start a new one
        excel_was_running = True
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except:
            excel_was_running = False
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = True
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(pandas_path))

        print(
            f"\nEdit {pandas_path.name} as needed (or paste in a fresh SAP download)")
        print("Type 'Y' here when you're done to save and continue.\n")

        while True:
            if input("> ").strip().upper() == "Y":
                break
            print("Type 'Y' when you're done editing.")

        print(f"Saving and closing {pandas_path.name}...")

        wb.Save()
        wb.Close(SaveChanges=False)

        if not excel_was_running:
            excel.Quit()

        print("Excel file saved and closed.")
        print("Running Hierarchy Maker...")

        output_file = HierarchyMaker.hiymaker()
        print("Hierarchy Maker run successfully!")

        os.startfile(output_file)

    except Exception as e:
        print(f"Error running hierarchy maker: {e}")


def run_nonlff_template_creator():
    try:
        os.system('cls')
        print("running nonLFF template creator")

        deathpath = Path('Inputs/PyDump.xlsx').resolve()
        outpath = Path('Inputs/Out.csv')

        excel_was_running = True
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except:
            excel_was_running = False
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = True
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(deathpath))

        print("\nEnter in the cost objects and the relevant details")
        print("Type 'Y' here when you're done to save and continue.\n")

        while True:
            if input("> ").strip().upper() == "Y":
                break
            print("Type 'Y' when you're done editing.")

        print("Saving and closing workbook...")

        wb.Save()
        wb.Close(SaveChanges=False)

        if not excel_was_running:
            excel.Quit()

        print("Excel file saved and closed.")
        print("Proceeding with mapping...")

        mappingnonlffaktion()
        print("Running now please wait for the CSV file to come out")
        os.startfile(outpath)

    except Exception as e:
        print(f"Error creating nonLFF template: {e}")


def run_lff_robot():
    try:
        os.system('cls')
        print("running LFF r0bot")
        mainloop()
        print("LFF robot run successfully!!!")
    except Exception as e:
        print(f"Error running LFF robot: {e}")


def open_gcohmapping():
    try:
        os.startfile("F3 2024 mappings.xlsx")
    except Exception as e:
        print(f"Where are you running this shit from bro")


def open_gcohstruct():
    try:
        os.startfile("GCOH Structure.xlsx")
    except Exception as e:
        print(f"Where are you running this shit from bro")


def open_pandascreator():
    try:
        os.startfile(f"inputs/pandas.xlsx")
    except Exception as e:
        print(f"Where are you running this shit from bro")


def run_git_commands(choice):
    if choice == 5:
        os.system('cls')
        subprocess.run(["git", "pull"])
        print("GIT pulled please check if there are any erros")
    elif choice == 6:
        os.system('cls')
        subprocess.run(["git", "commit", "-a", "-m", "new job"])
        subprocess.run(["git", "push"])
        print("GIT pushed please check if there are any erros")


def main():
    global x
    while True:
        x = get_user_choice()
        if x == 4:
            os.system('cls')
            print("thank u for using my pr0gram now fck off dummy")
            break
        elif x == 1:
            run_hierarchy_maker()
        elif x == 2:
            run_nonlff_template_creator()
        elif x == 3:
            run_lff_robot()
        elif x == 5:
            run_git_commands(5)
        elif x == 6:
            run_git_commands(6)
        elif x == 7:
            open_gcohmapping()
        elif x == 8:
            open_gcohstruct()
        elif x == 9:
            open_pandascreator()


if __name__ == "__main__":
    main()
