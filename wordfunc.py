class Fun:
    def __init__(self):
        self.fun = ""
        self.funtab = ["IF", "OR", "ROUNDUP", "SUM"]

    def isfunction(self, word):
        if str(word)[0:1] == "=":
            return True
        else:
            return False

    def whatfunction(self, functionn):
        self.fun = functionn
        pom1 = str(self.fun)
        try:
            pom2 = pom1.rfind("=")
            subfun = pom1[pom2 + 1:]
            if subfun.startswith("IF"):
                return "IF"
            elif subfun.startswith("OR"):
                return "OR"
            elif subfun.startswith("ROUNDUP"):
                return "ROUNDUP"
            elif subfun.startswith("SUM"):
                return "SUM"
        except Exception as e:
            print("Error:", e)

    # def cell(self, command):
    #     tab = []
    #     value = 0
    #     for char in command:
    #         if char.isalpha() and char.upper() <= 'Z':
    #             tab.append(ord(char.upper()) - ord('A'))
    #             value=value+1
    #         else:
    #             helpingvar=str(char)
    #             check=0
    #             while check==0:
    #                 try:
    #                     if command[value].isalpha():
    #                         helpingvar=helpingvar+str(command[value])
    #                         value=value+1
    #                 except:
    #                     tab.append(helpingvar)
    #                     check=1
    #
    #             # for j in range(len(command)):
    #             #     if not char.isalpha():
    #             #         helpingvar=char
    #             #         tab.append(char)
    #             #     else:
    #             #         continue
    #     return tab

    def cell(self, command):
        tab = []
        value = 0
        current_number = ""
        while value < len(command):
            char = command[value]
            if char.isalpha() and char.upper() <= 'Z':
                if current_number!="":
                    tab.append(int(current_number))
                    current_number=""
                tab.append(ord(char.upper()) - ord('A')+1)
                value += 1
            elif char.isdigit():
                current_number += char
                value += 1
        if current_number != "":
            tab.append(int(current_number))

        return tab

