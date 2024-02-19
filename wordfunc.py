class fun():
    def __init__(self):
        self.fun=""
        self.funtab=["IF","OR","ROUNDUP","SUM"]
    def check(self,word):
        if word[0:1] is "=":
            return True
        else:
            return False
    def function(self,functionn):
        self.fun=functionn
        pom1=str(fun)
        try:
            pom2 = pom1.rfind("=")
            subfun = fun[pom2 + 1:-1]
        except:
            print("",end="")

        if subfun[0]=="I":
            return "IF"
        if subfun[0]=="O":
            return "OR"
        if subfun[0]=="R":
            return "ROUNDUP"
        if subfun[0]=="SUM":
            return "SUM"