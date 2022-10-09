from sendEmail import *
import subprocess
import os

def main():
    Git_Dir = ""
    dir_path = os.path.dirname(os.path.realpath(__file__))
    with open("input.txt", "r") as obj:
        lines = obj.readlines()
    if(len(lines) > 0):
            for line in lines:
                if "Git_Dir" == line[:7]:
                     Git_Dir = line.split('"')[1]
    os.chdir(Git_Dir)
    for x in range(6):
        cmd_run = subprocess.Popen("git pull PS5 dev",shell=True,stdout=subprocess.PIPE)
        subprocess_return = cmd_run.stdout.read()
        subprocess_return_string = subprocess_return.decode("utf-8", "ignore") 
        if "Already up to date" in subprocess_return_string:
            os.chdir(dir_path)
            accout_emali = Account_Email()  
            accout_emali.send_all()
            break
    
if __name__ == "__main__":
    main()