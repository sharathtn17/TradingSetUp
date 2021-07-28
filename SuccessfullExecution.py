#to add html extention of the file
#!/usr/bin/python3
import string
file=open("Execution.txt","r")

file2=open("ExecutionParse1.txt","w")

for line in file:
      if "Execution successful" in line:
         
         Prev=Prev.split(":",1)
         Prev=Prev[1]
         print(Prev)
         if ".html" in Prev:
            Prev=Prev.split(".html")
            Prev=Prev[0]+"\n"
            print(Prev+"**")
         file2.write(Prev.replace(" ", ""))
      
      else:
          #print(line)
          Prev=line
        