import os
from collections import namedtuple
import csv

def main():
    scan_for_file()


def scan_for_file():
    list = []
    for item in os.listdir('.'):
        if item.endswith(".xlsx"):
            list.append(item)

    if len(list) == 1:
        read_file(list[0])
    else:
        for index, item in enumerate(list):
            print(str(index) + ": " + item)

        while True:
            x = input("which file to run? ")
            if x.isdigit():
                if int(x) < len(list):
                    read_file(list[int(x)])
                    return
            print("input not valid")


def read_file(filename):
    TestItem = namedtuple("TestItem",
                          "Number, Tool, Order, ID, Job, Activity, Index, PRI, Clicks, "
                          "Time, FUN, SIM, PER, STA, SCA, SEC, Comments")
    with open(filename, 'rb') as file:
    	reader = csv.reader(file)
    	for row in reader:
    		    print ' '.join(row)



if __name__ == '__main__':
    main()
