from binascii import a2b_hex
import csv
from pprint import pprint
from random import choice


def clean_age(age_range: str):
    if "+" in age_range[0]:
        age_list = [
            age_range[0].strip(),
        ]
        age_list.append(age_range[1].strip())
    else:
        age_list = age_range.split()

    n1, n2 = [int(x) for x in age_list[0].split("-")]

    random_age = choice([i for i in range(n1, n2 + 1)])

    return " ".join([str(random_age), age_list[1]])


def hide_age(row):
    age_hidden_dict = {}
    with open("dictionary/age_hidden.csv") as read_obj:
        csvreader = csv.reader(read_obj)
        for line in csvreader:
            if "+" in line[1]:
                age_hidden_dict[int(line[0])] = int(line[1].strip("+"))
            else:
                age_range = line[1].split("-")
                age_hidden_dict[int(line[0])] = [int(x) for x in age_range]

        print(age_hidden_dict)
        try:
            if isinstance(age_hidden_dict[row], list):
                return choice(
                    [
                        x
                        for x in range(
                            age_hidden_dict[row][0], age_hidden_dict[row][1] + 1
                        )
                    ]
                )
        except KeyError:
            return choice([x for x in range(age_hidden_dict[int(row)], 119)])


def convert_height_to_inches(string):
    stripped = string.strip().split(" ")
    try:
        n = int(stripped[0].replace("'", ""))
    except ValueError:
        pass

    try:
        n1 = int(stripped[1].replace('"', ""))
    except:
        pass
    try:
        n1 = int(stripped[0].replace('"', ""))
        return n1

    except:
        pass

    total_inches = (n * 12) + n1

    return total_inches
