import pandas as pd
from datetime import datetime, timedelta
import holidays
import os


def licence():
    print("""\033[1;91m
    *******************************************************************************
                                    LICENCE
                Rotagen by Expergefactor Copyright (c) 2024.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
    *******************************************************************************
    \033[0m""")


def clear_console():
    os.system('cls' if os.name == 'nt' else 'clear')


def banner():
    # https://texteditor.com/multiline-text-art/
    print("""
                                  ━━━━━━━━┏┓━━━━━━━━━━━━━━━━━━
                                  ━━━━━━━┏┛┗┓━━━━━━━━━━━━━━━━━
                                  ┏━┓┏━━┓┗┓┏┛┏━━┓━┏━━┓┏━━┓┏━━┓
                                  ┃┏┛┃┏┓┃━┃┃━┗━┓┃━┃┏┓┃┃┏┓┃┃┏┓┃
                                  ┃┃━┃┗┛┃━┃┗┓┃┗┛┗┓┃┗┛┃┃┃━┫┃┃┃┃
                                  ┗┛━┗━━┛━┗━┛┗━━━┛┗━┓┃┗━━┛┗┛┗┛
                                  ━━━━━━━━━━━━━━━━┏━┛┃━━━━━━━━
                                  ━━━━━━━━━━━━━━━━┗━━┛━━━━━━━━
                            \033[1;92mWorkforce rota generator by Expergefactor\033[0m
    """)


def choices():
    try:
        while True:
            search_type_choice = input(
                " \033[1;97m What would you like to do?\n\n\033[0m"
                " \033[1;93m1\033[1;97m Generate a new rota\n"
                " \033[1;93m2\033[1;97m Find out how rotagen works\n"
                " \033[1;93m3\033[1;97m Exit\n\n\033[0m"
                " \033[1;93mInput a value and press 'enter': \033[0m"
            )
            if search_type_choice == "1":
                clear_console()
                banner()
                rotagen()
                break
            if search_type_choice == "2":
                print("\n\033[1;97m SUMMARY OF FUNCTION...\n"
                      " Rotagen will create a rota to determine who is going to be working on any given day\n"
                      " between specified dates conforming to preset criteria.\n\n"
                      " It produces a simple excel spreadsheet comprising a single sheet with 4 headers:\n"
                      " 'Date', 'Day', 'On Call' & 'Public Holiday'.\n\n"
                      " The date value is the date and will list every date in the UK calendar for the period\n"
                      " requested. The day value is the correct day for the date value. The 'On Call' value is a\n"
                      " person parameter which specifies the person to work on that day. The person parameters are\n"
                      " user defined, comma seperated values. The 'Public Holiday' value will be populated with 'PH'\n"
                      " for every UK public holiday on any given date.\n\n"
                      " The script asks you to:\n"
                      " 1. Specify the person parameters i.e. 'John Doe, Jane Doe,' etc.\n"
                      " 2. The person parameter that was working on the last day of the current rota.\n"
                      " 3. The date that the current rota ends (this will determine the date that the new rota "
                      "begins).\n"
                      " 4. The date that the new rota should end.\n\n"
                      " Rules that will be applied:\n"
                      " * UK based calendar.\n"
                      " * Each person gets allocated days to work as equally and fairly as is possible.\n"
                      " * The person parameter that is allocated to any given Friday also gets allocated the\n"
                      "    Saturday and Sunday that immediately follow.\n"
                      " * The Friday-Sunday blocks of working are spaced as far apart as is possible for\n"
                      "    each person so that each person does not work weekends in quick succession.\n\n")
                licence()
                input("\033[1;93m Press 'Enter' key to continue\033[0m")
                clear_console()
                banner()
                choices()
            if search_type_choice == "3":
                print("\n\033[1;92m Program terminated by user...\033[0m\n")
                exit(0)
    except KeyboardInterrupt:
        print("\n\033[1;92m Exiting...\033[0m\n")
        exit(0)
    except Exception as e:
        print(f' Error: {e}')


# Function to generate the rota
def generate_rota(persons, start_date, last_person, end_date):
    uk_holidays = holidays.UK()
    rota = []
    current_date = start_date
    person_index = persons.index(last_person) + 1

    while current_date <= end_date:
        day = current_date.strftime('%A')
        public_holiday = 'PH' if current_date in uk_holidays else ''
        on_call = persons[person_index % len(persons)]

        # Ensure the same person works Friday to Sunday
        if day == 'Friday':
            for _ in range(3):
                if current_date > end_date:
                    break
                day = current_date.strftime('%A')
                public_holiday = 'PH' if current_date in uk_holidays else ''
                rota.append([current_date.strftime('%d/%m/%Y'), day, on_call, public_holiday])
                current_date += timedelta(days=1)
            person_index += 1
        else:
            rota.append([current_date.strftime('%d/%m/%Y'), day, on_call, public_holiday])
            current_date += timedelta(days=1)
            person_index += 1

    return rota


def rotagen():
    try:
        persons = input("\033[1;97m Specify people to be on the rota. "
                        "Separate multiple with commas:\033[0m\n").split(', ')
        last_person = input("\033[1;97m\n Who worked the last day on the previous rota?\n"
                            " Value to be same format as the people to be on the rota:\033[0m\n")
        while True:
            start_date_str = input("\033[1;97m\n Date the previous rota ends (dd/mm/yyyy):\033[0m\n")
            try:
                start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
                print("\033[1;92m Start date:", start_date.strftime("%d/%m/%Y"), "\033[0m")
                break
            except ValueError:
                print("\033[1;91m Invalid date format. Please enter the date in dd/mm/yyyy format.\033[0m")

        while True:
            end_date_str = input("\033[1;97m\n When do you want the new rota to end? (dd/mm/yyyy):\033[0m\n")
            try:
                end_date = datetime.strptime(end_date_str, "%d/%m/%Y")
                print("\033[1;92m End date:", end_date.strftime("%d/%m/%Y"), "\033[0m")
                break
            except ValueError:
                print("\033[1;91m Invalid date format. Please enter the date in dd/mm/yyyy format.\033[0m")

    except Exception as input_error:
        print(f' Error: {input_error}')
    except KeyboardInterrupt:
        print(f'\n\n\033[1;91m User initiated exit.\n Exiting...\n\033[0m ')
        exit(0)

    start_date = datetime.strptime(start_date_str, '%d/%m/%Y') + timedelta(days=1)
    end_date = datetime.strptime(end_date_str, '%d/%m/%Y')

    rota = generate_rota(persons, start_date, last_person, end_date)

    start_date_filename = start_date_str.replace("/", "-")
    end_date_filename = end_date_str.replace("/", "-")

    df = pd.DataFrame(rota, columns=['Date', 'Day', 'On Call', 'Public Holiday'])
    folder_path = f'New_Rotas/'
    filename = f'{folder_path}New_Rota_{start_date_filename}-{end_date_filename}.xlsx'

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    df.to_excel(filename, index=False)
    print(f"\n\033[1;92m New rota has been saved:\n\033[0m"
          f"\033[1;97m {folder_path}New_Rota_{start_date_filename}-{end_date_filename}.xlsx\n\033[0m\033[0m")
    choices()


# Main function
def main():
    clear_console()
    banner()
    choices()


if __name__ == "__main__":
    main()
