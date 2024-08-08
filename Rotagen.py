import pandas as pd
from datetime import datetime, timedelta
import holidays


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


# Main function
def main():
    persons = input("Enter the person parameters separated by commas: ").split(', ')
    start_date_str = input("Enter the date the previous rota ends (dd/mm/yyyy): ")
    last_person = input("Enter the name parameter that was working on the last day of the previous rota: ")
    end_date_str = input("Enter the date to run the new rota to (dd/mm/yyyy): ")

    start_date = datetime.strptime(start_date_str, '%d/%m/%Y') + timedelta(days=1)
    end_date = datetime.strptime(end_date_str, '%d/%m/%Y')

    rota = generate_rota(persons, start_date, last_person, end_date)

    df = pd.DataFrame(rota, columns=['Date', 'Day', 'On Call', 'Public Holiday'])
    df.to_excel('rota.xlsx', index=False)
    print("Rota has been saved to 'rota.xlsx'")


if __name__ == "__main__":
    main()
