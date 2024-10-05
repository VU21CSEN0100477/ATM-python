import getpass
import openpyxl
import random

def read_pins_and_balances_from_txt():
    try:
        with open("pin.txt", "r") as file:
            pin_data = [line.strip().split(',') for line in file.readlines()]
            pins = [int(line[0]) for line in pin_data]
            balances = {int(line[0]): int(line[1]) for line in pin_data}
            answers = {int(line[0]): line[2:] for line in pin_data}
        return pins, balances, answers
    except FileNotFoundError:
        print("PIN file not found.")
        exit()

def write_pins_and_balances_to_file(pins, balances, answers):
    try:
        with open("pin.txt", "w") as file:
            for pin in pins:
                balance = balances.get(pin)
                if balance is not None:
                    answer = answers.get(pin, ["", "", ""])
                    file.write(f"{pin},{balance},{','.join(answer)}\n")
                else:
                    print(f"Balance not found for PIN {pin}.")
    except FileNotFoundError:
        print("PIN file not found.")
        exit()

def append_pins_to_file(pin, balance, transaction_history):
    try:
        wb = openpyxl.load_workbook("database.xlsx")
        sheet = wb.active
        transaction_history_str = ', '.join(map(str, transaction_history))
        sheet.append([pin, balance, transaction_history_str])
        wb.save("database.xlsx")
        wb.close()
    except FileNotFoundError:
        print("Database file not found.")
        exit()

def print_mini_statement():
    try:
        with open("transactions.txt", "r") as file:
            lines = file.readlines()
            if len(lines) > 5:
                start_index = len(lines) - 5
            else:
                start_index = 0
            print("\nMini Statement:")
            for line in lines[start_index:]:
                print(line.strip())
    except FileNotFoundError:
        print("\nNo transactions found.")

while True:
    attempts = 3
    pin_list, balance_list, answer_list = read_pins_and_balances_from_txt()

    while attempts > 0:
        id_input = getpass.getpass(f"Please enter your 4-digit PIN (Attempts left: {attempts}): ")
        if id_input.isdigit() and 999 < int(id_input) < 10000 and int(id_input) in pin_list:
            print("PIN entered successfully!\n")
            print("Welcome! Thank you for choosing our ATM services.\n")
            print("Please choose from the below listed options: \n")
            print("0. Check Balance   1. Transaction History \n2. Withdraw        3. Deposit  \n4. Transfer        5. Reset PIN \n6. Mini Statement   7. Quit")
            break
        else:
            print("Your PIN is invalid")
            attempts -= 1
            if attempts == 0:
                print("ACCOUNT HAS BEEN LOCKED")
                exit()

    account_numbers = [123456789012, 234567890123, 345678901234, 456789012345, 67890123456, 678901234567,
                       789012345678, 890123456789, 901234567890]

    while True:
        balance = balance_list.get(int(id_input))
        if balance is not None:
            break
        else:
            print("Balance not found for this PIN. Please contact customer support.")
            exit()

    opt = int(input("\nEnter your choice~'0'/'1'/'2'/'3'/'4'/'5'/'6'/'7': "))
    pin_index = pin_list.index(int(id_input))

    def switch_case(value, pin_index, balance):
        if value == 0:
            print("Your current balance is: ", balance)
        elif value == 1:
            file = open("transactions.txt", "r")
            data = file.readlines()
            for line in data:
                print(line.strip())
            file.close()
        elif value == 2:
            try:
                withdraw_amount = int(input("Enter the amount you want to withdraw: "))
                if withdraw_amount > 0 and balance >= withdraw_amount:
                    new_balance = balance - withdraw_amount
                    balance_list[int(id_input)] = new_balance
                    write_pins_and_balances_to_file(pin_list, balance_list, answer_list)
                    append_pins_to_file(int(id_input), new_balance, ["Withdrawal: " + str(withdraw_amount)])
                    print("Withdrawal successful!")
                else:
                    print("Invalid amount or insufficient balance.")
            except ValueError:
                print("Invalid amount.")
        elif value == 3:
            try:
                deposit_amount = int(input("Enter the amount you want to deposit: "))
                if deposit_amount > 0:
                    new_balance = balance + deposit_amount
                    balance_list[int(id_input)] = new_balance
                    write_pins_and_balances_to_file(pin_list, balance_list, answer_list)
                    append_pins_to_file(int(id_input), new_balance, ["Deposit: " + str(deposit_amount)])
                    print("Deposit successful!")
                else:
                    print("Invalid amount for deposit.")
            except ValueError:
                print("Invalid amount.")
        elif value == 4:
            try:
                transfer_amount = int(input("Enter the amount you want to transfer: "))
                if transfer_amount > 0 and balance >= transfer_amount:
                    transfer_account = int(input("Enter the account number you want to transfer to: "))
                    if transfer_account in account_numbers:
                        new_balance_sender = balance - transfer_amount
                        new_balance_receiver = balance_list.get(transfer_account, 0) + transfer_amount
                        balance_list[int(id_input)] = new_balance_sender
                        balance_list[transfer_account] = new_balance_receiver
                        write_pins_and_balances_to_file(pin_list, balance_list, answer_list)
                        append_pins_to_file(int(id_input), new_balance_sender, ["Transfer to " + str(transfer_account) + ": " + str(transfer_amount)])
                        append_pins_to_file(transfer_account, new_balance_receiver, ["Transfer from " + str(id_input) + ": " + str(transfer_amount)])
                        print("Transfer successful!")
                    else:
                        print("Invalid account number!")
                else:
                    print("Invalid amount or insufficient balance.")
            except ValueError:
                print("Invalid amount.")
        elif value == 5:
            security_questions = [
                "What is your friend's name?",
                "What city were you born in?",
                "What is your favorite color?"
            ]
            random_question = random.choice(security_questions)
            print("Security Question:", random_question)
            answer = input("Enter your answer: ")
            if answer in answer_list[int(id_input)]:
                new_pin = getpass.getpass("Enter your new 4-digit PIN: ")
                if new_pin.isdigit() and 999 < int(new_pin) < 10000:
                    old_pin = pin_list[pin_index]
                    pin_list[pin_index] = int(new_pin)
                    answer_list[int(new_pin)] = answer_list.pop(old_pin)
                    balance_list[int(new_pin)] = balance_list.pop(old_pin)
                    write_pins_and_balances_to_file(pin_list, balance_list, answer_list)
                    print("PIN reset successful!")
                else:
                    print("Invalid PIN format. PIN should be a 4-digit number.")
            else:
                print("Incorrect answer.")
        elif value == 6:
            print_mini_statement()
        elif value == 7:
            print("Thank you for using our services!")
            exit()
        else:
            print("Invalid option! Please try again.")

    switch_case(opt, pin_index, balance)