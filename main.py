from send_email import send_email
from quantity import quantity, articles, type_object
from datetime import date

def main():
    articles()
    quantity()
    print(send_email(f'ozon_{date.today().strftime("%d-%b-%Y")}.xlsx'))


if __name__ == '__main__':
    main()