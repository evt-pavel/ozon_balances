from send_email import send_email
from quantity import quantity, articles, value_comparison, today, yesterday


def main():
    articles()
    quantity()
    value_comparison()
    print(send_email(f'comparison_{today}.xlsx', f'ozon_{yesterday}.xlsx'))


if __name__ == '__main__':
    main()