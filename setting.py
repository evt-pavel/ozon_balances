import os
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

sender = os.environ.get('SENDER')
password = os.environ.get('PASSWORD')
recipient = os.environ.get('RECIPIENT')
