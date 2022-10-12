import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, date

def setup_custom_logger(name):
    today = date.today()
    d3 = today.strftime("%Y-%m-%d")
    filename = './Logs/'+str(d3)+" log.txt"

    logging.basicConfig(handlers=[RotatingFileHandler(filename, maxBytes=10485760, backupCount=5)],
                        level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s ')
    logger = logging.getLogger()
    return logger

def setup_custom_logger_file(name,RootFolder,OrderDate,ClientCode):

    #converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    print('Here')
    filename = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/80-logs/logs.txt"
    
    logging.basicConfig(handlers=[RotatingFileHandler(filename, maxBytes=10485760, backupCount=5)],
                        level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s ')
    file_logger = logging.getLogger()
    return file_logger