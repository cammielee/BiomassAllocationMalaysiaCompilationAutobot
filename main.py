import datetime
import logging
from logging.handlers import SMTPHandler
import os

from utils.excel_autobot import excel_run
from utils.config import directory
from utils.checkreport_month import wantedFileName
logger = logging.getLogger(__name__)


dirr = directory()
rawdatafilename, report_monthyear, reportmonth_datetime = wantedFileName() # Choose latest month

now = datetime.datetime.now()
today = now.today()
today = today.strftime("%d-%m-%Y") # dd-mm-yyyy

def logging_func(sender_email=None, sender_password=None, recipients=[], send_email=False):
    """Logging setting, if do not put emails, then dont send error log"""
    # # ------ Based on current date-------


    # ------------- directory ------------
    master_dir = dirr.path
    log_dir = os.path.join(master_dir, 'Log')

    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    logger = logging.getLogger(__name__)
    # ------------ Formaters & Handlers ------------
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    filehandler = logging.FileHandler(os.path.join(
        log_dir, f"autobot_log_{today}.log"))
    filehandler.setFormatter(formatter)
    streamhandler = logging.StreamHandler()
    streamhandler.setFormatter(formatter)
    logging.basicConfig(level=logging.INFO, handlers=[filehandler, streamhandler])

    # --------- for error handlers  ---------------
    error_handler = SMTPHandler(mailhost=("smtp.office365.com", 587),
                            fromaddr=sender_email,
                            toaddrs=recipients,
                            subject=u"[Biomass Allocation Capacity Preprocessing Autobot] autobot_error_log",
                            credentials=(sender_email, sender_password),
                            secure=())
    error_handler.setLevel(logging.ERROR)
    # Sending email if error
    if sender_email != None and sender_password != None and send_email != False:
        logger.addHandler(error_handler)
    return logger

#  === For sending error logs =====
senderEmail = 'tmobot@topglove.com.my' 
senderPass = 'Passw0rd@123'
recipients = ['leesc@topglove.com.my'] # Error log recipients

if __name__ == "__main__":
    # try:
        logger = logging_func(senderEmail, senderPass, recipients, True)
        logger.info("Autobot starts...")
        logger.info(f"Report={report_monthyear}, today={today}")

        excel_run()
  
        logger.info("Done running Autobot")
    # except:
    #     logger.exception('Autobot failed to run, error as follows: \n\n')